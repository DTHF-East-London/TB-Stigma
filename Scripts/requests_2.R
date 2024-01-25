library(openxlsx)
library(tidyverse)
library(dplyr)
library(redcapAPI)
library(RMySQL)
library(summarytools)
library(readxl)
library(haven)
library(xlsx)
library(survival)
library(conflicted)
library(lubridate)

#source("Scripts/dataset_generator_1.R")

no_gxp <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc < 14 & 
                   (is.na(raw_data_baseline_arm_1$tbr_smear_res_1) & (!is.na(raw_data_baseline_arm_1$tbr_genex_result_1))))

no_gxp <- no_gxp[c('record_id', 'tbip_sc_q5', 'tbip_sc_part_name', 'tbip_sc_part_surname', 'tbip_sc_dob', 'tbip_sc_consent_part', 'tbip_sc_ini_date', 'tbr_smear_res_1','tbr_genex_result_1', 'tbr_gx_done_res_2')]

write.table(no_gxp,"Data/no gxp.csv", sep = ",", row.names = FALSE)


###
smear_ni_results <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc < 14 &
                             (!is.na(raw_data_baseline_arm_1$tbr_smear_res_1) & (is.na(raw_data_baseline_arm_1$tbr_gx_done_res_2))))

smear_ni_results <- smear_ni_results[c('record_id', 'tbr_smear_res_1', 'tbr_smear_res_2', 'tbr_gx_done_res_2')]

write.table(smear_ni_results, "Data/smear ni results.csv", sep = ",", row.names = FALSE)


####
returned <- subset(raw_data_follow_up_1_arm_1, !is.na(raw_data_follow_up_1_arm_1$tbr_smear_res_1) &
                     (raw_data_follow_up_1_arm_1$index_follow_up_questionnaire_3_complete=='Incomplete'))

returned <- returned[c('record_id', 'tbr_smear_res_1', 'tbr_smear_res_2', 'index_follow_up_questionnaire_3_complete')]

write.table(returned, "Data/fu questionnaire not completed.csv", sep = ",", row.names = FALSE)


#############
gxp_returned <- subset(raw_data_follow_up_1_arm_1, !is.na(raw_data_follow_up_1_arm_1$tbr_genex_result_1))

gxp_returned <- gxp_returned[c('record_id', 'tbr_genex_result_1', 'tbr_gx_done_res_2')]

write.table(gxp_returned, "Data/gxp returned.csv", sep = ",", row.names = FALSE)


######################
smear_returned <- subset(raw_data_follow_up_1_arm_1, raw_data_follow_up_1_arm_1$tbr_smear_1=='Yes')

smear_returned <- smear_returned[c('record_id', 'tbr_smear_1')]


new_df <- rbind(raw_data_baseline_arm_1,raw_data_baseline_arm_1)

df2 <- merge(x=raw_data_baseline_arm_1,y=raw_data_baseline_arm_1, 
             by="record_id", all.x=TRUE)

##############################################################################################################

may_enrolment <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q5=="John Dube Clinic" & 
                          (raw_data_baseline_arm_1$tbip_sc_ini_days_calc < 14) &
                        (as.Date(raw_data_baseline_arm_1$tbip_sc_date, format="%d-%m-%Y") < as.Date("2023-05-31")))

may_enrolment <- may_enrolment[c('record_id', 'tbip_sc_date', 'tbip_sc_q5')]


#June
June_enrolment <- subset(raw_data_baseline_arm_1, (raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes') & (raw_data_baseline_arm_1$tbip_sc_q5=='Ginsberg Clinic') &
                 (as.Date(raw_data_baseline_arm_1$tbip_sc_date, format="%d-%m-%Y") >= as.Date("2023-05-31"))&
                   (as.Date(raw_data_baseline_arm_1$tbip_sc_date, format="%d-%m-%Y") <= as.Date("2023-07-01")))

June_enrolment <- June_enrolment[c('record_id', 'tbip_sc_consent_part', 'tbip_sc_date')]


########
smear <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc < 14 &
                  (!is.na(raw_data_baseline_arm_1$tbr_smear_res_1)))

smear <- smear[c('record_id', 'tbr_smear_res_1', 'tbr_smear_res_2')]

write.table(smear, 'Data/Smear results.csv', sep = ",", row.names = FALSE)



###################
gxp_only <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc < 14 & 
                   (is.na(raw_data_baseline_arm_1$tbr_smear_res_1) & (!is.na(raw_data_baseline_arm_1$tbr_genex_result_1))))

gxp_only <- gxp_only[c('record_id', 'tbr_smear_res_1','tbr_genex_result_1', 'tbr_gx_done_res_2')]

write.table(no_gxp,"Data/no gxp.csv", sep = ",", row.names = FALSE)



#############################################################Follow Up Datatset#########################################
data_1 <- raw_data_baseline_arm_1[c('record_id', 'tbip_sc_date', 'tbip_sc_ini_date','tbip_sc_ini_days_calc','tbip_sc_q5', 'tbip_sc_consent_part', 'tbr_fu_present', 'tbr_smear_1', 'tbr_smear_res_1', 'tbr_smear_2', 'tbr_smear_res_2', 'tbr_genex_1', 'tbr_genex_result_1', 'tbr_genex_2', 'tbr_gx_done_res_2', 'treat_out_outcome')]

#data_1$days_since_initiation <- difftime(tbip_sc_ini_date, Sys.Date(), units = "days")

#data_1 <- relocate(data_1, days_since_initiation, .after = tbip_sc_ini_days_calc)

data_2 <- raw_data_follow_up_1_arm_1[c('record_id', 'tbr_sputum_collected', 'tbr_fu_present', 'tbr_smear_1', 'tbr_smear_res_1', 'tbr_smear_2', 'tbr_smear_res_2', 'tbr_genex_1', 'tbr_genex_result_1', 'tbr_genex_2', 'tbr_gx_done_res_2', 'index_follow_up_questionnaire_1_complete', 'index_follow_up_questionnaire_2_complete', 'index_follow_up_questionnaire_3_complete')]

follow_up_data <- left_join(data_1, data_2, by='record_id')

write.table(follow_up_data, "Data/follow up data.csv", sep = ",", row.names = FALSE)


####################################
converted <- subset(follow_up_data, !is.na(follow_up_data$tbr_fu_present.x) | !is.na(follow_up_data$tbr_fu_present.y))

converted <- converted[c('record_id', 'tbr_fu_present.x', 'tbr_fu_present.y')]

write.table(converted, "Data/Converted.csv", sep = ",", row.names = FALSE)



new_variable <- subset(follow_up_data, follow_up_data$tbr_fu_present.x=='Yes' | follow_up_data$tbr_fu_present.y=='Yes')

new_variable <- new_variable[c('record_id', 'tbip_sc_ini_days_calc','tbr_fu_present.x', 'tbr_fu_present.y')]

write.table(new_variable, "Data/new variable.csv", sep = ",", row.names = FALSE)



follow_up_converted <- subset(raw_data_follow_up_1_arm_1, raw_data_follow_up_1_arm_1$tbr_fu_present=='Yes')

follow_up_converted <- follow_up_converted[c('record_id', 'tbr_fu_present', 'tbr_smear_res_1', 'tbr_smear_res_2')]


hhc_refused <- subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhc_sc_consent_provided=='No' | raw_data_hhci_info_arm_1$hhc_sc_competent=='No')

hhc_refused <- hhc_refused[c('record_id', 'redcap_repeat_instance', 'hhc_sc_date', 'hhc_sc_consent_provided', 'hhc_sc_competent')]

write.table(hhc_refused, "Data/hhc refused.csv", sep = ",", row.names = FALSE)



###########################################################################
only_gxp <- subset(follow_up_data, follow_up_data$tbip_sc_ini_days_calc < 14 & !is.na(follow_up_data$tbr_genex_result_1.x) &
                     (is.na(follow_up_data$tbr_smear_res_1.y)))

only_gxp <- only_gxp[c('record_id', 'tbr_genex_result_1.x', 'tbr_smear_res_1.y')]

