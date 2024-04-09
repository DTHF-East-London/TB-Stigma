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



#Follow Up Dataset
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




#############################################################
df1 <- follow_up_data

df2 <- raw_data_follow_up_2_arm_1[c('record_id', 'tbr_exp_sputum_collection_date', 'tbr_sputum_collected', 'index_follow_up_questionnaire_3_complete')]

follow_up_2_data <- left_join(df1, df2, by='record_id')

write.table(follow_up_2_data, "Data/follow up 2 data.csv", sep = ",", row.names = FALSE)



#####################################################
refused <- subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhc_sc_consent_provided=='No' |
                    raw_data_hhci_info_arm_1$hhc_sc_competent=='No')

refused <- refused [c('record_id', 'hhc_sc_consent_provided', 'hhc_sc_competent')]

write.table(refused, "Data/refused.csv", sep = ",", row.names = FALSE)



perc <- subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhc_sc_consent_provided=='Yes')

perc <- perc[c('record_id', 'hhc_sc_consent_provided', 'hhc_sc_age_calc')]

clean_perc <- na.omit(perc)


percentiles <- c(0.25, 0.5, 0.75)

percentile_values <- quantile(perc$hhc_sc_age_calc, probs = percentiles, na.rm = TRUE)

print(percentile_values)


#################################################################
pre_tutt_enr <- subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$record_id < 1105 &
                         (raw_data_hhci_info_arm_1$hhc_sc_consent_provided=='Yes'))

pre_tutt_enr <- pre_tutt_enr[c('record_id', 'hhc_sc_consent_provided')]


pre_tutt_hhc <- raw_data_hhci_info_arm_1 %>%
  dplyr::filter(hhc_sc_consent_provided == 'Yes') %>%
  dplyr::filter(record_id < 1105) %>%
  dplyr::distinct(record_id)




##################################################################

hh_listed <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' | 
                      raw_data_baseline_arm_1$tbip_sc_cgiver_permission=='Yes')

###################################################################

hh_listed_ni <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' | raw_data_baseline_arm_1$tbip_sc_cgiver_permission == 'Yes' & 
                         (raw_data_baseline_arm_1$tbip_sc_ini_days_calc < 14))

hh_listed_ni <- hh_listed_ni[c('record_id', 'tbip_sc_consent_part', 'tbip_sc_cgiver_permission', 'tbip_sc_ini_days_calc')]


#Aaron's request
aaron_request1 <- subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhc_sc_consent_provided=="Yes" &
                           (raw_data_hhci_info_arm_1$hhc_sc_night_sweat=="Yes" |
                              raw_data_hhci_info_arm_1$hhc_sc_coughing=="Yes, lasted more than 2 weeks." |
                              raw_data_hhci_info_arm_1$hhc_sc_coughing=="Yes, lasted less than 2 weeks." |
                              raw_data_hhci_info_arm_1$hhc_sc_fever=="Yes"))

aaron_request1 <- aaron_request1[c('record_id', 'hhc_sc_date_cons','hhc_sc_age', 'hhc_sc_weight_loss', 'hhc_sc_night_sweat', 'hhc_sc_coughing', 'hhc_sc_coughing_blood', 'hhc_sc_fever', 'hhcd_gender', 'hhcd_s1_q1', 'hhcd_s1_q1_other', 'hhcd_place_of_birth_rsa', 'hhcd_place_of_birth_nrsa', 'hhcd_place_of_birth_nrsa_oth', 'hhcd_s1_q5', 'hhcd_living_arrangement', 
                                   'hhcd_emp_q1', 'hhcd_emp_q1_oth', 'hhcd_s1_q14', 'hhcd_ac_prison', 'hhcd_ac_mines', 'hhc_q1_language_prefer', 'hhc_q1_ch_q1', 'hhc_q1_s3_complete_tb_treatment', 'hhc_q1_tb_m', 'hhc_q1_s3_they_complete_tb_treatment', 'hhc_q1_s2_q3', 'hhc_q1_tbk_q1___1', 'hhc_q1_tbk_q1___2', 'hhc_q1_tbk_q1___3', 'hhc_q1_tbk_q1___4',
                                   'hhc_q1_tbk_q1___5', 'hhc_q1_tbk_q1___6', 'hhc_q1_tbk_q1___7', 'hhc_q1_tbk_q1___8', 'hhc_q1_tbk_q1___99', 'hhc_q1_tbk_q1_other', 'hhc_q1_tbk_q2___1', 'hhc_q1_tbk_q2___2', 'hhc_q1_tbk_q2___3', 'hhc_q1_tbk_q2___4', 'hhc_q1_tbk_q2___5', 'hhc_q1_tbk_q2___6', 'hhc_q1_tbk_q2___7', 'hhc_q1_tbk_q2_other', 'hhc_q1_tbk_q3___1',
                                   'hhc_q1_tbk_q3___2', 'hhc_q1_tbk_q3___3', 'hhc_q1_tbk_q3___4', 'hhc_q1_tbk_q3___5', 'hhc_q1_tbk_q3___6', 'hhc_q1_tbk_q3___7', 'hhc_q1_tbk_q3___8', 'hhc_q1_tbk_q3___9', 'hhc_q1_tbk_q3___10', 'hhc_q1_tbk_q3___11', 'hhc_q1_tbk_q3_other', 'hhc_q1_tbk_q4', 'hhc_q1_tbk_q5', 'hhc_q1_highest_education', 'hhc_q1_highest_primary',
                                   'hhc_q2_mos1', 'hhc_q2_mos2', 'hhc_q2_mos3', 'hhc_q2_mos4', 'hhc_q2_mos5', 'hhc_q2_mos6', 'hhc_q2_mos7', 'hhc_q2_mos8', 'hhc_q2_mos9', 'hhc_q2_mos10', 'hhc_q2_mos11', 'hhc_q2_mos12', 'hhc_q2_mos13', 'hhc_q2_mos14', 'hhc_q2_mos15', 'hhc_q2_mos16', 'hhc_q2_mos17', 'hhc_q2_mos18', 'hhc_q2_mos19', 'hhc_q2_sst1', 'hhc_q2_sst2',
                                   'hhc_q2_sst3', 'hhc_q2_sst4', 'hhc_q2_sst5', 'hhc_q2_sst6', 'hhc_q2_sst7', 'hhc_q2_sst8', 'hhc_q2_sst9', 'hhc_q2_sst10', 'hhc_q2_sst11', 'hhc_q2_sst12', 'hhc_q2_sst13', 'hhc_q2_sst14', 'hhc_q2_sst15', 'hhc_q2_sst16', 'hhc_q2_sst17', 'hhc_q2_sst18', 'hhc_q2_sst19', 'hhc_q2_sst20', 'hhc_q2_sst21', 'hhc_q2_sst22', 'hhc_q2_sst23',
                                   'hhc_q2_sst24', 'hhc_q2_sst25', 'hhc_q2_sst26', 'hhc_q2_sst27', 'hhc_q2_sst28', 'hhc_q2_fs_c1', 'hhc_q2_fs_c2', 'hhc_q2_fs_c3', 'hhc_q2_fs_c4', 'hhc_q2_fs_c4a', 'hhc_q2_fs_c5', 'hhc_q2_fs_c6', 'hhc_q2_fs_c7', 'hhc_q2_fs_c8', 'hhc_q2_fs_c8a', 'hhc_q2_fs_c_child_age', 'hhc_q2_fs_c9', 'hhc_q2_fs_c10', 'hhc_q2_fs_c11', 'hhc_q2_fs_c12', 'hhc_q2_fs_c13', 'hhc_q2_fs_c13a', 'hhc_q2_fs_c14', 'hhc_q2_fs_c15', 'hhc_q2_psp_d1', 'hhc_q2_psp_d2', 'hhc_q2_psp_d3', 'hhc_q2_psp_d4', 'hhc_q2_psp_d5', 'hhc_q2_psp_d6', 'hhc_q2_psp_d7', 'hhc_q2_psp_d8', 'hhc_q2_psp_d9', 'hhc_q2_psp_d10', 'hhc_q2_psp_d11', 'hhc_q2_psp_d12', 'hhc_q2_psp_d13', 'hhc_q2_psp_d14', 'hhc_q2_psp_d15', 'hhc_q2_psp_d16', 'hhc_q2_psp_d17', 'hhc_q2_psp_d18', 'hhc_q2_psp_d19', 'hhc_q2_psp_d20', 'hhc_q2_psp_d21', 'hhc_q2_psp_d22', 'hhc_q2_psp_d23', 'hhc_q2_ips_q1', 'hhc_q2_ips_q2', 'hhc_q2_ips_q3', 'hhc_q2_ips_q4', 'hhc_q2_eps_q1', 'hhc_q2_eps_q2', 'hhc_q2_eps_q3', 'hhc_q2_eps_q4', 'hhc_q3_cart1', 'hhc_q3_cart2', 'hhc_q3_cart3', 'hhc_q3_cart4', 'hhc_q3_cart5', 'hhc_q3_cart6', 'hhc_q3_cart7', 'hhc_q3_cart8', 'hhc_q3_cart9', 'hhc_q3_cart10', 'hhc_q3_cart11', 'hhc_q3_cart12', 'hhc_q3_cart13', 'hhc_q3_cart29', 'hhc_q3_cart30', 'hhc_q3_cart31', 'hhc_q3_cart32', 'hhc_q3_cart33', 'hhc_q3_cart34', 'hhc_q3_cart35', 'hhc_q3_cart36', 'hhc_q3_cart37', 'hhc_q3_cart38', 'hhc_q3_cart39', 'hhc_q3_cart40', 'hhc_q3_cart41', 'hhc_q3_cart42', 'hhc_q3_cart43', 'hhc_q3_cart44', 'hhc_q3_hiv_know1', 'hhc_q3_hiv_know2', 'hhc_q3_hiv_know3', 'hhc_q3_hiv_know4', 'hhc_q3_hiv_know5', 'hhc_q3_hiv_know6', 'hhc_q3_hiv_know7', 'hhc_q3_hiv_know8', 'hhc_q3_hiv_know9', 'hhc_q3_hiv_know10', 'hhc_q3_hiv_know11', 'hhc_q3_hiv_know12', 'hhc_q3_hiv_know13', 'hhc_q3_hiv_know14', 'hhc_q3_hiv_know15', 'hhc_q3_hiv_know16', 'hhc_q3_hiv_know17', 'hhc_q3_hiv_know18', 'hhc_q3_hiv_stig_s5_q1', 'hhc_q3_hiv_stig_s5_q2', 'hhc_q3_hiv_stig_s5_q3', 'hhc_q3_hiv_stig_s5_q4', 'hhc_q3_hiv_stig_s5_q5', 'hhc_q3_hiv_stig_s5_q6', 'hhc_q3_hiv_stig_s5_q7', 'hhc_q3_hiv_stig_s5_q8', 'hhc_q3_hiv_stig_s5_q9', 'hhc_q3_hiv_stig_s5_q10', 'hhc_q3_hiv_stig_s5_q11', 'hhc_q3_hiv_stig_s5_q12', 'hhc_q3_hiv_stig_s5_q13', 'hhc_q3_hiv_stig_s5_q14', 'hhc_q3_hiv_stig_s5_q15', 'hhc_q3_tb_know1', 'hhc_q3_tb_know2', 'hhc_q3_tb_know3', 'hhc_q3_tb_know4', 'hhc_q3_tb_know5', 'hhc_q3_tb_know6', 'hhc_q3_tb_know7', 'hhc_q3_tb_know8', 'hhc_q3_tb_know9', 'hhc_q3_tb_know10', 'hhc_q3_tb_know11', 'hhc_q3_s6_q1', 'hhc_q3_s6_q2', 'hhc_q3_s6_q3', 'hhc_q3_s6_q4', 'hhc_q3_s6_q5', 'hhc_q3_s6_q6', 'hhc_q3_s6_q7', 'hhc_q3_s6_q8', 'hhc_q3_s6_q9', 'hhc_q3_s6_q10', 'hhc_q3_s6_q11', 'hhc_q3_s6_q12', 'hhc_q3_tbs_s6_q1', 'hhc_q3_tbs_s6_q2', 'hhc_q3_tbs_s6_q3', 'hhc_q3_tbs_s6_q4', 'hhc_q3_tbs_s6_q5', 'hhc_q3_tbs_s6_q6', 'hhc_q3_tbs_s6_q7', 'hhc_q3_tbs_s6_q8', 'hhc_q3_tbs_s6_q9', 'hhc_q3_tbs_s6_q10', 'hhc_pc_present_referral', 'hhc_pt_return_clinic')]


data1 <- aaron_request1

data2 <- raw_data_hhci_visit_info_arm_1[c('record_id', 'hhc_css_area', 'hhc_css_study_community')]

aar_data_req <- left_join(data1, data2, by='record_id')

write.table(aar_data_req, "Data/data request.csv", sep = ",", row.names = FALSE)


#
hhc_not_presented <- subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhc_sc_consent_provided=='Yes' &(
                            is.na(raw_data_hhci_info_arm_1$hhc_pc_been_facility) | 
                              is.na(raw_data_hhci_info_arm_1$hhc_pt_return_clinic)))

hhc_not_presented <- hhc_not_presented[c('record_id', 'hhcl_member_name', 'hhcl_member_surname', 'hhc_q1_tb_care_clinic', 'hhc_q1_ac_hiv_clinic_prefer_test', 'hhc_q1_ac_hiv_clinic_prefer_treat', 'hhc_pc_been_facility', 'hhc_pt_return_clinic')]


write.table(hhc_not_presented, "Data/hhc_not_presented.csv", sep = ",", row.names = FALSE)


#
ex_ne <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc >= 14 &
                  (raw_data_baseline_arm_1$tbip_sc_q8=='No'))

ex_ne <- ex_ne[c('record_id', 'tbip_sc_ini_days_calc', 'tbip_sc_q8')]



#####
df_hhs_mas <- dataset_master[c('record_id', 'attempt', 'consent_date', 'timestamp_end_of_consent', 'participant_gender', 'calculated_age')]

df_hhs_hhd <- dataset_hhd

hhc_l_data <- left_join(df_hhs_mas, df_hhs_hhd, by='record_id')

write.table(hhc_l_data, "Data/HH Data.csv", sep = ",", row.names = FALSE)
