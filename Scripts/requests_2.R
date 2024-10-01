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


###############GXP NE
gxp_ne <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc < 14 & 
                   (!is.na(raw_data_baseline_arm_1$tbr_genex_result_1) &
                      (is.na(raw_data_baseline_arm_1$tbr_smear_res_1))))

gxp_ne <- gxp_ne[c('record_id', 'tbip_sc_ini_days_calc', 'tbr_genex_result_1', 'tbr_smear_res_1')]


########Smear NE
smear_ne <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc < 14 & 
                     (!is.na(raw_data_baseline_arm_1$tbr_smear_res_1) & 
                        (is.na(raw_data_baseline_arm_1$tbr_genex_result_1))))

smear_ne <- smear_ne[c('record_id', 'tbip_sc_ini_days_calc', 'tbr_smear_res_1', 'tbr_genex_result_1')]


#########smear and gxp
smear_gxp <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_date < 14 & 
                      (!is.na(raw_data_baseline_arm_1$tbr_smear_res_1) &
                         (!is.na(raw_data_baseline_arm_1$tbr_genex_result_1))))

smear_gxp <- smear_gxp[c('record_id', 'tbip_sc_ini_date', 'tbr_smear_res_1', 'tbr_genex_result_1')]


##
lpa <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc < 14 &
                (!is.na(raw_data_baseline_arm_1$tbr_lpa_done_res_1) &
                   (is.na(raw_data_baseline_arm_1$tbr_smear_res_1) & (is.na(raw_data_baseline_arm_1$tbr_genex_result_1)))))

lpa <- lpa[c('record_id', 'tbr_lpa_done_res_1', 'tbr_smear_res_1', 'tbr_genex_result_1')]



##########Enrolled no extraction
enrol_no_ex <- subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhc_sc_consent_provided=='Yes' &
                        (raw_data_hhci_info_arm_1$hhc_pc_days_since_referral > 30))

enrol_no_ex<- enrol_no_ex[c('record_id', 'hhc_sc_consent_provided', 'hhc_pc_days_since_referral', 'hhc_pc_intro', 'hhc_pc_attempt_outcome_1', 'hhc_pc_attempt_outcome_2', 'hhc_pc_attempt_outcome_3', 'hhc_pc_attempt_outcome_4', 'hhc_pc_attempt_outcome_5', 'hhc_pc_been_facility', 'hhc_pt_intro')]

write.table(enrol_no_ex, "Data/extraction data.csv", sep = ",", row.names = FALSE)


#
no_fu2_data <- subset(raw_data_follow_up_2_arm_1, (is.na(raw_data_follow_up_2_arm_1$tbr_exp_sputum_collection_date) |
                                                     is.na(raw_data_follow_up_2_arm_1$tbr_sputum_collected) |
                                                     is.na(raw_data_follow_up_2_arm_1$tbr_act_sputum_collection_date)|
                                                     is.na(raw_data_follow_up_2_arm_1$tbr_results_available)))

no_fu2_data <- no_fu2_data[c('record_id', 'tbr_exp_sputum_collection_date', 'tbr_sputum_collected', 'tbr_act_sputum_collection_date', 'tbr_results_available')]


v <- subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhc_sc_clinic_visit=='Yes')


#Presented No clinic extraction
no_extraction <- subset(raw_data_hhci_info_arm_1, !is.na(raw_data_hhci_info_arm_1$hhc_pc_present_referral))

no_extraction <- no_extraction[c('record_id', 'redcap_repeat_instance', 'hhc_pc_facility', 'hhc_pc_present_referral', 'hhc_pt_intro', 'hhc_pt_return_clinic')]

write.table(no_extraction, "Data/No extraction.csv", sep = ",", row.names = FALSE)


##########
hhc_not_visited <- subset(raw_data_hhci_info_arm_1, is.na(raw_data_hhci_info_arm_1$hhc_sc_intro) &
                            (is.na(raw_data_hhci_info_arm_1$hhc_sn_needed)))

hhc_not_visited <- hhc_not_visited[c('record_id', 'hhc_sc_intro', 'hhc_sn_needed')]

write.table(hhc_not_visited, "Data/hhc_not_visited.csv", sep = ",", row.names = FALSE)


##################################################################################################
aaron_request <- full_dataset_master[c('record_id', 'ra_instruct', 'house_number', 'attempt', 'ra_instruct_1stattempt', 'google_map_pic', 'ra_instruct_2nd_3rdattempt',
                                       'ins_13', 'labelling_hh_on_google_maps_complete', 'area', 'study_community', 'square_number', 'sub_square_letter', 'confirm_infrontofgate_door', 'instruct_gps', 'gps_coordinates', 'gps_coordinates_2', 'hhmember_present', 'intro_script1', 'hh_not_present_confirm', 
                                       'hoh_at_home', 'intro_script_2', 'interest_status', 'thanks', 'screening_for_eligibility', 'name', 'surname', 'how_old_are_you', 'dob', 'calculated_age', 'participant_gender', 'lang_fluent', 'community', 'ra_the_participant_is_elig', 'ra_the_person_is_not_eligi', 'note_to_ra_you_will_now_st', 'preferred_language_to_be_u',
                                       'icf_e', 'did_the_person_consent_to', 'participant_can_sign', 'i_name_surname_hereby_prov', 'signature_x', 'witness', 'witness_name', 'i_name_surname_hereby_prov_2', 'signature_x_2', 'consent_date', 'pin_setup1', 'pin_setup2', 'pin_setup3', 'pin_setup4', 'pin_setup5', 'refconfirm', 'consented_the_person_has_a', 'pin',
                                       'pin_display', 'consent_qc_1', 'consent_qc_2', 'consent_qc_3', 'contact_no', 'contact_no_2', 'contact_owner_other', 'timestamp_end_of_consent', 'screen_save', 'save_instruct_2', 'screening_and_consenting_complete', 'oversampled', 'questionnaires_complete')]



aaron_request <- full_dataset_master[-c(100)]


#####################################################################################################
sec_attempt <- subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhc_sc_attempt_1_present=='No' &
                        (is.na(raw_data_hhci_info_arm_1$hhc_sc_attempt_2_date)))


sec_attempt <- sec_attempt[c('record_id', 'redcap_repeat_instance','hhc_sc_attempt_1_present', 'hhc_sc_attempt_2_date')]


third_attempt <- subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhc_sc_attempt_2_present=='No' &
                          (is.na(raw_data_hhci_info_arm_1$hhc_sc_attempt_3_date)))


third_attempt <- third_attempt[c('record_id', 'redcap_repeat_instance','hhc_sc_attempt_2_present', 'hhc_sc_attempt_3_date')]



write.table(sec_attempt, "Data/no second attempt.csv", sep = ",", row.names = FALSE)

write.table(third_attempt, "Data/no third attempt.csv", sep = ",", row.names = FALSE)




###############################################################################################################
hhc_enroled <- subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhc_sc_consent_provided=='Yes')


no_clinic_enrol <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes')

no_clinic_enrol <- no_clinic_enrol[c('record_id', 'tbip_sc_q5')]

unique(no_clinic_enrol$tbip_sc_q5)



##############################################################
baseline <- raw_data_baseline_arm_1[c('record_id', 'tbip_sc_q5', 'tbip_sc_ini_days_calc', 'tbip_sc_q14', 'tbip_sc_consent_part')]

sc <- raw_data_hhci_visit_info_arm_1[c('record_id', 'hhc_css_area','hhc_css_study_community')]

study_community <- left_join(baseline, sc, by='record_id')

write.table(study_community, "Data/study community.csv", sep = ",", row.names = FALSE)



# PRESENTED HHC
presented_hhc <- subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhc_sc_consent_provided=='Yes' &
                          (raw_data_hhci_info_arm_1$hhc_pc_been_facility=='Yes, I remember the date' | raw_data_hhci_info_arm_1$hhc_pc_been_facility=="Yes, I don't remember the date" | 
                             raw_data_hhci_info_arm_1$hhc_pt_return_clinic=='Yes'))

presented_hhc <- presented_hhc[c('record_id', 'redcap_repeat_instance', 'hhc_pc_been_facility', 'hhc_pc_been_facility', 'hhc_pt_return_clinic')]

write.table(presented_hhc, "Data/presented hhc.csv", sep = ",", row.names = FALSE)



# NOT PRESENTED HHC
not_presented <- subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhc_sc_consent_provided=='Yes' &
                          (raw_data_hhci_info_arm_1$hhc_pc_been_facility=='No') & (raw_data_hhci_info_arm_1$hhc_pt_return_clinic=='No'))

not_presented <- not_presented[c('record_id', 'redcap_repeat_instance', 'hhc_pc_been_facility', 'hhc_pt_return_clinic')]

write.table(not_presented, "Data/not present hhc.csv", sep = ",", row.names = FALSE)


#############################################################################################################
less_than_30 <- subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhc_sc_consent_provided=='Yes')

less_than_30 <- less_than_30[c('record_id', 'redcap_repeat_instance', 'hhc_sc_consent_provided', 'hhc_sc_date_cons', 'hhc_pc_presentation_date', 'hhc_pt_return_date')]

write.table(less_than_30, "Data/less_than_30.csv", sep = ",", row.names = FALSE)


##############################################################################################################
hhc_hiv <- subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhc_sc_consent_provided=='Yes')

hhc_hiv <- hhc_hiv[c('record_id', 'hhc_sc_consent_provided','hhc_q1_hiv_th_q1', 'hhc_q1_hiv_status')]

write.table(hhc_hiv, "Data/hhc hiv.csv", sep = ",", row.names = FALSE)



###########################################################################################################################
adherence <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' &
                      (raw_data_baseline_arm_1$tbip_sc_ini_days_calc < 14))

adherence <- adherence[c('record_id', 'treat_out_visit_1_calc', 'treat_out_visit_1_calc_2', 'treat_out_visit_1_calc_3', 'treat_out_visit_1_calc_4', 'treat_out_visit_1_calc_5', 'treat_out_visit_1_calc_6', 'treat_out_visit_1_calc_7', 'treat_out_visit_1_calc_8', 'treat_out_visit_1_calc_9', 'treat_out_visit_1_calc_10', 'treat_out_visit_1_calc_11')]

result <- apply(adherence,2, function(x) x[x > 30])



#################################################################################################################################
no_ques <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & (raw_data_baseline_arm_1$index_questionnaire_3_complete=='Unverified' |
                    raw_data_baseline_arm_1$index_questionnaire_3_complete=='Incomplete'))

no_ques <- no_ques[c('record_id', 'tbip_sc_consent_part', 'index_questionnaire_3_complete', 'sn_add_notes')]

write.table(no_ques, "Data/incomplete questionnaires.csv", sep = ",", row.names = FALSE)
