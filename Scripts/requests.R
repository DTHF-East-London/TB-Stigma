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
library(table1)

#source("Scripts/dataset_generator_1.R")

baseline_ques <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$index_questionnaire_3_complete=='Incomplete')

baseline_ques <- baseline_ques[c('record_id', 'tbip_sc_ini_days_calc','tbip_sc_date','tbip_sc_q5','tbip_sc_consent_part', 'tbid_eligible', 'tbip_q1_language_prefer', 'tbip_q2_ss1', 'tbip_q3_cart1', 'index_questionnaire_3_complete')]


#follow-up
sputum <- subset(raw_data_follow_up_1_arm_1, raw_data_follow_up_1_arm_1$tbr_sputum_collected=='Yes')

sputum <- sputum[c('record_id', 'tbr_sputum_collected')]


#Majiza
liya <- subset(raw_data_follow_up_1_arm_1, raw_data_follow_up_1_arm_1$tbr_sputum_collected=='Yes')

liya <- liya[c('record_id', 'tbr_sputum_collected', 'tbr_results_available')]

write.table(liya, "Data/follow-up.csv", sep = ",", row.names = FALSE)


follow_up_returned <- subset(raw_data_hhci_info_arm_1, (!is.na(raw_data_hhci_info_arm_1$tbr_act_sputum_collection_date)))

follow_up_returned <- follow_up_returned[c('record_id', 'tbr_act_sputum_collection_date')]

positive <- subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$tbr_smear_res_1=='Postive 1+' | raw_data_hhci_info_arm_1$tbr_smear_res_1=='Postive 2+' | raw_data_hhci_info_arm_1$tbr_smear_res_1=='Postive 3+')

positive <- positive[c('record_id')]


#follow up
tb_results <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & (raw_data_baseline_arm_1$tbr_smear_res_1=='Negative'))

tb_results <- tb_results[c('record_id')]


tb_results_pos <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 &
                           (raw_data_baseline_arm_1$tbr_smear_res_1=='Positive 1+' | raw_data_baseline_arm_1$tbr_smear_1=='Positive 2+' | raw_data_baseline_arm_1$tbr_smear_res_1=='Positive 3+'))

tb_results_pos <- tb_results_pos[c('record_id')]

#missing results
missing_results <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes'
                          & (is.na(raw_data_baseline_arm_1$tbr_med_data_extract=='Proceed')))

missing_results <- missing_results[c('record_id', 'tbip_sc_ini_days_calc', 'tbip_sc_q5','tbip_sc_part_name', 'tbip_sc_part_surname', 'tbip_sc_dob','tbr_sputum_collected')]

write.table(missing_results, "Data/missing results.csv", sep = ",", row.names = FALSE)


#results
results <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc < 14 &
                    (raw_data_baseline_arm_1$tbr_smear_res_1=='Negative'))

results <- results[c('record_id', 'tbip_sc_ini_days_calc', 'tbr_smear_res_1')]

#results 2
results_2 <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc < 14 &
                      (!is.na(raw_data_baseline_arm_1$tbip_image_4)))

results_2 <- results_2[c('record_id', 'tbip_image_4', 'tbr_smear_res_1')]


#GeneXpert
results_gxp <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc < 14 &
                      (!is.na(raw_data_baseline_arm_1$tbip_image_5)))

results_gxp <- results_gxp[c('record_id', 'tbip_image_5', 'tbr_genex_result_1')]





#datasets
write.table(raw_data_baseline_arm_1, "Data/Baseline.csv", sep = ",", row.names = FALSE)

write.table(raw_data_baseline_ex_arm_1, "Data/Baseline_experienced.csv", sep = ",", row.names = FALSE)

write.table(raw_data_baseline_ni_arm_1, "Data/Baseline_ni.csv", sep = ",", row.names = FALSE)

write.table(raw_data_follow_up_1_arm_1, "Data/Follow-up.csv", sep = ",", row.names = FALSE)

write.table(raw_data_hhci_info_arm_1, "Data/hhci_info.csv", sep = ",", row.names = FALSE)

write.table(raw_data_hhci_visit_info_arm_1, "Data/HHCI visit info.csv", sep = ",", row.names = FALSE)

write.table(raw_data_follow_up_2_arm_1, "Data/follow_up_2.csv", sep = ",", row.names = FALSE)


#follow Up ques
fu_ques <- subset(raw_data_follow_up_1_arm_1, raw_data_follow_up_1_arm_1$tbr_smear_1=='Negative')

fu_ques <- fu_ques[c('record_id', 'index_document_upload_complete')]


#clinic extraction
clinic_ex <- subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhc_pt_collect_sputum=='Yes')

clinic_ex <- clinic_ex[c('record_id', 'hhc_pt_collect_sputum')]


#outcomes
outcomes <- subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$days_since_referral<= 30 &
                     (raw_data_hhci_info_arm_1$hhc_sc_consent_provided=='Yes'))

outcomes <- outcomes[c('record_id')]


#visit date
visit_date <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' &
                       (!is.na(raw_data_baseline_arm_1$tbip_sc_date_visit_hh1)))

visit_date <- visit_date[c('record_id', 'tbip_sc_ini_days_calc','tbip_sc_q5','tbip_sc_suburb','tbip_sc_street_add', 'tbip_sc_date_time_hh1', 'tbip_sc_date_visit_hh1')]

write.table(visit_date, "Data/visit date.csv", sep = ",", row.names = FALSE)


#visit confirmation
visit_date_2 <- subset(raw_data_hhci_visit_info_arm_1, raw_data_hhci_visit_info_arm_1$hhc_sch_call_attempt_con=='Proceed' &
                         (!is.na(raw_data_hhci_visit_info_arm_1$hhc_sch_hhi_date_visit)))

visit_date_2 <- visit_date_2[c('record', 'hhc_sch_hhi_date_visit', 'hhc_sch_visit_timeslot')]


#baseline gxp
baseline_gxp_neg <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' &
                         (raw_data_baseline_arm_1$tbr_genex_result_1=='Negative (Not Detected)'))

baseline_gxp_neg <- baseline_gxp_neg[c('record_id', 'tbip_sc_q5','tbr_genex_result_1')]


#baseline gxp pos
baseline_gxp_pos <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc < 14 &
                             (raw_data_baseline_arm_1$tbr_genex_result_1=='Positive (Micro bacterium Detected)' | raw_data_baseline_arm_1$tbr_smear_res_1=='Positive 1+' | raw_data_baseline_arm_1$tbr_smear_res_1=='Positive 2+' | raw_data_baseline_arm_1$tbr_smear_res_1=='Positive 3+'))

baseline_gxp_pos <- baseline_gxp_pos[c('record_id', 'tbip_sc_ini_days_calc', 'tbr_genex_result_1', 'tbr_smear_res_1')]


#follow up
follow_up_eligible <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc < 14 &
                               (raw_data_baseline_arm_1$tbr_smear_res_1=='Negative'))

follow_up_eligible <- follow_up_eligible[c('record_id', 'tbip_sc_ini_date', 'tbip_sc_q5', 'tbr_smear_res_1')]

follow_up_eligible$week_7 <- tbip_sc_ini_date + weeks(7)

follow_up_eligible <- relocate(follow_up_eligible, week_7, .after = tbr_smear_res_1 )


#questionnaire 3
questionnaire <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' &
                        (raw_data_baseline_arm_1$tbip_sc_ini_days_calc > 13))

questionnaire <- questionnaire[c('record_id', 'tbip_sc_ini_days_calc', 'tbip_f3_ql_25')]


#'tbip_f3_ql_q2', 'tbip_f3_ql_q3', 'tbip_f3_ql_q4', 'tbip_f3_ql_q5', 'tbip_f3_ql_q6', 'tbip_f3_ql_q7', 'tbip_f3_ql_q8', 'tbip_f3_ql_q9', 'tbip_f3_ql_10', 'tbip_f3_ql_11', 'tbip_f3_ql_12', 'tbip_f3_ql_13', 'tbip_f3_ql_14', 'tbip_f3_ql_q15', 'tbip_f3_ql_16', 'tbip_f3_ql_17', 'tbip_f3_ql_18', 'tbip_f3_ql_19', 'tbip_f3_ql_20', 'tbip_f3_ql_21', 'tbip_f3_ql_22', 'tbip_f3_ql_23', 'tbip_f3_ql_24', 'tbip_f3_ql_25', 'tbip_f3_ql_q26')]



#follow up
follow_up_eligible_2 <- subset(raw_data_baseline_arm_1, 
                             tbip_sc_ini_days_calc < 14 &
                               tbr_smear_res_1 == 'Negative')

follow_up_eligible_2 <- follow_up_eligible_2 %>%select(record_id, tbip_sc_ini_date, tbip_sc_q5, tbr_smear_res_1)

follow_up_eligible_2 <- follow_up_eligible_2 %>%mutate(week_7 = tbip_sc_ini_date + weeks(7))

follow_up_eligible_2 <- follow_up_eligible_2 %>%relocate(week_7, .after = tbr_smear_res_1)

follow_up_eligible_2 <- follow_up_eligible_2[c('record_id', 'tbip_sc_ini_date', 'tbip_sc_q5', 'tbr_smear_res_1', 'week_7')]

write.table(follow_up_eligible_2, "Data/follow up 2 eligible.csv", sep = ",", row.names = FALSE)


#eligible
not_tb_positive <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q8=='No')

not_tb_positive <- not_tb_positive[c('record_id', 'tbip_sc_q8')]


#follow up reasons
follow_up_reasons <- subset(raw_data_follow_up_1_arm_1, !is.na(raw_data_follow_up_1_arm_1$sn_nature))

follow_up_reasons <- follow_up_reasons[c('record_id', 'sn_nature')]
                            
                            
#Follow up questionnaire not completed
fu_no_ques <- subset(raw_data_follow_up_1_arm_1, raw_data_follow_up_1_arm_1$tbr_smear_res_1=='Negative' & 
                       (raw_data_follow_up_1_arm_1$index_follow_up_questionnaire_3_complete=='Incomplete'))

fu_no_ques <- fu_no_ques[c('record_id', 'tbr_smear_res_1', 'index_follow_up_questionnaire_3_complete')]

write.table(fu_no_ques, "Data/Follow Questionnaire not completed.csv", sep = ",", row.names = FALSE)


#recommended dates
rec_date_2 <- subset(raw_data_hhci_info_arm_1, !is.na(raw_data_hhci_info_arm_1$hhc_sc_attempt_1_rec_date))

rec_date_2 <- rec_date_2[c('record_id', 'redcap_repeat_instance', 'hhc_sc_attempt_1_rec_date')]


rec_date_3 <- subset(raw_data_hhci_info_arm_1, !is.na(raw_data_hhci_info_arm_1$hhc_sc_attempt_1_rec_date_2))

rec_date_3 <- rec_date_3[c('record_id', 'redcap_repeat_instance', 'hhc_sc_attempt_1_rec_date_2')]


#Confirmation date
confirmation_date <- subset(raw_data_hhci_visit_info_arm_1, !is.na(raw_data_hhci_visit_info_arm_1$hhc_sch_hhi_date_visit))

confirmation_date <- confirmation_date[c('record_id', 'redcap_repeat_instance','hhc_sch_hhi_date_visit', 'hhc_sch_hhi_visit_time')]


#Incomplete questionnaires
incomplete_ques <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part &
                            (raw_data_baseline_arm_1$index_demographic_complete=='Incomplete'))


#gxp
gxp <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc < 14 & 
                (!is.na(raw_data_baseline_arm_1$tbr_genex_result_1)))

gxp <- gxp[c('record_id', 'tbip_sc_ini_days_calc', 'tbr_genex_result_1', 'tbr_gx_done_res_2')]


pos <- subset(raw_data_follow_up_1_arm_1, raw_data_follow_up_1_arm_1$tbr_smear_res_1=='Positive 1+' | raw_data_follow_up_1_arm_1$tbr_smear_res_1=='Positive 2+' | raw_data_follow_up_1_arm_1=='Positive 3+')

pos <- pos[c('record_id', 'tbr_smear_res_1')]


#
gxp_2 <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc < 14 & 
                (is.na(raw_data_baseline_arm_1$tbr_smear_res_1) & (!is.na(raw_data_baseline_arm_1$tbr_genex_result_1))))

gxp_2 <- gxp_2[c('record_id', 'tbr_smear_res_1', 'tbr_genex_result_1')]


#baseline Positive smear
base_pos_smear <- subset(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & (raw_data_baseline_arm_1$tbr_smear_res_1=='Positive 1+' | raw_data_baseline_arm_1$tbr_smear_1=='Positive 2+' | raw_data_baseline_arm_1$tbr_smear_res_1=='Positive 3+')))
  
base_pos_smear <- base_pos_smear[c('record_id', 'tbr_smear_1')]

write.table(base_pos_smear, "Data/ base_pos_smear.csv", sep = ",", row.names = FALSE)


#
pos_returned <- subset(raw_data_follow_up_1_arm_1, raw_data_follow_up_1_arm_1$tbr_sputum_collected=='Yes')

pos_returned <- pos_returned[c('record_id', 'tbr_sputum_collected')]


#Questionnaire completed
fu_qu <- subset(raw_data_follow_up_1_arm_1, 
                raw_data_follow_up_1_arm_1$index_follow_up_questionnaire_3_complete=='Unverified' | raw_data_follow_up_1_arm_1$index_follow_up_questionnaire_3_complete=="Complete")

fu_qu <- fu_qu[c('record_id', 'index_follow_up_questionnaire_3_complete')]


#
hhc_treatment <- subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhc_sc_on_treatment=='Yes')

hhc_treatment <- hhc_treatment[c('record_id', 'hhc_sc_on_treatment')]


hhc_age <- subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhc_sc_age_calc < 18)

hhc_age <- hhc_age[c('record_id', 'hhc_sc_age_calc')]



tb_asymp <- subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhc_sc_weight_loss=='No' &
                     (raw_data_hhci_info_arm_1$hhc_sc_night_sweat=='No') &
                     (raw_data_hhci_info_arm_1$hhc_sc_coughing=='No') &
                     (raw_data_hhci_info_arm_1$hhc_sc_fever=='No'))

tb_asymp <- tb_asymp[c('record_id', 'hhc_sc_weight_loss', 'hhc_sc_night_sweat', 'hhc_sc_coughing', 'hhc_sc_fever')]


#baseline smear done
smear_done <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbr_smear_1=='Yes' &
                       (is.na(raw_data_baseline_arm_1$tbr_smear_res_1)))

smear_done <- smear_done[c('record_id', 'tbip_sc_q5', 'tbr_res_date_1','tbr_smear_res_1')]


#no smear results
no_smear <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc < 14 &
                     is.na(raw_data_baseline_arm_1$tbr_smear_res_1) & (!is.na(raw_data_baseline_arm_1$tbr_genex_result_1)))

no_smear <- no_smear[c('record_id', 'tbip_sc_q5', 'tbip_sc_part_name', 'tbip_sc_part_surname', 'tbip_sc_dob', 'tbip_sc_consent_part', 'tbip_sc_ini_date','tbr_smear_res_1', 'tbr_genex_result_1')]

write.table(no_smear,"Data/no smear.csv", sep = ",", row.names = FALSE)


#No GeneXpert
no_gxp <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc < 14 & 
                   is.na(raw_data_baseline_arm_1$tbr_genex_result_1))

no_gxp <- no_gxp[c('record_id', 'tbip_sc_q5', 'tbip_sc_part_name', 'tbip_sc_part_surname', 'tbip_sc_dob', 'tbip_sc_consent_part', 'tbip_sc_ini_date', 'tbr_genex_result_1')]

write.table(no_gxp,"Data/no gxp.csv", sep = ",", row.names = FALSE)


#no smear/no gxp
no_results <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc < 14 & (raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes') &
                       (is.na(raw_data_baseline_arm_1$tbr_smear_res_1) & (is.na(raw_data_baseline_arm_1$tbr_genex_result_1))))

no_results <- no_results[c('record_id', 'tbip_sc_q5', 'tbip_sc_consent_part','tbip_sc_part_name', 'tbip_sc_part_surname', 'tbip_sc_dob', 'tbip_sc_ini_date', 'tbr_smear_res_1', 'tbr_genex_result_1')]

write.table(no_results,"Data/no results.csv", sep = ",", row.names = FALSE)

#pending visits
hhc_refused <- subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhc_sc_consent_provided=='No')

hhc_refused <- hhc_refused[c('record_id', 'redcap_repeat_instance','hhc_sc_consent_provided')]

write.table(hhc_refused, "Data/Household Contacts Refused.csv", sep = ",", row.names = FALSE)


#Under Age
under_age <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_below_age=='Yes' | raw_data_baseline_arm_1$tbip_sc_below_age=='No')

under_age <- under_age[c('record_id', 'tbip_sc_q5','tbip_sc_age_calc', 'tbip_sc_below_age')]

write.table(under_age,"Data/Under 18.csv", sep = ",", row.names = FALSE)


#Approached In the past 7 days
approached <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$approach_date >= Sys.Date() - 6 & 
                       (raw_data_baseline_arm_1$approach_facility=='Empilweni Gompo CHC'))

approached <- approached[c('record_id','approach_date', 'approach_facility')]

write.table(approached,"Data/approached.csv", sep = ",", row.names = FALSE)


#Follow up questionnaire not completed
fu_ques <- subset(raw_data_follow_up_1_arm_1, raw_data_follow_up_1_arm_1$tbr_smear_res_1=='Negative' & 
                       (raw_data_follow_up_1_arm_1$index_follow_up_questionnaire_3_complete=='Incomplete'))

fu_ques <- fu_no_ques[c('record_id', 'tbr_smear_res_1', 'index_follow_up_questionnaire_3_complete')]

write.table(fu_no_ques, "Data/Follow Questionnaire incomplete.csv", sep = ",", row.names = FALSE)


#New variable
fu_converted <- subset(raw_data_baseline_arm_1, !is.na(raw_data_baseline_arm_1$tbr_fu_present))

fu_converted <- fu_converted[c('record_id', 'tbip_sc_ini_days_calc','tbr_fu_present')]

write.table(fu_converted, "Data/fu converted.csv", sep = ",", row.names = FALSE)



#######################################
fu_converted_2 <- subset(raw_data_follow_up_1_arm_1, !is.na(raw_data_follow_up_1_arm_1$tbr_fu_present))

fu_converted_2 <- fu_converted_2[c('record_id','tbr_fu_present')]

write.table(fu_converted_2, "Data/fu converted follow up.csv", sep = ",", row.names = FALSE)


#################First attempt Adherence
first_attempt <- subset(temp, temp$visit_adherence_1=='Yes')

first_attempt <- first_attempt[c('record_id', 'visit_adherence_1')]

write.table(first_attempt, "Data/first attempt.csv", sep = ",", row.names = FALSE)


sec_attempt <- subset(temp, temp$visit_adherence_2=='Yes')

sec_attempt <- sec_attempt[c('record_id', 'visit_adherence_2')]


#####################################################################
clean_df <- na.omit(raw_data_baseline_arm_1)

table1::table1(~tbip_sc_q5+tbip_sc_q8+tbip_sc_age_calc | tbip_sc_consent_part, data=raw_data_baseline_arm_1 )


hoh <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_hoh=='Yes')

hoh <- hoh[c('record_id', 'tbip_sc_hh_visit_d_confirm', 'tbip_sc_date_visit_hh1', 'tbip_sc_date_time_hh1')]



#############################################################
tb_adherence <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$treat_out_intro=='Proceed')

tb_adherence <- tb_adherence[c('record_id', 'tbip_sc_ini_days_calc', 'treat_out_intro', 'treat_out_ex_visit_1', 'treat_out_ac_visit_1', 'treat_out_visit_1_calc', 'treat_out_proc_2', 'treat_out_ex_visit_2', 'treat_out_ac_visit_2', 'treat_out_visit_1_calc_2', 'treat_out_descr_2', 'treat_out_proc_3', 'treat_out_ex_visit_3')]


####################################HoH refused Household visit
ref_hh_visit <- subset(raw_data_hhci_visit_info_arm_1, raw_data_hhci_visit_info_arm_1$hhc_sch_confirm_visit_2=='No')

ref_hh_visit <- ref_hh_visit[c('record_id', 'hhc_sch_confirm_visit_2')]


write.table(dataset_hhd, "Data/HH Survey.csv", sep = ",", row.names = FALSE)

write.table(dataset_master, "Data/HH Survey Master.csv", sep = ",", row.names = FALSE)



#
out_com <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q17=='No')

out_com <- out_com[c('record_id', 'tbip_sc_date', 'tbip_sc_q5', 'tbip_sc_ini_date', 'tbip_sc_q17')]

write.table(out_com, "Data/outside community.csv", sep = ",", row.names = FALSE)


#
no_adherence <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' &
                         (is.na(raw_data_baseline_arm_1$treat_out_intro)))

no_adherence <- no_adherence[c('record_id',  'tbip_sc_q5', 'tbip_sc_consent_part','treat_out_intro')]

write.table(no_adherence, "Data/no adherence.csv", sep = ",", row.names = FALSE)
