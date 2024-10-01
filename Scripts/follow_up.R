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

#Follow Up Dataset
data_1 <- raw_data_baseline_arm_1[c('record_id', 'tbip_sc_date', 'tbip_sc_ini_date','tbip_sc_ini_days_calc','tbip_sc_q5', 'tbip_sc_consent_part','tbr_fu_present', 'tbr_smear_1', 'tbr_smear_res_1', 'tbr_smear_2', 'tbr_smear_res_2', 'tbr_genex_1', 'tbr_genex_result_1', 'tbr_genex_2', 'tbr_gx_done_res_2', 'treat_out_outcome')]

data_2 <- raw_data_follow_up_1_arm_1[c('record_id', 'tbip_f1_date', 'tbr_sputum_collected', 'tbr_act_sputum_collection_date', 'tbr_fu_present', 'tbr_smear_1', 'tbr_res_date_1', 'tbr_smear_res_1', 'tbr_smear_2', 'tbr_smear_res_2', 'tbr_res_date_2','tbr_genex_1', 'tbr_genex_result_1', 'tbr_genex_2', 'tbr_gx_done_res_2', 'index_follow_up_questionnaire_1_complete', 'index_follow_up_questionnaire_2_complete', 'index_follow_up_questionnaire_3_complete')]

follow_up_data <- left_join(data_1, data_2, by='record_id')

#Follow Up 1 Period Open
follow_up_data <- follow_up_data %>%mutate(fu1_period_open = tbr_act_sputum_collection_date + weeks(1))

follow_up_data <- follow_up_data %>%relocate(fu1_period_open, .after = tbr_act_sputum_collection_date)


#Follow Up period 2 close date
follow_up_data <- follow_up_data %>%mutate(fu1_period_close = fu1_period_open + weeks(4))

follow_up_data <- follow_up_data %>%relocate(fu1_period_close, .after = fu1_period_open)


#Follow up 1 window for participants who had positive smears open
follow_up_data <- follow_up_data %>%mutate(fu1_pos_per_open = tbr_res_date_2 + weeks(1))

follow_up_data <- follow_up_data %>%relocate(fu1_pos_per_open, .after = tbr_res_date_1)


#Follow up 1 window for participants who had positive smears close
follow_up_data <- follow_up_data %>%mutate(fu1_pos_per_close = fu1_pos_per_open + weeks(4))

follow_up_data <- follow_up_data %>%relocate(fu1_pos_per_close, .after = fu1_pos_per_open)

write.table(follow_up_data, "Data/follow up data.csv", sep = ",", row.names = FALSE)



########################################################################################################################################################
#Follow up 2
df1 <- follow_up_data

df2 <- raw_data_follow_up_2_arm_1[c('record_id', 'tbip_f1_date','tbr_exp_sputum_collection_date', 'tbr_sputum_collected', 'tbr_act_sputum_collection_date', 'tbr_results_available', 'index_follow_up_questionnaire_3_complete')]

follow_up_2_data <- left_join(df1, df2, by='record_id')


#FU 2 window opens
follow_up_2_data <- follow_up_2_data %>%mutate(fu2_period_open = tbip_sc_ini_date + weeks(21))

follow_up_2_data <- follow_up_2_data %>%relocate(fu2_period_open, .after = tbr_act_sputum_collection_date.y)


#FU 2 window close
follow_up_2_data <- follow_up_2_data %>%mutate(fu2_period_close = fu2_period_open + weeks(6))

follow_up_2_data <- follow_up_2_data %>%relocate(fu2_period_close, .after = fu2_period_open)


#FU 2 calling window opens
follow_up_2_data <- follow_up_2_data %>%mutate(fu2_calling_period_open = fu2_period_open + weeks(3))

follow_up_2_data <- follow_up_2_data %>%relocate(fu2_calling_period_open, .after = tbr_results_available)


#FU 2 calling window close
follow_up_2_data <- follow_up_2_data %>%mutate(fu2_calling_period_close = fu2_calling_period_open + weeks(3))

follow_up_2_data <- follow_up_2_data %>%relocate(fu2_calling_period_close, .after = fu2_calling_period_open)


write.table(follow_up_2_data, "Data/follow up 2 data.csv", sep = ",", row.names = FALSE)
