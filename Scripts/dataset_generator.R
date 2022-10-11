if (!require("tidyverse")) install.packages("tidyverse", dependencies = TRUE)

library(openxlsx)
library(tidyverse)
library(dplyr)
library(redcapAPI)
library(RMySQL)

source("Scripts/functions.R")

#Get REDCap connection
print("getting REDCap connection")
rcon <- getREDCapConnection(1)
path <- "./Data/"
output_file <- paste0('dataset',format(Sys.time(), '%d_%B_%Y'),'.xlsx')

TBStigmaHouseholdSurvey_DataDictionary_2022_08_05 <- read_csv("Metadata/TBStigmaHouseholdSurvey_DataDictionary_2022-08-05.csv")

#metadata_eng <- subset(TBStigmaHouseholdSurvey_DataDictionary_2022_08_05, TBStigmaHouseholdSurvey_DataDictionary_2022_08_05$English=='1')
  
#metadata_xho <- subset(TBStigmaHouseholdSurvey_DataDictionary_2022_08_05, TBStigmaHouseholdSurvey_DataDictionary_2022_08_05$Xhosa=='1')

dataset_master <- getREDCapRecords(NULL, NULL, NULL)

metadata_hhd <- subset(TBStigmaHouseholdSurvey_DataDictionary_2022_08_05, TBStigmaHouseholdSurvey_DataDictionary_2022_08_05$`Form Name` =='head_of_household_demographics' | TBStigmaHouseholdSurvey_DataDictionary_2022_08_05$`Variable / Field Name`=='record_id')

#dataset_eng <- subset(dataset_master, dataset_master$language_prefer=='1')

#dataset_xho <- subset(dataset_master, dataset_master$language_prefer=='2')

#names(dataset_xho)<- names(dataset_eng)

#dataset_master_2 <- rbind(dataset_eng, dataset_xho)


############################# Labelling HH# on Google Maps #####################

forms <- c('labelling_hh_on_google_maps')

dataset_hh <- getREDCapRecords(NULL, forms, NULL)


############################# Screening And Consenting #########################

forms <- c('labelling_hh_on_google_maps', 'screening_and_consenting')

dataset_sc <- getREDCapRecords(NULL, forms, NULL)

dataset_sc <- dataset_sc[c(1:3,13:65)]

dataset_sc <- subset(dataset_sc, !is.na(dataset_sc$redcap_repeat_instrument))

dataset_scv <- dataset_sc

dataset_sc_1 <- subset(dataset_sc, dataset_sc$redcap_repeat_instance==1)

dataset_sc_1 <- dataset_sc_1 %>% rename_with(~ paste(.x, "1", sep = "_"), -c(1))

dataset_sc_2 <- subset(dataset_sc, dataset_sc$redcap_repeat_instance==2)

dataset_sc_2 <- dataset_sc_2 %>% rename_with(~ paste(.x, "2", sep = "_"), -c(1))

dataset_sc_3 <- subset(dataset_sc, dataset_sc$redcap_repeat_instance==3)

dataset_sc_3 <- dataset_sc_3 %>% rename_with(~ paste(.x, "3", sep = "_"), -c(1))

dataset_sc_4 <- subset(dataset_sc, dataset_sc$redcap_repeat_instance==4)

dataset_sc_4 <- dataset_sc_4 %>% rename_with(~ paste(.x, "4", sep = "_"), -c(1))

dataset_sch <- left_join(dataset_sc_1, dataset_sc_2, by = 'record_id') %>% left_join(., dataset_sc_3, by = 'record_id') %>% left_join(., dataset_sc_4, by = 'record_id')

rm()
############################# head_of_household_demographics###################

forms <- c('labelling_hh_on_google_maps', 'head_of_household_demographics')

dataset_hhd <- getREDCapRecords(NULL, forms, NULL)

dataset_hhd <- subset(dataset_hhd, !is.na(dataset_hhd$language_prefer))

dataset_hhd_eng <- subset(dataset_hhd, dataset_hhd$language_prefer=='1')

dataset_hhd_xho <- subset(dataset_hhd, dataset_hhd$language_prefer=='2')

metadata_hhd_eng <- subset(metadata_hhd, metadata_hhd$English=='1')

metadata_hhd_xho <- subset(metadata_hhd, metadata_hhd$Xhosa=='1')

dataset_hhd_eng <- dataset_hhd_eng[,metadata_hhd_eng$`Variable / Field Name`]

dataset_hhd_xho <- dataset_hhd_xho[,metadata_hhd_xho$`Variable / Field Name`]

names(dataset_hhd_xho) <- names(dataset_hhd_eng)

dataset_hhd <- rbind(dataset_hhd_eng, dataset_hhd_xho)

############################# Questionnaires ##################################

forms <- c('labelling_hh_on_google_maps', 'questionnaires')

metadata_que <- subset(TBStigmaHouseholdSurvey_DataDictionary_2022_08_05, TBStigmaHouseholdSurvey_DataDictionary_2022_08_05$`Form Name` =='questionnaires' | TBStigmaHouseholdSurvey_DataDictionary_2022_08_05$`Variable / Field Name`=='record_id')

dataset_que <- getREDCapRecords(NULL, forms, NULL)

dataset_que <- subset(dataset_que, !is.na(dataset_que$ques_language))

dataset_que_eng <- subset(dataset_que, dataset_que$ques_language=='1')

dataset_que_xho <- subset(dataset_que, dataset_que$ques_language=='2')

metadata_que_eng <- subset(metadata_que, metadata_que$English=='1')

metadata_que_xho <- subset(metadata_que, metadata_que$Xhosa=='1')

dataset_que_eng <- dataset_que_eng[,metadata_que_eng$`Variable / Field Name`]

dataset_que_xho <- dataset_que_xho[,metadata_que_xho$`Variable / Field Name`]

names(dataset_que_xho) <- names(dataset_que_eng)

dataset_que <- rbind(dataset_que_eng, dataset_que_xho)


###################### Proof Of Reimbursement And Snack ########################

forms <- c('labelling_hh_on_google_maps', 'proof_of_reimbursement_and_snack')

metadata_prs <- subset(TBStigmaHouseholdSurvey_DataDictionary_2022_08_05, TBStigmaHouseholdSurvey_DataDictionary_2022_08_05$`Form Name` =='proof_of_reimbursement_and_snack' | TBStigmaHouseholdSurvey_DataDictionary_2022_08_05$`Variable / Field Name`=='record_id')

metadata_prs <- subset(metadata_prs, metadata_prs$English=='1')

dataset_prs <- getREDCapRecords(NULL, forms, NULL)

dataset_prs <- dataset_prs[,metadata_prs$`Variable / Field Name`]


###################### Proof Of Reimbursement And Snack ########################

forms <- c('labelling_hh_on_google_maps', 'study_notes')

metadata_sns <- subset(TBStigmaHouseholdSurvey_DataDictionary_2022_08_05, TBStigmaHouseholdSurvey_DataDictionary_2022_08_05$`Form Name` =='study_notes' | TBStigmaHouseholdSurvey_DataDictionary_2022_08_05$`Variable / Field Name`=='record_id')

metadata_sns <- subset(metadata_sns, metadata_sns$English=='1')

dataset_sns <- getREDCapRecords(NULL, forms, NULL)

dataset_sns <- dataset_sns[c(1:3,13:20)]

dataset_sns_1 <- subset(dataset_sns, dataset_sns$redcap_repeat_instance==1)

dataset_sns_1 <- dataset_sns_1 %>% rename_with(~ paste(.x, "1", sep = "_"), -c(1))

dataset_sns_2 <- subset(dataset_sns, dataset_sns$redcap_repeat_instance==2)

dataset_sns_2 <- dataset_sns_2 %>% rename_with(~ paste(.x, "2", sep = "_"), -c(1))

dataset_sns_3 <- subset(dataset_sns, dataset_sns$redcap_repeat_instance==3)

dataset_sns_3 <- dataset_sns_3 %>% rename_with(~ paste(.x, "3", sep = "_"), -c(1))

dataset_sns <- left_join(dataset_sns_1, dataset_sns_2, by = "record_id") %>% left_join(., dataset_sns_3, by = "record_id") 

######################### Internal Quality Control #############################

forms <- c('labelling_hh_on_google_maps', 'internal_quality_control')

metadata_iqc <- subset(TBStigmaHouseholdSurvey_DataDictionary_2022_08_05, TBStigmaHouseholdSurvey_DataDictionary_2022_08_05$`Form Name` =='internal_quality_control' | TBStigmaHouseholdSurvey_DataDictionary_2022_08_05$`Variable / Field Name`=='record_id')

metadata_iqc <- subset(metadata_iqc, metadata_iqc$English=='1')

dataset_iqc <- getREDCapRecords(NULL, forms, NULL)

dataset_iqc <- dataset_iqc[,metadata_iqc$`Variable / Field Name`]

#df_labelling_hh <- subset(dataset, is.na(dataset$redcap_repeat_instrument))

#df_labelling_hh <- df_labelling_hh %>% discard(~all(is.na(.) | . ==""))

#df_labelling_hh_e <- subset(df_labelling_hh, df_labelling_hh$language_prefer=='1')

#df_labelling_hh_e <- df_labelling_hh_e[,metadata_eng$`Variable / Field Name`]

#df_labelling_hh_e <- df_labelling_hh_e %>% discard(~all(is.na(.) | . ==""))

#df_labelling_hh_x <- subset(df_labelling_hh, df_labelling_hh$language_prefer=='2')

#df_labelling_hh_x <- df_labelling_hh_x[,metadata_xho$`Variable / Field Name`]
#df_labelling_hh_x <- df_labelling_hh_x %>% discard(~all(is.na(.) | . ==""))

#write.table(as.data.frame(metadata_xho$`Variable / Field Name`), file = 'Metadata/metadata_xho.csv', sep = ",", row.names = FALSE)
#write.table(as.data.frame(metadata_eng$`Variable / Field Name`), file = 'Metadata/metadata_eng.csv', sep = ",", row.names = FALSE)

#if (mysqlHasDefault()) {
  # connect to a database and load some data
  con <- dbConnect(RMySQL::MySQL(), 
                   dbname = "long_covid", 
                   host = 'localhost',
                   port = 3306,
                   user = "freedom",
                   password = "silverWATER9!8"
                   )
  dbWriteTable(con, "household", dataset_hh, overwrite = TRUE, row.names = FALSE,
               field.types = c(record_id = 'integer',
                              redcap_repeat_instrument = 'VARCHAR(30)',
                              redcap_repeat_instance = 'integer',
                              ra_name_label = 'VARCHAR(30)',
                              labelling_date = 'date',
                              labelling_time = 'time',
                              ra_instruct = 'integer',
                              attempt = 'integer',
                              ra_instruct_1stattempt = 'integer',
                              ra_instruct_2nd_3rdattempt = 'integer',
                              labelling_hh_on_google_maps_complete = 'integer'))
  dbWriteTable(con, "screeningh", dataset_sch, overwrite = TRUE, row.names = FALSE)
  dbWriteTable(con, "screeningv", dataset_scv, overwrite = TRUE, row.names = FALSE)
  dbWriteTable(con, "hhd", dataset_hhd, overwrite = TRUE, row.names = FALSE)
  dbWriteTable(con, "que", dataset_que, overwrite = TRUE, row.names = FALSE,
               field.types = c(record_id = "integer",
                               staff4 = "VARCHAR(30)",
                               date4 = "date",
                               time4 = "time",
                               ques_language = "integer",
                               hoh_last_clinic_any_e = "integer",
                               symptom_preffered_clinic = "integer",
                               tb_diagnosed_clinic_prefer = "integer",
                               hiv_clinic_prefer_test = "integer",
                               hiv_clinic_prefer_treat = "integer",
                               non_tb_care_util = "integer",
                               type_health_util___1 = "integer",
                               type_health_util___2 = "integer",
                               type_health_util___3 = "integer",
                               type_health_util___4 = "integer",
                               prison = "integer", prison_time = "integer",
                               prison_release = "integer", mines = "integer",
                               mines_time = "integer", mines_long_ago = "integer",
                               ss1 = "integer", ss2 = "integer", ss3 = "integer",
                               ss4 = "integer", ss5 = "integer", ss6 = "integer",
                               ss7 = "integer", ss8 = "integer", ss9 = "integer",
                               ss10 = "integer", ss11 = "integer", ss12 = "integer",
                               e1_eng = "integer", e2_eng = "integer", 
                               e3_eng = "integer", e4_eng = "integer",
                               e5_eng = "integer", e6_eng = "integer",
                               e7_eng = "integer", e8_eng = "integer",
                               e9_eng = "integer", e10_eng = "integer",
                               e11_eng = "integer", e12_eng = "integer",
                               e13_eng = "integer", e14_eng = "integer",
                               e15_eng = "integer", e16_eng = "integer",
                               e17_eng = "integer", e18_eng = "integer",
                               e21_eng = "integer", e22_eng = "integer",
                               e23_eng = "integer", e24_eng = "integer",
                               e25_eng = "integer", e26_eng = "integer",
                               e27_eng = "integer", e28_eng = "integer",
                               e29_eng = "integer", e30_eng = "integer",
                               rights_equity_1 = "integer", 
                               rights_equity_2 = "integer",
                               rights_equity_3 = "integer",
                               rights_equity_4 = "integer",
                               rights_equity_5 = "integer",
                               rights_equity_6 = "integer",
                               rights_equity_7 = "integer",
                               rights_equity_8 = "integer",
                               rights_equity_9 = "integer",
                               rights_equity_10 = "integer",
                               rights_equity_11 = "integer",
                               rights_equity_12 = "integer",
                               rights_equity_13 = "integer",
                               rights_equity_14 = "integer",
                               gender_norms_1 = "integer",
                               gender_norms_2 = "integer",
                               gender_norms_3 = "integer",
                               gender_norms_4 = "integer",
                               gender_norms_5 = "integer",
                               gender_norms_6 = "integer",
                               gender_norms_7 = "integer",
                               gender_norms_8 = "integer",
                               gender_norms_9 = "integer",
                               gender_norms_10 = "integer",
                               gender_norms_11 = "integer",
                               gender_norms_12 = "integer",
                               gender_norms_13 = "integer",
                               gender_norms_14 = "integer",
                               gender_norms_15 = "integer",
                               gender_norms_16 = "integer",
                               gender_norms_17 = "integer",
                               gender_role_1 = "integer",
                               gender_role_2 = "integer",
                               gender_role_3 = "integer",
                               gender_role_4 = "integer",
                               gender_role_5 = "integer",
                               gender_role_6 = "integer",
                               gender_role_7 = "integer",
                               gender_role_8 = "integer",
                               gender_role_9 = "integer",
                               gender_role_10 = "integer",
                               gender_role_11 = "integer",
                               gender_role_12 = "integer",
                               gender_role_13 = "integer",
                               gender_role_14 = "integer",
                               gender_role_15 = "integer",
                               gender_role_16 = "integer",
                               gender_role_17 = "integer",
                               gender_role_18 = "integer",
                               gender_role_19 = "integer",
                               gender_role_20 = "integer",
                               gender_role_21 = "integer",
                               gender_role_22 = "integer",
                               gender_role_23 = "integer",
                               gender_role_24 = "integer",
                               masc_chronic_d1 = "integer",
                               masc_chronic_d2 = "integer",
                               masc_chronic_d3 = "integer",
                               masc_chronic_d4 = "integer",
                               masc_chronic_d5 = "integer",
                               masc_chronic_d6 = "integer",
                               masc_chronic_d7 = "integer",
                               masc_chronic_d8 = "integer",
                               masc_chronic_d9 = "integer",
                               c1_eng = "integer",
                               c2_eng = "integer",
                               c3_eng = "integer",
                               c4_eng = "integer",
                               c4a_eng = "integer",
                               c5_eng = "integer",
                               c6_eng = "integer",
                               c7_eng = "integer",
                               c8_eng = "integer",
                               c8a_eng = "integer",
                               c_child_age = "integer",
                               c9_eng = "integer",
                               c10_eng = "integer",
                               c11_eng = "integer",
                               c12_eng = "integer",
                               c13_eng = "integer",
                               c13a_eng = "integer",
                               c14_eng = "integer",
                               c15_eng = "integer",
                               d1_eng = "integer",
                               d2_eng = "integer",
                               d3_eng = "integer",
                               d4_eng = "integer",
                               d5_eng = "integer",
                               d6_eng = "integer",
                               d7_eng = "integer",
                               d8_eng = "integer",
                               d9_eng = "integer",
                               d10_eng = "integer",
                               d11_eng = "integer",
                               d12_eng = "integer",
                               d13_eng = "integer",
                               d14_eng = "integer",
                               d15_eng = "integer",
                               d16_eng = "integer",
                               d17_eng = "integer",
                               d18_eng = "integer",
                               d19_eng = "integer",
                               d20_eng = "integer",
                               d21_eng = "integer",
                               d22_eng = "integer",
                               d23_eng = "integer",
                               hiv_know1 = "integer",
                               hiv_know2 = "integer",
                               hiv_know3 = "integer",
                               hiv_know4 = "integer",
                               hiv_know5 = "integer",
                               hiv_know6 = "integer",
                               hiv_know7 = "integer",
                               hiv_know8 = "integer",
                               hiv_know9 = "integer",
                               hiv_know10 = "integer",
                               hiv_know11 = "integer",
                               hiv_know12 = "integer",
                               hiv_know13 = "integer",
                               hiv_know14 = "integer",
                               hiv_know15 = "integer",
                               hiv_know16 = "integer",
                               hiv_know17 = "integer",
                               hiv_know18 = "integer",
                               s5_q1_p = "integer",
                               s5_q2_p = "integer",
                               s5_q3_p = "integer",
                               s5_q4_p = "integer",
                               s5_q5_p = "integer",
                               s5_q6_p = "integer",
                               s5_q7_p = "integer",
                               s5_q8_p = "integer",
                               s5_q9_p = "integer",
                               s5_q10_p = "integer",
                               s5_q11_p = "integer",
                               s5_q12_p = "integer",
                               s5_q13_p = "integer",
                               s5_q14_p = "integer",
                               tb_know1 = "integer",
                               tb_know2 = "integer",
                               tb_know3 = "integer",
                               tb_know4 = "integer",
                               tb_know5 = "integer",
                               tb_know6 = "integer",
                               tb_know7 = "integer",
                               tb_know8 = "integer",
                               tb_know9 = "integer",
                               tb_know10 = "integer",
                               tb_know11 = "integer",
                               s6_q1 = "integer",
                               s6_q2 = "integer",
                               s6_q3 = "integer",
                               s6_q4 = "integer",
                               s6_q5 = "integer",
                               s6_q6 = "integer",
                               s6_q7 = "integer",
                               s6_q8 = "integer",
                               s6_q9 = "integer",
                               s6_q10 = "integer",
                               a1_eng = "integer",
                               a2_eng = "integer",
                               a3_eng = "integer",
                               a4_eng = "integer",
                               a5_eng = "integer",
                               a6_eng = "integer",
                               a7_eng = "integer",
                               a8_eng = "integer",
                               a9_eng = "integer",
                               a10_eng = "integer",
                               a11_eng = "integer",
                               a12_eng = "integer",
                               a17_eng = "integer",
                               a29_eng = "integer",
                               a30_eng = "integer",
                               a31_eng = "integer",
                               a32_eng = "integer",
                               a33_eng = "integer",
                               a34_eng = "integer",
                               a35_eng = "integer",
                               a36_eng = "integer",
                               a37_eng = "integer",
                               a38_eng = "integer",
                               a39_eng = "integer",
                               a40_eng = "integer",
                               a41_eng = "integer",
                               a42_eng = "integer",
                               a43_eng = "integer",
                               a44_eng = "integer",
                               b1_eng = "integer",
                               b2_eng = "integer",
                               b3_eng = "integer",
                               b4_eng = "integer",
                               b5_eng = "integer",
                               b6_eng = "integer",
                               b7_eng = "integer",
                               b8_eng = "integer",
                               b9_eng = "integer",
                               b10_eng = "integer",
                               b11_eng = "integer",
                               b12_eng = "integer",
                               b13_eng = "integer",
                               b14_eng = "integer",
                               b15_eng = "integer",
                               b16_eng = "integer",
                               b17_eng = "integer",
                               b18_eng = "integer",
                               b19_eng = "integer",
                               b20_eng = "integer",
                               b21_eng = "integer",
                               b22_eng = "integer",
                               b23_eng = "integer",
                               b24_eng = "integer",
                               b25_eng = "integer",
                               b26_eng = "integer",
                               b27_eng = "integer",
                               b28_eng = "integer",
                               b29_eng = "integer",
                               b30_eng = "integer",
                               b31_eng = "integer",
                               b32_eng = "integer",
                               maks1 = "integer",
                               maks2 = "integer",
                               maks3 = "integer",
                               maks4 = "integer",
                               maks5 = "integer",
                               maks6 = "integer",
                               maks7 = "integer",
                               maks8 = "integer",
                               maks9 = "integer",
                               maks10 = "integer",
                               maks11 = "integer",
                               maks12 = "integer",
                               bog_1 = "integer",
                               bog_2 = "integer",
                               bog_3 = "integer",
                               bog_4 = "integer",
                               bog_5 = "integer",
                               bog_6 = "integer",
                               pss1_eng = "integer",
                               pss2_eng = "integer",
                               pss3_eng = "integer",
                               pss4_eng = "integer",
                               pss5_eng = "integer",
                               pss6_eng = "integer",
                               pss7_eng = "integer",
                               pss8_eng = "integer",
                               pss9_eng = "integer",
                               pss10_eng = "integer",
                               stig1_eng = "integer",
                               stig2_eng = "integer",
                               stig3_eng = "integer",
                               stig4_eng = "integer",
                               end_instruct = "integer",
                               end_time = "time",
                               questionnaires_complete = "integer"
               ))
  dbWriteTable(con, "prs", dataset_prs, overwrite = TRUE, row.names = FALSE)
  dbWriteTable(con, "sns", dataset_sns, overwrite = TRUE, row.names = FALSE)
  dbWriteTable(con, "iqc", dataset_iqc, overwrite = TRUE, row.names = FALSE)
  
  # query
  #rs <- dbSendQuery(con, "SELECT * FROM USArrests")
  #d1 <- dbFetch(rs, n = 10)      # extract data in chunks of 10 rows
  #dbHasCompleted(rs)
  #d2 <- dbFetch(rs, n = -1)      # extract all remaining data
  #dbHasCompleted(rs)
  #dbClearResult(rs)
  ##dbListTables(con)
  
  # clean up
  #dbRemoveTable(con, "USArrests")
  dbDisconnect(con)
#}
