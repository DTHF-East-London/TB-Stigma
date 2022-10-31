if (!require("tidyverse")) install.packages("tidyverse", dependencies = TRUE)

library(openxlsx)
library(tidyverse)
library(dplyr)
library(redcapAPI)
library(RMySQL)
library(summarytools)

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


#}

#Generate summary
stview(dfSummary(dataset_hhd))
summarytools::dfSummary()
save(dfSummary(dataset_hhd))
