if (!require("tidyverse")) install.packages("tidyverse", dependencies = TRUE)
library(openxlsx)
library(tidyverse)
library(dplyr)
library(redcapAPI)
library(RMySQL)
library(summarytools)
library(readxl)

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

dataset_hh_final = subset(dataset_hh, select = -c(2,3) )


############################# Screening And Consenting #########################

forms <- c('labelling_hh_on_google_maps', 'screening_and_consenting')

dataset_sc <- getREDCapRecords(NULL, forms, NULL)

dataset_sc <- dataset_sc[c(1:3,13:57)]

dataset_sc <- subset(dataset_sc, !is.na(dataset_sc$redcap_repeat_instrument))

dataset_screening <- subset(dataset_sc, select = -c(16:48))

dataset_screening_1 <- subset(dataset_screening, dataset_screening$redcap_repeat_instance==1)

dataset_screening_1 <- dataset_screening_1 %>% rename_with(~ paste(.x, "1", sep = "_"), -c(1))

dataset_screening_1 <- subset(dataset_screening_1, select = -c(2:3))

dataset_screening_2 <- subset(dataset_screening, dataset_screening$redcap_repeat_instance==2)

dataset_screening_2 <- dataset_screening_2 %>% rename_with(~ paste(.x, "2", sep = "_"), -c(1))

dataset_screening_2 <- subset(dataset_screening_2, select = -c(2:3))

dataset_screening_3 <- subset(dataset_screening, dataset_screening$redcap_repeat_instance==3)

dataset_screening_3 <- dataset_screening_3 %>% rename_with(~ paste(.x, "3", sep = "_"), -c(1))

dataset_screening_3 <- subset(dataset_screening_3, select = -c(2:3))

#dataset_screening_4 <- subset(dataset_screening, dataset_screening$redcap_repeat_instance==4)

#dataset_screening_4 <- dataset_screening_4 %>% rename_with(~ paste(.x, "4", sep = "_"), -c(1))

#dataset_screening_4 <- subset(dataset_screening_4, select = -c(2:3))

dataset_screening <- left_join(dataset_screening_1, dataset_screening_2, by = 'record_id') %>% left_join(., dataset_screening_3, by = 'record_id')

dataset_consenting <- subset(dataset_sc, select = -c(2:15))

dataset_consenting <- subset(dataset_consenting, !is.na(dataset_consenting$intro_script1))

dataset_sc_final <- left_join(dataset_screening, dataset_consenting, by = 'record_id')

dataset_sc_final <- subset(dataset_sc_final, dataset_sc_final$did_the_person_consent_to==1)

rm()
############################# head_of_household_demographics###################

forms <- c('labelling_hh_on_google_maps', 'head_of_household_demographics')

dataset_hhd <- getREDCapRecords(NULL, forms, NULL)

dataset_hhd <- subset(dataset_hhd, select = -c(2:12))

dataset_hhd <- subset(dataset_hhd, !is.na(dataset_hhd$language_prefer))

dataset_hhd_eng <- subset(dataset_hhd, dataset_hhd$language_prefer=='1')

dataset_hhd_xho <- subset(dataset_hhd, dataset_hhd$language_prefer=='2')

metadata_hhd_eng <- subset(metadata_hhd, metadata_hhd$English=='1')

metadata_hhd_xho <- subset(metadata_hhd, metadata_hhd$Xhosa=='1')

dataset_hhd_eng <- dataset_hhd_eng[,metadata_hhd_eng$`Variable / Field Name`]

dataset_hhd_xho <- dataset_hhd_xho[,metadata_hhd_xho$`Variable / Field Name`]

names(dataset_hhd_xho) <- names(dataset_hhd_eng)

dataset_hhd_final <- rbind(dataset_hhd_eng, dataset_hhd_xho)

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

dataset_que_final <- rbind(dataset_que_eng, dataset_que_xho)


###################### Proof Of Reimbursement And Snack ########################

forms <- c('labelling_hh_on_google_maps', 'proof_of_reimbursement_and_snack')

metadata_prs <- subset(TBStigmaHouseholdSurvey_DataDictionary_2022_08_05, TBStigmaHouseholdSurvey_DataDictionary_2022_08_05$`Form Name` =='proof_of_reimbursement_and_snack' | TBStigmaHouseholdSurvey_DataDictionary_2022_08_05$`Variable / Field Name`=='record_id')

metadata_prs <- subset(metadata_prs, metadata_prs$English=='1')

dataset_prs <- getREDCapRecords(NULL, forms, NULL)

dataset_prs_final <- dataset_prs[,metadata_prs$`Variable / Field Name`]


###################### Proof Of Reimbursement And Snack ########################

forms <- c('labelling_hh_on_google_maps', 'study_notes')

metadata_sns <- subset(TBStigmaHouseholdSurvey_DataDictionary_2022_08_05, TBStigmaHouseholdSurvey_DataDictionary_2022_08_05$`Form Name` =='study_notes' | TBStigmaHouseholdSurvey_DataDictionary_2022_08_05$`Variable / Field Name`=='record_id')

metadata_sns <- subset(metadata_sns, metadata_sns$English=='1')

dataset_sns <- getREDCapRecords(NULL, forms, NULL)

dataset_sns <- dataset_sns[c(1:3,13:21)]

dataset_sns_1 <- subset(dataset_sns, dataset_sns$redcap_repeat_instance==1)

dataset_sns_1 <- dataset_sns_1 %>% rename_with(~ paste(.x, "1", sep = "_"), -c(1))

dataset_sns_1 <- subset(dataset_sns_1, select = -c(2:3))

dataset_sns_2 <- subset(dataset_sns, dataset_sns$redcap_repeat_instance==2)

dataset_sns_2 <- dataset_sns_2 %>% rename_with(~ paste(.x, "2", sep = "_"), -c(1))

dataset_sns_2 <- subset(dataset_sns_2, select = -c(2:3))

dataset_sns_3 <- subset(dataset_sns, dataset_sns$redcap_repeat_instance==3)

dataset_sns_3 <- dataset_sns_3 %>% rename_with(~ paste(.x, "3", sep = "_"), -c(1))

dataset_sns_3 <- subset(dataset_sns_3, select = -c(2:3))

dataset_sns_final <- left_join(dataset_sns_1, dataset_sns_2, by = "record_id") %>% left_join(., dataset_sns_3, by = "record_id") 

######################### Internal Quality Control #############################

forms <- c('labelling_hh_on_google_maps', 'internal_quality_control')

metadata_iqc <- subset(TBStigmaHouseholdSurvey_DataDictionary_2022_08_05, TBStigmaHouseholdSurvey_DataDictionary_2022_08_05$`Form Name` =='internal_quality_control' | TBStigmaHouseholdSurvey_DataDictionary_2022_08_05$`Variable / Field Name`=='record_id')

metadata_iqc <- subset(metadata_iqc, metadata_iqc$English=='1')

dataset_iqc <- getREDCapRecords(NULL, forms, NULL)

dataset_iqc <- subset(dataset_iqc, select = -c(2:12))

dataset_iqc_final <- subset(dataset_iqc, !is.na(dataset_iqc$ra_name3))

#dataset_iqc <- dataset_iqc[,metadata_iqc$`Variable / Field Name`]

full_dataset_master <- right_join(dataset_hh_final, dataset_sc_final, by = "record_id") %>% left_join(., dataset_hhd_final, by = "record_id") %>% left_join(., dataset_que_final, by = "record_id") %>% left_join(., dataset_prs_final, by = "record_id") %>% left_join(., dataset_sns_final, by = "record_id") %>% left_join(., dataset_iqc_final, by = "record_id")

#Oversampled Squares
#Import from excel file
Oversurveyed_squares <- readxl::read_excel("Metadata/Oversurveyed_squares.xlsx", 
                                   col_types = c("text", "numeric", "numeric", 
                                                 "text"))


Oversurveyed_squares$AreaCode <- NA

Oversurveyed_squares$AreaCode[Oversurveyed_squares$Area == "Duncan Village"] <- 1
Oversurveyed_squares$AreaCode[Oversurveyed_squares$Area == "Scenery Park"] <- 2
Oversurveyed_squares$AreaCode[Oversurveyed_squares$Area == "Nompumelelo"] <- 3
Oversurveyed_squares$AreaCode[Oversurveyed_squares$Area == "Ducats"] <- 4
Oversurveyed_squares$AreaCode[Oversurveyed_squares$Area == "Mdantsane"] <- 5
Oversurveyed_squares$AreaCode[Oversurveyed_squares$Area == "Ndevana"] <- 6
Oversurveyed_squares$AreaCode[Oversurveyed_squares$Area == "Buffalo Flats"] <- 8

Oversurveyed_squares <- Oversurveyed_squares %>% relocate(AreaCode, .before = Area)

Oversurveyed_squares$LetterCode <- NA

Oversurveyed_squares$LetterCode[Oversurveyed_squares$Letter == "A"] <- 1
Oversurveyed_squares$LetterCode[Oversurveyed_squares$Letter == "B"] <- 2
Oversurveyed_squares$LetterCode[Oversurveyed_squares$Letter == "C"] <- 3
Oversurveyed_squares$LetterCode[Oversurveyed_squares$Letter == "D"] <- 4

Oversurveyed_squares <- Oversurveyed_squares %>% relocate(LetterCode, .before = Letter)

Oversurveyed_squares$AreaCode <- as.integer(Oversurveyed_squares$AreaCode)
Oversurveyed_squares$Community <- as.integer(Oversurveyed_squares$Community)
Oversurveyed_squares$Square <- as.integer(Oversurveyed_squares$Square)


full_dataset_master$oversampled <- 0

full_dataset_master <- full_dataset_master %>% relocate(oversampled, .after = "record_id")

#for(i in 1:nrow(Oversurveyed_squares)){  print(Oversurveyed_squares[i,5])
#    full_dataset_master <- within(full_dataset_master, {
#      
#      print(area_1 == as.integer(Oversurveyed_squares[i,2]))
#      oversampled[(area_1 == as.integer(Oversurveyed_squares[i,1]) | area_1 == Oversurveyed_squares[i,2]) & 
#                    study_community_1 == as.integer(Oversurveyed_squares[i,3]) &
#                    square_number_1 == as.integer(Oversurveyed_squares[i,4]) &
#                    (sub_square_letter_1 == as.integer(Oversurveyed_squares[i,5]) |
#                       is.na(Oversurveyed_squares[i,5]) |
#                       sub_square_letter_1 == as.integer(Oversurveyed_squares[i,6]))] <- 1
#      
#    })
#}

#Ducats
full_dataset_master$oversampled[(full_dataset_master$area_1=="Ducats" | full_dataset_master$area_1== 4) & full_dataset_master$study_community_1 == 1 & full_dataset_master$square_number_1 == 5] <- 1
full_dataset_master$oversampled[(full_dataset_master$area_1=="Ducats" | full_dataset_master$area_1== 4) & full_dataset_master$study_community_1 == 1 & full_dataset_master$square_number_1 == 6] <- 1
full_dataset_master$oversampled[(full_dataset_master$area_1=="Ducats" | full_dataset_master$area_1== 4) & full_dataset_master$study_community_1 == 1 & full_dataset_master$square_number_1 == 8] <- 1

#Nompumelelo
full_dataset_master$oversampled[(full_dataset_master$area_1=="Nompumelelo" | full_dataset_master$area_1== 3) & full_dataset_master$study_community_1 == 2 & full_dataset_master$square_number_1 == 7] <- 1

#Buffalo Flats
full_dataset_master$oversampled[(full_dataset_master$area_1=="Buffalo Flats" | full_dataset_master$area_1== 8) & full_dataset_master$study_community_1 == 4 & full_dataset_master$square_number_1 == 1 & (full_dataset_master$sub_square_letter_1 == "C" | full_dataset_master$sub_square_letter_1 == 3)] <- 1
full_dataset_master$oversampled[(full_dataset_master$area_1=="Buffalo Flats" | full_dataset_master$area_1== 8) & full_dataset_master$study_community_1 == 4 & full_dataset_master$square_number_1 == 1 & (full_dataset_master$sub_square_letter_1 == "D" | full_dataset_master$sub_square_letter_1 == 4)] <- 1
full_dataset_master$oversampled[(full_dataset_master$area_1=="Buffalo Flats" | full_dataset_master$area_1== 8) & full_dataset_master$study_community_1 == 4 & full_dataset_master$square_number_1 == 4 & (full_dataset_master$sub_square_letter_1 == "A" | full_dataset_master$sub_square_letter_1 == 1)] <- 1

#Scenery Park
full_dataset_master$oversampled[(full_dataset_master$area_1=="Scenery Park" | full_dataset_master$area_1== 2) & full_dataset_master$study_community_1 == 7 & full_dataset_master$square_number_1 == 3 & (full_dataset_master$sub_square_letter_1 == "B" | full_dataset_master$sub_square_letter_1 == 2)] <- 1
full_dataset_master$oversampled[(full_dataset_master$area_1=="Scenery Park" | full_dataset_master$area_1== 2) & full_dataset_master$study_community_1 == 7 & full_dataset_master$square_number_1 == 3 & (full_dataset_master$sub_square_letter_1 == "C" | full_dataset_master$sub_square_letter_1 == 3)] <- 1
full_dataset_master$oversampled[(full_dataset_master$area_1=="Scenery Park" | full_dataset_master$area_1== 2) & full_dataset_master$study_community_1 == 8 & full_dataset_master$square_number_1 == 2 & (full_dataset_master$sub_square_letter_1 == "B" | full_dataset_master$sub_square_letter_1 == 2)] <- 1

#Duncan
full_dataset_master$oversampled[(full_dataset_master$area_1=="Duncan Village" | full_dataset_master$area_1== 1) & full_dataset_master$study_community_1 == 3 & full_dataset_master$square_number_1 == 1 & (full_dataset_master$sub_square_letter_1 == "D" | full_dataset_master$sub_square_letter_1 == 4)] <- 1
full_dataset_master$oversampled[(full_dataset_master$area_1=="Duncan Village" | full_dataset_master$area_1== 1) & full_dataset_master$study_community_1 == 3 & full_dataset_master$square_number_1 == 4 & (full_dataset_master$sub_square_letter_1 == "C" | full_dataset_master$sub_square_letter_1 == 3)] <- 1
full_dataset_master$oversampled[(full_dataset_master$area_1=="Duncan Village" | full_dataset_master$area_1== 1) & full_dataset_master$study_community_1 == 3 & full_dataset_master$square_number_1 == 4 & (full_dataset_master$sub_square_letter_1 == "D" | full_dataset_master$sub_square_letter_1 == 4)] <- 1
full_dataset_master$oversampled[(full_dataset_master$area_1=="Duncan Village" | full_dataset_master$area_1== 1) & full_dataset_master$study_community_1 == 4 & full_dataset_master$square_number_1 == 2 & (full_dataset_master$sub_square_letter_1 == "B" | full_dataset_master$sub_square_letter_1 == 2)] <- 1
full_dataset_master$oversampled[(full_dataset_master$area_1=="Duncan Village" | full_dataset_master$area_1== 1) & full_dataset_master$study_community_1 == 4 & full_dataset_master$square_number_1 == 2 & (full_dataset_master$sub_square_letter_1 == "C" | full_dataset_master$sub_square_letter_1 == 3)] <- 1
full_dataset_master$oversampled[(full_dataset_master$area_1=="Duncan Village" | full_dataset_master$area_1== 1) & full_dataset_master$study_community_1 == 4 & full_dataset_master$square_number_1 == 2 & (full_dataset_master$sub_square_letter_1 == "D" | full_dataset_master$sub_square_letter_1 == 4)] <- 1
full_dataset_master$oversampled[(full_dataset_master$area_1=="Duncan Village" | full_dataset_master$area_1== 1) & full_dataset_master$study_community_1 == 4 & full_dataset_master$square_number_1 == 4 & (full_dataset_master$sub_square_letter_1 == "B" | full_dataset_master$sub_square_letter_1 == 2)] <- 1
full_dataset_master$oversampled[(full_dataset_master$area_1=="Duncan Village" | full_dataset_master$area_1== 1) & full_dataset_master$study_community_1 == 4 & full_dataset_master$square_number_1 == 4 & (full_dataset_master$sub_square_letter_1 == "D" | full_dataset_master$sub_square_letter_1 == 4)] <- 1
full_dataset_master$oversampled[(full_dataset_master$area_1=="Duncan Village" | full_dataset_master$area_1== 1) & full_dataset_master$study_community_1 == 4 & full_dataset_master$square_number_1 == 4 & (full_dataset_master$sub_square_letter_1 == "D" | full_dataset_master$sub_square_letter_1 == 4)] <- 1
full_dataset_master$oversampled[(full_dataset_master$area_1=="Duncan Village" | full_dataset_master$area_1== 1) & full_dataset_master$study_community_1 == 4 & full_dataset_master$square_number_1 == 6 & (full_dataset_master$sub_square_letter_1 == "A" | full_dataset_master$sub_square_letter_1 == 1)] <- 1
full_dataset_master$oversampled[(full_dataset_master$area_1=="Duncan Village" | full_dataset_master$area_1== 1) & full_dataset_master$study_community_1 == 4 & full_dataset_master$square_number_1 == 6 & (full_dataset_master$sub_square_letter_1 == "B" | full_dataset_master$sub_square_letter_1 == 2)] <- 1
full_dataset_master$oversampled[(full_dataset_master$area_1=="Duncan Village" | full_dataset_master$area_1== 1) & full_dataset_master$study_community_1 == 5 & full_dataset_master$square_number_1 == 2 & (full_dataset_master$sub_square_letter_1 == "C" | full_dataset_master$sub_square_letter_1 == 3)] <- 1
full_dataset_master$oversampled[(full_dataset_master$area_1=="Duncan Village" | full_dataset_master$area_1== 1) & full_dataset_master$study_community_1 == 6 & full_dataset_master$square_number_1 == 4 & (full_dataset_master$sub_square_letter_1 == "B" | full_dataset_master$sub_square_letter_1 == 2)] <- 1
full_dataset_master$oversampled[(full_dataset_master$area_1=="Duncan Village" | full_dataset_master$area_1== 1) & full_dataset_master$study_community_1 == 13 & full_dataset_master$square_number_1 == 10] <- 1
full_dataset_master$oversampled[(full_dataset_master$area_1=="Duncan Village" | full_dataset_master$area_1== 1) & full_dataset_master$study_community_1 == 14 & full_dataset_master$square_number_1 == 7] <- 1

#Ndevana
full_dataset_master$oversampled[(full_dataset_master$area_1=="Ndevana" | full_dataset_master$area_1== 6) & full_dataset_master$study_community_1 == 3 & full_dataset_master$square_number_1 == 8] <- 1
full_dataset_master$oversampled[(full_dataset_master$area_1=="Ndevana" | full_dataset_master$area_1== 6) & full_dataset_master$study_community_1 == 3 & full_dataset_master$square_number_1 == 9] <- 1
full_dataset_master$oversampled[(full_dataset_master$area_1=="Ndevana" | full_dataset_master$area_1== 6) & full_dataset_master$study_community_1 == 5 & full_dataset_master$square_number_1 == 11] <- 1

#Mdantsane
full_dataset_master$oversampled[(full_dataset_master$area_1=="Ndevana" | full_dataset_master$area_1== 5) & full_dataset_master$study_community_1 == 1 & full_dataset_master$square_number_1 == 4] <- 1
full_dataset_master$oversampled[(full_dataset_master$area_1=="Ndevana" | full_dataset_master$area_1== 5) & full_dataset_master$study_community_1 == 3 & full_dataset_master$square_number_1 == 6] <- 1
full_dataset_master$oversampled[(full_dataset_master$area_1=="Ndevana" | full_dataset_master$area_1== 5) & full_dataset_master$study_community_1 == 6 & full_dataset_master$square_number_1 == 7] <- 1
full_dataset_master$oversampled[(full_dataset_master$area_1=="Ndevana" | full_dataset_master$area_1== 5) & full_dataset_master$study_community_1 == 6 & full_dataset_master$square_number_1 == 8] <- 1
full_dataset_master$oversampled[(full_dataset_master$area_1=="Ndevana" | full_dataset_master$area_1== 5) & full_dataset_master$study_community_1 == 6 & full_dataset_master$square_number_1 == 9] <- 1
full_dataset_master$oversampled[(full_dataset_master$area_1=="Ndevana" | full_dataset_master$area_1== 5) & full_dataset_master$study_community_1 == 6 & full_dataset_master$square_number_1 == 10] <- 1

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
#stview(dfSummary(dataset_hhd))
#summarytools::dfSummary()
#save(dfSummary(dataset_hhd))

library(haven)
require(foreign)
write.csv(full_dataset_master,'C:/Users/nkagisangn/OneDrive - foundation.co.za/Documents/GitHub/TB-Stigma/Data/full_dataset_master.csv',row.names = FALSE)
write.dta(full_dataset_master,'C:/Users/nkagisangn/OneDrive - foundation.co.za/Documents/GitHub/TB-Stigma/Data/full_dataset_master.dta')


