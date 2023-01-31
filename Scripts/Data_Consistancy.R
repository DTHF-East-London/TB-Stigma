############################### Library ########################################
library(redcapAPI)

library(dplyr)

###############################################################################

#study <- 1

#getREDCapConnection(study)

# Import data from REDCap


TBStigmaHouseholdSurvey_DataDictionary_2022_08_05 <- read_csv("Metadata/TBStigmaHouseholdSurvey_DataDictionary_2022-08-05.csv")

forms <- c('screening_and_consenting', 'head_of_household_demographics', 'questionnaire')

dataset_master <- getREDCapRecords(NULL, NULL, NULL)

metadata_hhd <- subset(TBStigmaHouseholdSurvey_DataDictionary_2022_08_05, TBStigmaHouseholdSurvey_DataDictionary_2022_08_05$`Form Name` =='head_of_household_demographics' | TBStigmaHouseholdSurvey_DataDictionary_2022_08_05$`Variable / Field Name`=='record_id')


baseline_fields <- c("record_id", 'area', "	screening_date", "lang_fluent", "interest_status", 
                     "preferred_language_to_be_u", "did_the_person_consent_to")

#df_screening_consenting <- df_baseline[screening_and_consenting_fields]