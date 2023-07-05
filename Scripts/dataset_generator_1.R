if (!require("tidyverse")) install.packages("tidyverse", dependencies = TRUE)
library(openxlsx)
library(tidyverse)
library(dplyr)
library(redcapAPI)
library(RMySQL)
library(summarytools)
library(readxl)
library(haven)
library(xlsx)
library(tableone)
library(survival)


source("Scripts/functions.R")


#Get REDCap connection
print("getting REDCap connection")
rcon <- getREDCapConnection(2)
path <- "./Data/"
output_file <- paste0('dataset',format(Sys.time(), '%d_%B_%Y'),'.xlsx')

events <- exportEvents(rcon)

events <- as.list(events$unique_event_name)

instruments <- exportMappings(rcon)

for(event in events){
  forms <- subset(instruments, instruments$unique_event_name==event)
  forms <- as.vector(forms$form)

  if(event!="baseline_arm_1"){
    forms <- append(forms, "index_screening_and_consent",1)
  }
  
  if(event!="follow_up_1_arm_1" & event!="follow_up_2_arm_1"){
    
    temp <- getREDCapRecords(event, forms, NULL, TRUE)
    
    if(event!="baseline_arm_1"){
      temp <- temp[-c(5:100)]
    }
  
    assign(paste('raw_data', event, sep = '_'), temp)
  }
  
}

#Adding additional calculated variables
raw_data_hhci_info_arm_1$hhc_

#Code missing information as NI
levels(raw_data_hhci_info_arm_1$hhc_sc_clinic_visit) <- c('No', 'Yes', 'NI')
raw_data_hhci_info_arm_1$hhc_sc_clinic_visit[!is.na(raw_data_hhci_info_arm_1$hhc_sc_verbal_consent)] <- 'NI'
levels(raw_data_hhci_info_arm_1$hhc_sc_provide_sputum) <- c('No', 'Yes', 'NI')
raw_data_hhci_info_arm_1$hhc_sc_provide_sputum[!is.na(raw_data_hhci_info_arm_1$hhc_sc_verbal_consent)] <- 'NI'

raw_data_baseline_arm_1$tbip_sc_eligible[raw_data_baseline_arm_1$record_id == '145'] <- 'Proceed'

#raw_data_baseline_arm_1 <- raw_data_baseline_arm_1 %>% 
#  select(record_id, contains("tbip_q1_pc_q")) %>% 
#  rowwise() %>%
#  mutate(tbip_q1_pc_count = sum(!is.na)
         
#raw_data_baseline_arm_1 <- raw_data_baseline_arm_1 %>%       
#  mutate(tbip_q1_pc_count = rowSums(!is.na(contains("tbip_q1_pc_q")))) %>%
#  relocate(.data, after=tbip_q1_pc_q14)

