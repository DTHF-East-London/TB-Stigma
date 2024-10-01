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
library(survival)
library(conflicted)

source("Scripts/functions.R")

#Get REDCap connection
print("getting REDCap connection")
rcon <- getREDCapConnection(2)
path <- "./Data/"
output_file <- paste0('dataset',format(Sys.time(), '%d_%B_%Y'),'.xlsx')

events <- exportEvents(rcon)

events <- as.list(events$unique_event_name)

instruments <- exportMappings(rcon)

today <- as.POSIXct(Sys.time())

for(event in events){
  forms <- subset(instruments, instruments$unique_event_name==event)
  forms <- as.vector(forms$form)
  
  print(paste0("Event: ", event))

  if(event!="baseline_arm_1"){
    forms <- append(forms, "index_screening_and_consent",0)
  }
  
  if(event!="follow_up_1_arm_1" & event!="follow_up_2_arm_1"){
    
    temp <- getREDCapRecords(event, forms, NULL, TRUE)
    
    if(event!="baseline_arm_1"){
      temp <- temp[-c(5:100)]
    }
   temp$record_id <- as.numeric(temp$record_id)
    assign(paste('raw_data', event, sep = '_'), temp)
  }
  
  if(event!="follow_up_1_arm_1"){
    
    temp <- getREDCapRecords(event, forms, NULL, TRUE)
    
    if(event!="baseline_arm_1"){
      temp <- temp[-c(5:100)]
    }
    temp$record_id <- as.numeric(temp$record_id)
    assign(paste('raw_data', event, sep = '_'), temp)
  }
  
  if(event!="follow_up_2_arm_1"){
    
    temp <- getREDCapRecords(event, forms, NULL, TRUE)
    
    if(event!="baseline_arm_1"){
      temp <- temp[-c(5:100)]
    }
    temp$record_id <- as.numeric(temp$record_id)
    assign(paste('raw_data', event, sep = '_'), temp)
  }
  
}

#Drop record 53

raw_data_baseline_arm_1 <- subset(raw_data_baseline_arm_1, record_id!='53')
raw_data_hhci_visit_info_arm_1 <- subset(raw_data_hhci_visit_info_arm_1, record_id!='53')
raw_data_hhci_info_arm_1 <- subset(raw_data_hhci_info_arm_1, record_id!='53')

#Drop record 386

raw_data_baseline_arm_1 <- subset(raw_data_baseline_arm_1, record_id!='386')
raw_data_hhci_visit_info_arm_1 <- subset(raw_data_hhci_visit_info_arm_1, record_id!='386')
raw_data_hhci_info_arm_1 <- subset(raw_data_hhci_info_arm_1, record_id!='386')

#Drop record 352

raw_data_baseline_arm_1 <- subset(raw_data_baseline_arm_1, record_id!='352')
raw_data_hhci_visit_info_arm_1 <- subset(raw_data_hhci_visit_info_arm_1, record_id!='352')
raw_data_hhci_info_arm_1 <- subset(raw_data_hhci_info_arm_1, record_id!='352')


#Drop record 286
raw_data_baseline_arm_1 <- subset(raw_data_baseline_arm_1, record_id!='286')
raw_data_hhci_visit_info_arm_1 <- subset(raw_data_hhci_visit_info_arm_1, record_id!='286')
raw_data_hhci_info_arm_1 <- subset(raw_data_hhci_info_arm_1, record_id!='286')

#Drop 671
raw_data_baseline_arm_1 <- subset(raw_data_baseline_arm_1, record_id!='671')
raw_data_hhci_visit_info_arm_1 <- subset(raw_data_hhci_visit_info_arm_1, record_id!='671')
raw_data_hhci_info_arm_1 <- subset(raw_data_hhci_info_arm_1, record_id!='671')


#Temporary drop record 1007 
raw_data_baseline_arm_1 <- subset(raw_data_baseline_arm_1, record_id!='1007')
raw_data_hhci_visit_info_arm_1 <- subset(raw_data_hhci_visit_info_arm_1, record_id!='1007')
raw_data_hhci_info_arm_1 <- subset(raw_data_hhci_info_arm_1, record_id!='1007')


#Missing Records
raw_data_hhci_visit_info_arm_1 <- subset(raw_data_hhci_visit_info_arm_1, record_id!='1027')
raw_data_hhci_info_arm_1 <- subset(raw_data_hhci_info_arm_1, record_id!='1027')

raw_data_hhci_visit_info_arm_1 <- subset(raw_data_hhci_visit_info_arm_1, record_id!='1046')
raw_data_hhci_info_arm_1 <- subset(raw_data_hhci_info_arm_1, record_id!='1046')

raw_data_hhci_visit_info_arm_1 <- subset(raw_data_hhci_visit_info_arm_1, record_id!='1024')
raw_data_hhci_info_arm_1 <- subset(raw_data_hhci_info_arm_1, record_id!='1024')

raw_data_hhci_visit_info_arm_1 <- subset(raw_data_hhci_visit_info_arm_1, record_id!='1030')
raw_data_hhci_info_arm_1 <- subset(raw_data_hhci_info_arm_1, record_id!='1030')

raw_data_hhci_visit_info_arm_1 <- subset(raw_data_hhci_visit_info_arm_1, record_id!='1058')
raw_data_hhci_info_arm_1 <- subset(raw_data_hhci_info_arm_1, record_id!='1058')

raw_data_hhci_visit_info_arm_1 <- subset(raw_data_hhci_visit_info_arm_1, record_id!='1039')
raw_data_hhci_info_arm_1 <- subset(raw_data_hhci_info_arm_1, record_id!='1039')

raw_data_hhci_visit_info_arm_1 <- subset(raw_data_hhci_visit_info_arm_1, record_id!='1040')
raw_data_hhci_info_arm_1 <- subset(raw_data_hhci_info_arm_1, record_id!='1040')

raw_data_hhci_visit_info_arm_1 <- subset(raw_data_hhci_visit_info_arm_1, record_id!='1067')
raw_data_hhci_info_arm_1 <- subset(raw_data_hhci_info_arm_1, record_id!='1067')

raw_data_hhci_visit_info_arm_1 <- subset(raw_data_hhci_visit_info_arm_1, record_id!='1028')
raw_data_hhci_info_arm_1 <- subset(raw_data_hhci_info_arm_1, record_id!='1028')

raw_data_hhci_visit_info_arm_1 <- subset(raw_data_hhci_visit_info_arm_1, record_id!='1038')
raw_data_hhci_info_arm_1 <- subset(raw_data_hhci_info_arm_1, record_id!='1038')

raw_data_hhci_visit_info_arm_1 <- subset(raw_data_hhci_visit_info_arm_1, record_id!='1081')
raw_data_hhci_info_arm_1 <- subset(raw_data_hhci_info_arm_1, record_id!='1081')









#Drop records that went missing on REDCap


#Adding additional calculated variables
raw_data_baseline_arm_1 <- raw_data_baseline_arm_1%>% mutate(tbip_sc_days_since_init = difftime(today, as.POSIXct(as.Date(tbip_sc_ini_date, format = '%Y-%m-%d')), units = 'days')) %>% relocate(tbip_sc_days_since_init, .after = 'tbip_sc_ini_days_calc')


#raw_data_hhci_info_arm_1$hhc_

#raw_data_hhci_info_arm_1 <- raw_data_hhci_info_arm_1 %>% mutate(hhc_days_since_referral = difftime(today, as.POSIXct(as.Date(hhc_sc_date_cons, format = '%Y-%m-%d')), units = 'days')) %>% relocate(hhc_days_since_referral, .after = 'hhc_sc_date_cons')
#raw_data_hhci_info_arm_1 <- raw_data_hhci_info_arm_1 %>% mutate(hhc_pc_days_to_present = difftime(as.POSIXct(as.Date(hhc_pc_presentation_date, format = '%Y-%m-%d')), as.POSIXct(as.Date(hhc_sc_date_cons, format = '%Y-%m-%d')), units = 'days')) %>% relocate(hhc_pc_days_to_present, .after = 'hhc_sc_date_cons')
#raw_data_hhci_info_arm_1 <- raw_data_hhci_info_arm_1 %>% mutate(hhc_pt_days_to_present = difftime(as.POSIXct(as.Date(hhc_pt_return_date, format = '%Y-%m-%d')), as.POSIXct(as.Date(hhc_sc_date_cons, format = '%Y-%m-%d')), units = 'days')) %>% relocate(hhc_pt_days_to_present, .after = 'hhc_sc_date_cons')

raw_data_hhci_info_arm_1 <- raw_data_hhci_info_arm_1 %>% mutate(hhc_days_since_referral = difftime(today, as.POSIXct(as.Date(hhc_sc_date_cons, format = '%Y-%m-%d')), units = 'days')) %>% relocate(hhc_days_since_referral, .after = 'hhc_sc_date_cons')
raw_data_hhci_info_arm_1 <- raw_data_hhci_info_arm_1 %>% mutate(hhc_pc_days_to_present = difftime(as.POSIXct(as.Date(hhc_pc_presentation_date, format = '%Y-%m-%d')), as.POSIXct(as.Date(hhc_sc_date_cons, format = '%Y-%m-%d')), units = 'days')) %>% relocate(hhc_pc_days_to_present, .after = 'hhc_sc_date_cons')
raw_data_hhci_info_arm_1 <- raw_data_hhci_info_arm_1 %>% mutate(hhc_pt_days_to_present = difftime(as.POSIXct(as.Date(hhc_pt_return_date, format = '%Y-%m-%d')), as.POSIXct(as.Date(hhc_sc_date_cons, format = '%Y-%m-%d')), units = 'days')) %>% relocate(hhc_pt_days_to_present, .after = 'hhc_pc_days_to_present')
raw_data_hhci_info_arm_1 <- raw_data_hhci_info_arm_1 %>% mutate(hhc_pt_days_to_present = difftime(as.POSIXct(as.Date(hhc_pt_return_date, format = '%Y-%m-%d')), as.POSIXct(as.Date(hhc_sc_date_cons, format = '%Y-%m-%d')), units = 'days')) %>% relocate(hhc_pt_days_to_present, .after = 'hhc_sc_date_cons')

#Code missing information as NI
levels(raw_data_hhci_info_arm_1$hhc_sc_clinic_visit) <- c('No', 'Yes', 'NI')
raw_data_hhci_info_arm_1$hhc_sc_clinic_visit[!is.na(raw_data_hhci_info_arm_1$hhc_sc_verbal_consent)] <- 'NI'
levels(raw_data_hhci_info_arm_1$hhc_sc_provide_sputum) <- c('No', 'Yes', 'NI')
raw_data_hhci_info_arm_1$hhc_sc_provide_sputum[!is.na(raw_data_hhci_info_arm_1$hhc_sc_verbal_consent)] <- 'NI'

#raw_data_baseline_arm_1$tbip_sc_eligible[raw_data_baseline_arm_1$record_id == '145'] <- 'Proceed'

raw_data_hhci_info_arm_1 <- raw_data_hhci_info_arm_1 %>% mutate(hhc_days_since_referral = difftime(today, as.POSIXct(as.Date(hhc_sc_date_cons, format = '%Y-%m-%d')), units = 'days')) %>% relocate(hhc_days_since_referral, .after = 'hhc_sc_date_cons')
raw_data_hhci_info_arm_1 <- raw_data_hhci_info_arm_1 %>% mutate(hhc_pc_days_to_present = difftime(as.POSIXct(as.Date(hhc_pc_presentation_date, format = '%Y-%m-%d')), as.POSIXct(as.Date(hhc_sc_date_cons, format = '%Y-%m-%d')), units = 'days')) %>% relocate(hhc_pc_days_to_present, .after = 'hhc_sc_date_cons')
raw_data_hhci_info_arm_1 <- raw_data_hhci_info_arm_1 %>% mutate(hhc_pt_days_to_present = difftime(as.POSIXct(as.Date(hhc_pt_return_date, format = '%Y-%m-%d')), as.POSIXct(as.Date(hhc_sc_date_cons, format = '%Y-%m-%d')), units = 'days')) %>% relocate(hhc_pt_days_to_present, .after = 'hhc_pc_days_to_present')
raw_data_hhci_info_arm_1 <- raw_data_hhci_info_arm_1 %>% mutate(hhc_pt_days_to_present = difftime(as.POSIXct(as.Date(hhc_pt_return_date, format = '%Y-%m-%d')), as.POSIXct(as.Date(hhc_sc_date_cons, format = '%Y-%m-%d')), units = 'days')) %>% relocate(hhc_pt_days_to_present, .after = 'hhc_sc_date_cons')

ip_enrollment <- raw_data_baseline_arm_1[c("record_id", "tbip_sc_date")]

raw_data_hhci_info_arm_1 <- left_join(raw_data_hhci_info_arm_1, ip_enrollment) %>% relocate(tbip_sc_date, .after = record_id)

raw_data_hhci_info_arm_1 <- raw_data_hhci_info_arm_1 %>% mutate(hhc_collection_point = case_when(as.Date(tbip_sc_date) == as.Date(hhcl_date) ~ "Clinic", as.Date(tbip_sc_date) != as.Date(hhcl_date) ~ "HH")) %>% relocate(hhc_collection_point, .after = tbip_sc_date)

#Visit 1
visit_attempt <- as.data.frame.table(table(raw_data_hhci_info_arm_1$record_id, raw_data_hhci_info_arm_1$hhc_sc_visit_attempt___1))

temp_1 <- subset(visit_attempt, Var2=="Checked")
temp_1 <- temp_1[c(1,3)]
names(temp_1) <- c("record_id", "hhc_members_visited_1")

temp_2 <- subset(visit_attempt, Var2=="Unchecked")
temp_2 <- temp_2[c(1,3)]
names(temp_2) <- c("record_id", "hhc_members_not_visited_1")

visit_attempt <- left_join(temp_1, temp_2, by="record_id")

visit_attempt$record_id <- as.numeric(as.character(visit_attempt$record_id))

raw_data_baseline_arm_1 <- left_join(raw_data_baseline_arm_1, visit_attempt)

#Visit 2
visit_attempt <- as.data.frame.table(table(raw_data_hhci_info_arm_1$record_id, raw_data_hhci_info_arm_1$hhc_sc_visit_attempt___2))

temp_1 <- subset(visit_attempt, Var2=="Checked")
temp_1 <- temp_1[c(1,3)]
names(temp_1) <- c("record_id", "hhc_members_visited_2")

temp_2 <- subset(visit_attempt, Var2=="Unchecked")
temp_2 <- temp_2[c(1,3)]
names(temp_2) <- c("record_id", "hhc_members_not_visited_2")

visit_attempt <- left_join(temp_1, temp_2)

visit_attempt$record_id <- as.numeric(visit_attempt$record_id)

raw_data_baseline_arm_1 <- left_join(raw_data_baseline_arm_1, visit_attempt)


#Visit 3
visit_attempt <- as.data.frame.table(table(raw_data_hhci_info_arm_1$record_id, raw_data_hhci_info_arm_1$hhc_sc_visit_attempt___3))

temp_1 <- subset(visit_attempt, Var2=="Checked")
temp_1 <- temp_1[c(1,3)]
names(temp_1) <- c("record_id", "hhc_members_visited_3")

temp_2 <- subset(visit_attempt, Var2=="Unchecked")
temp_2 <- temp_2[c(1,3)]
names(temp_2) <- c("record_id", "hhc_members_not_visited_3")

visit_attempt <- left_join(temp_1, temp_2)

visit_attempt$record_id <- as.numeric(visit_attempt$record_id)

raw_data_baseline_arm_1 <- left_join(raw_data_baseline_arm_1, visit_attempt)
#raw_data_baseline_arm_1 <- raw_data_baseline_arm_1 %>% 
#  select(record_id, contains("tbip_q1_pc_q")) %>% 
#  rowwise() %>%
#  mutate(tbip_q1_pc_count = sum(!is.na)
         
#raw_data_baseline_arm_1 <- raw_data_baseline_arm_1 %>%       
#  mutate(tbip_q1_pc_count = rowSums(!is.na(contains("tbip_q1_pc_q")))) %>%
#  relocate(.data, after=tbip_q1_pc_q14)

