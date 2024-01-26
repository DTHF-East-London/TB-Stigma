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

#recommended dates
rec_date_2 <- subset(raw_data_hhci_info_arm_1, !is.na(raw_data_hhci_info_arm_1$hhc_sc_attempt_1_rec_date))

rec_date_2 <- rec_date_2[c('record_id', 'redcap_repeat_instance', 'hhc_sc_attempt_1_rec_date')]


rec_date_3 <- subset(raw_data_hhci_info_arm_1, !is.na(raw_data_hhci_info_arm_1$hhc_sc_attempt_1_rec_date_2))

rec_date_3 <- rec_date_3[c('record_id', 'redcap_repeat_instance', 'hhc_sc_attempt_1_rec_date_2')]


#Confirmation date
confirmation_date <- subset(raw_data_hhci_visit_info_arm_1, !is.na(raw_data_hhci_visit_info_arm_1$hhc_sch_hhi_date_visit))

confirmation_date <- confirmation_date[c('record_id', 'redcap_repeat_instance','hhc_sch_hhi_date_visit', 'hhc_sch_hhi_visit_time')]
