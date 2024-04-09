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
library(furniture)
library(dplyr)
library(ggplot2)


enrol_date <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes')

enrol_date <- enrol_date[c('record_id', 'tbip_sc_date', 'tbip_sc_consent_date')]

tbip_sc_date <- as.Date(enrol_date$tbip_sc_date)

tbip_sc_consent_date <- as.Date(enrol_date$tbip_sc_consent_date)

enrol_date$date_diff <- difftime(tbip_sc_date, tbip_sc_consent_date, units = "days")

enrol_date <- relocate(enrol_date, date_diff, .after = tbip_sc_consent_date)


#########################################################################################################
