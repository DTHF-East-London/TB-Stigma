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

#datasets
write.table(raw_data_baseline_arm_1, "Data/Stigma Baseline.csv", sep = ",", row.names = FALSE)

write.table(raw_data_baseline_ex_arm_1, "Data/Stigma Baseline_experienced.csv", sep = ",", row.names = FALSE)

write.table(raw_data_baseline_ni_arm_1, "Data/Stigma Baseline_ni.csv", sep = ",", row.names = FALSE)

write.table(raw_data_follow_up_1_arm_1, "Data/Stigma Follow-up.csv", sep = ",", row.names = FALSE)

write.table(raw_data_hhci_info_arm_1, "Data/Stigma hhci_info.csv", sep = ",", row.names = FALSE)

write.table(raw_data_hhci_visit_info_arm_1, "Data/Stigma HHCI visit info.csv", sep = ",", row.names = FALSE)

write.table(raw_data_follow_up_2_arm_1, "Data/Stigma follow_up_2.csv", sep = ",", row.names = FALSE)
