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

wb1 <- xlsx::loadWorkbook("Metadata/HHCI Weekly Enrolment Template.xlsx")

works_sheets <- xlsx::getSheets(wb1)

tmp_sheet <- works_sheets[["Cumulative"]]

rows <- getRows(tmp_sheet)

cells <- getCells(rows)

today <- format(Sys.time(), "%Y-%m-%d")

filename_new <- paste("Data/Weekly Enrolment Chart",today,".xlsx")


setCellValue(cells[["8.4"]], nrow(subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhc_sc_date >= Sys.Date() - 7 &
                                           (raw_data_hhci_info_arm_1$hhc_sc_intro=='Proceed'))))

setCellValue(cells[["10.4"]], nrow(subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhc_sc_date >= Sys.Date() - 7 &
                                            (raw_data_hhci_info_arm_1$hhc_sc_attempt_1_present=='Yes' |
                                               raw_data_hhci_info_arm_1$hhc_sc_attempt_2_present=='Yes' |
                                               raw_data_hhci_info_arm_1$hhc_sc_attempt_3_present=='Yes'))))

setCellValue(cells[["11.4"]], nrow(subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhc_sc_date >= Sys.Date() - 7 &
                                            (raw_data_hhci_info_arm_1$hhc_sc_verbal_consent=='No'))))

setCellValue(cells[["12.4"]], nrow(subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhc_sc_date >= Sys.Date() - 7 &
                                            (!is.na(raw_data_hhci_info_arm_1$hhc_sc_dob)))))

setCellValue(cells[["13.4"]], nrow(subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhc_sc_date >= Sys.Date() - 7 &
                                            (raw_data_hhci_info_arm_1$hhc_sc_age_calc > 17) &
                                               (raw_data_hhci_info_arm_1$hhc_sc_language=='Yes') &
                                               (raw_data_hhci_info_arm_1$hhc_sc_on_treatment=='No'))))

setCellValue(cells[["14.4"]], nrow(subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhc_sc_date >= Sys.Date() - 7 &
                                            (raw_data_hhci_info_arm_1$hhc_sc_age < 17 |
                                               raw_data_hhci_info_arm_1$hhc_sc_language=='No' |
                                               raw_data_hhci_info_arm_1$hhc_sc_on_treatment=='Yes'))))


setCellValue(cells[["15.4"]], nrow(subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhc_sc_date >= Sys.Date() - 7 &
                                            (raw_data_hhci_info_arm_1$hhc_sc_competent=='No'))))

setCellValue(cells[["16.4"]], nrow(subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhc_sc_date >= Sys.Date() - 7 &
                                            (raw_data_hhci_info_arm_1$hhc_sc_consent_provided=='No'))))


setCellValue(cells[["17.4"]], nrow(subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhc_sc_date >= Sys.Date() - 7 &
                                            (raw_data_hhci_info_arm_1$hhc_sc_existing=='Yes'))))


setCellValue(cells[["18.4"]], nrow(subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhc_sc_date >= Sys.Date() - 7 &
                                            (raw_data_hhci_info_arm_1$hhc_sc_consent_provided=='Yes'))))



xlsx::saveWorkbook(wb1, filename_new)

