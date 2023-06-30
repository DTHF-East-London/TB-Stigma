library(dplyr)
library(xlsx)
source("Scripts/functions.R")

source("Scripts/dataset_generator_1.R")

options(java.parameters = "- Xmx2048m")

#report_date <- as.POSIXct('2023-06-24 00:00:00',tz="Africa/Johannesburg")

#raw_data_baseline_arm_1 <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date < report_date)

#raw_data_hhci_info_arm_1 <- subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhcl_date < report_date)

#raw_data_hhci_visit_info_arm_1 <- subset(raw_data_hhci_visit_info_arm_1, raw_data_hhci_visit_info_arm_1$hhc_date_sched < report_date)

wb <- xlsx::loadWorkbook("Metadata/TB_Stigma_Aim_1_2_template.xlsx")

works_sheets <- xlsx::getSheets(wb)

tmp_sheet <- works_sheets[["Aim 1 - Index Patients"]]

rows <- getRows(tmp_sheet)

cells <- getCells(rows)

today <- format(Sys.time(), "%Y-%m-%d")

filename_new <- paste("Data/TB_Stigma_Aim_1_2", today, ".xlsx")

#Screening
setCellValue(cells[["4.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed')))
setCellValue(cells[["5.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_end_ip=='End')))
setCellValue(cells[["6.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_below_age=='No')))
setCellValue(cells[["7.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q17=='No')))
setCellValue(cells[["8.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q8=='No')))
setCellValue(cells[["9.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_language=='No')))
setCellValue(cells[["10.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q13=='No')))
setCellValue(cells[["11.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_existing=='Yes')))
setCellValue(cells[["12.3"]], nrow(raw_data_baseline_arm_1))

#Enrolment
setCellValue(cells[["17.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes')))
setCellValue(cells[["18.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='No')))
setCellValue(cells[["19.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_refuse___1=='I\'m not interested')))
setCellValue(cells[["20.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_refuse___2=='I am enrolled in another study')))
setCellValue(cells[["21.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_refuse___3=='I do not have time')))
setCellValue(cells[["22.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_refuse___4=='I am tired')))
setCellValue(cells[["23.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_refuse___5=='Other')))


#Caregiver eligibility
setCellValue(cells[["28.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_below_age=='No')))
setCellValue(cells[["29.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_below_age=='Yes')))
setCellValue(cells[["30.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_cgiver_permission=='Yes')))
setCellValue(cells[["31.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_cgiver_permission=='No')))
setCellValue(cells[["32.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_below_age=='Yes' & raw_data_baseline_arm_1$tbip_sc_below_age=='No')))


#Aim 1 - Participant groups Newly initiated & Aim 1- Study visits / Retention

setCellValue(cells[["39.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed')))
setCellValue(cells[["39.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes')))
setCellValue(cells[["39.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes')))
#setCellValue(cells[["39.8"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & (raw_data_baseline_arm_1$index_questionnaire_3_complete=='Complete' | raw_data_baseline_arm_1$index_questionnaire_3_complete=='Unverified'))))
#setCellValue(cells[["39.9"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & difftime(Sys.time(), raw_data_baseline_arm_1$tbip_sc_ini_date, unit ='days')>=60)))
#setCellValue(cells[["39.10"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & difftime(Sys.time(), raw_data_baseline_arm_1$tbip_sc_ini_date, unit ='days')>=60)))
#setCellValue(cells[["39.12"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_cgiver_permission=='Yes')))
#setCellValue(cells[["39.13"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_cgiver_permission=='Yes')))

#Clinic screening rate
setCellValue(cells[["47.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Empilweni Gompo CHC')))
setCellValue(cells[["48.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Pefferville Clinic')))
setCellValue(cells[["49.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Duncan Village CHC')))
setCellValue(cells[["50.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo C Jabavu Clinic')))
setCellValue(cells[["51.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Chris Hani Clinic')))
setCellValue(cells[["52.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Luyolo NU 9 Clinic')))
setCellValue(cells[["53.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Alphendale Clinic')))
setCellValue(cells[["54.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='John Dube Clinic')))
setCellValue(cells[["55.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Fezeka NU 3 Clinic')))
setCellValue(cells[["56.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo A Ndende Clinic')))
setCellValue(cells[["57.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Ndevana Clinic')))
setCellValue(cells[["58.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Philani NU 1 Clinic')))
setCellValue(cells[["59.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Aspiranza Clinic')))
setCellValue(cells[["60.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Ginsberg Clinic')))
setCellValue(cells[["61.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Zwelitsha Zone 5 Clinic')))
setCellValue(cells[["62.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Masakhane Clinic (Zwelitsha)')))
setCellValue(cells[["63.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo B Jwayi Clinic')))
setCellValue(cells[["64.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='NU 12 Clinic')))

#Clinic eligible rate
setCellValue(cells[["47.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Empilweni Gompo CHC')))
setCellValue(cells[["48.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Pefferville Clinic')))
setCellValue(cells[["49.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Duncan Village CHC')))
setCellValue(cells[["50.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo C Jabavu Clinic')))
setCellValue(cells[["51.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Chris Hani Clinic')))
setCellValue(cells[["52.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Luyolo NU 9 Clinic')))
setCellValue(cells[["53.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Alphendale Clinic')))
setCellValue(cells[["54.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='John Dube Clinic')))
setCellValue(cells[["55.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Fezeka NU 3 Clinic')))
setCellValue(cells[["56.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo A Ndende Clinic')))
setCellValue(cells[["57.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Ndevana Clinic')))
setCellValue(cells[["58.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Philani NU 1 Clinic')))
setCellValue(cells[["59.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Aspiranza Clinic')))
setCellValue(cells[["60.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Ginsberg Clinic')))
setCellValue(cells[["61.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Zwelitsha Zone 5 Clinic')))
setCellValue(cells[["62.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Masakhane Clinic (Zwelitsha)')))
setCellValue(cells[["63.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo B Jwayi Clinic')))
setCellValue(cells[["64.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='NU 12 Clinic')))


#Clinic enrolled rate
setCellValue(cells[["47.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Empilweni Gompo CHC')))
setCellValue(cells[["48.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Pefferville Clinic')))
setCellValue(cells[["49.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Duncan Village CHC')))
setCellValue(cells[["50.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo C Jabavu Clinic')))
setCellValue(cells[["51.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Chris Hani Clinic')))
setCellValue(cells[["52.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Luyolo NU 9 Clinic')))
setCellValue(cells[["53.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Alphendale Clinic')))
setCellValue(cells[["54.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='John Dube Clinic')))
setCellValue(cells[["55.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Fezeka NU 3 Clinic')))
setCellValue(cells[["56.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo A Ndende Clinic')))
setCellValue(cells[["57.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Ndevana Clinic')))
setCellValue(cells[["58.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Philani NU 1 Clinic')))
setCellValue(cells[["59.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Aspiranza Clinic')))
setCellValue(cells[["60.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Ginsberg Clinic')))
setCellValue(cells[["61.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Zwelitsha Zone 5 Clinic')))
setCellValue(cells[["62.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Masakhane Clinic (Zwelitsha)')))
setCellValue(cells[["63.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo B Jwayi Clinic')))
setCellValue(cells[["64.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='NU 12 Clinic')))

##Participant time
setCellValue(cells[["69.3"]], mean(raw_data_baseline_arm_1$tbip_q3_duration, na.rm = TRUE))
setCellValue(cells[["70.3"]], median(raw_data_baseline_arm_1$tbip_q3_duration, na.rm = TRUE))
setCellValue(cells[["71.3"]], names(sort(-table(raw_data_baseline_arm_1$tbip_q3_duration)))[1])

xlsx::saveWorkbook(wb, filename_new)

xlsx::forceFormulaRefresh(filename_new)

wb <- xlsx::loadWorkbook(filename_new)

works_sheets <- xlsx::getSheets(wb)

tmp_sheet <- works_sheets[["Aim 1 - Newly Initiated IPs"]]

rows <- getRows(tmp_sheet)

cells <- getCells(rows)

#Screening
setCellValue(cells[["4.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed')))
setCellValue(cells[["5.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_end_ip=='End')))
setCellValue(cells[["6.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_below_age=='Yes' | raw_data_baseline_arm_1$tbip_sc_below_age=='No')))
setCellValue(cells[["7.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_q17=='No')))
setCellValue(cells[["8.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_language=='No')))
setCellValue(cells[["9.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_q14=='No')))
setCellValue(cells[["10.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_existing=='Yes')))
setCellValue(cells[["11.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14)))

#Enrolment
setCellValue(cells[["16.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes')))
setCellValue(cells[["17.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='No')))
setCellValue(cells[["18.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_refuse___1=='I\'m not interested')))
setCellValue(cells[["19.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_refuse___2=='I am enrolled in another study')))
setCellValue(cells[["20.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_refuse___3=='I do not have time')))
setCellValue(cells[["21.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_refuse___4=='I am tired')))
setCellValue(cells[["22.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_refuse___5=='Other')))


#Aim 1 - Participant groups Newly initiated & Aim 1- Study visits / Retention
setCellValue(cells[["28.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed')))
setCellValue(cells[["28.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes')))
setCellValue(cells[["28.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & (raw_data_baseline_arm_1$index_questionnaire_3_complete=='1' | raw_data_baseline_arm_1$index_questionnaire_3_complete=='2'))))
#setCellValue(cells[["28.7"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed')))
#setCellValue(cells[["28.8"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes')))
#setCellValue(cells[["28.10"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes')))
#setCellValue(cells[["28.11"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_cgiver_permission=='Yes')))


#Clinic screening rate
setCellValue(cells[["33.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q5=='Empilweni Gompo CHC')))
setCellValue(cells[["34.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q5=='Pefferville Clinic')))
setCellValue(cells[["35.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q5=='Duncan Village CHC')))
setCellValue(cells[["36.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q5=='Gompo C Jabavu Clinic')))
setCellValue(cells[["37.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q5=='Chris Hani Clinic')))
setCellValue(cells[["38.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q5=='Luyolo NU 9 Clinic')))
setCellValue(cells[["39.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q5=='Alphendale Clinic')))
setCellValue(cells[["40.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q5=='John Dube Clinic')))
setCellValue(cells[["41.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q5=='Fezeka NU 3 Clinic')))
setCellValue(cells[["42.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q5=='Gompo A Ndende Clinic')))
setCellValue(cells[["43.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q5=='Ndevana Clinic')))
setCellValue(cells[["44.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q5=='Philani NU 1 Clinic')))
setCellValue(cells[["45.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q5=='Aspiranza Clinic')))
setCellValue(cells[["46.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q5=='Ginsberg Clinic')))
setCellValue(cells[["47.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q5=='Zwelitsha Zone 5 Clinic')))
setCellValue(cells[["48.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q5=='Masakhane Clinic (Zwelitsha)')))
setCellValue(cells[["49.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q5=='Gompo B Jwayi Clinic')))
setCellValue(cells[["50.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q5=='NU 12 Clinic')))

#Clinic eligible rate
setCellValue(cells[["33.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Empilweni Gompo CHC')))
setCellValue(cells[["34.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Pefferville Clinic')))
setCellValue(cells[["35.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Duncan Village CHC')))
setCellValue(cells[["36.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo C Jabavu Clinic')))
setCellValue(cells[["37.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Chris Hani Clinic')))
setCellValue(cells[["38.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Luyolo NU 9 Clinic')))
setCellValue(cells[["39.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Alphendale Clinic')))
setCellValue(cells[["40.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='John Dube Clinic')))
setCellValue(cells[["41.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Fezeka NU 3 Clinic')))
setCellValue(cells[["42.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo A Ndende Clinic')))
setCellValue(cells[["43.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Ndevana Clinic')))
setCellValue(cells[["44.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Philani NU 1 Clinic')))
setCellValue(cells[["45.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Aspiranza Clinic')))
setCellValue(cells[["46.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Ginsberg Clinic')))
setCellValue(cells[["47.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Zwelitsha Zone 5 Clinic')))
setCellValue(cells[["48.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Masakhane Clinic (Zwelitsha)')))
setCellValue(cells[["49.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo B Jwayi Clinic')))
setCellValue(cells[["50.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='NU 12 Clinic')))


#Clinic enrolled rate
setCellValue(cells[["33.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Empilweni Gompo CHC')))
setCellValue(cells[["34.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Pefferville Clinic')))
setCellValue(cells[["35.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Duncan Village CHC')))
setCellValue(cells[["36.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo C Jabavu Clinic')))
setCellValue(cells[["37.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Chris Hani Clinic')))
setCellValue(cells[["38.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Luyolo NU 9 Clinic')))
setCellValue(cells[["39.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Alphendale Clinic')))
setCellValue(cells[["40.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='John Dube Clinic')))
setCellValue(cells[["41.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Fezeka NU 3 Clinic')))
setCellValue(cells[["42.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo A Ndende Clinic')))
setCellValue(cells[["43.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Ndevana Clinic')))
setCellValue(cells[["44.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Philani NU 1 Clinic')))
setCellValue(cells[["45.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Aspiranza Clinic')))
setCellValue(cells[["46.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Ginsberg Clinic')))
setCellValue(cells[["47.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Zwelitsha Zone 5 Clinic')))
setCellValue(cells[["48.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Masakhane Clinic (Zwelitsha)')))
setCellValue(cells[["49.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo B Jwayi Clinic')))
setCellValue(cells[["50.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='NU 12 Clinic')))

##Participant time
setCellValue(cells[["55.3"]], mean(raw_data_baseline_arm_1$tbip_q3_duration, na.rm = TRUE))
setCellValue(cells[["56.3"]], median(raw_data_baseline_arm_1$tbip_q3_duration, na.rm = TRUE))
setCellValue(cells[["57.3"]], names(sort(-table(raw_data_baseline_arm_1$tbip_q3_duration)))[1])


xlsx::saveWorkbook(wb, filename_new)

xlsx::forceFormulaRefresh(filename_new)


wb <- xlsx::loadWorkbook(filename_new)

works_sheets <- xlsx::getSheets(wb)

tmp_sheet <- works_sheets[["Aim 1 - Household Contacts"]]

rows <- getRows(tmp_sheet)

cells <- getCells(rows)

#HHCI: Screening & Enrolment

#Households listed by IPs

setCellValue(cells[["6.3"]], nrow(raw_data_hhci_info_arm_1 %>% distinct(record_id)))

#Total households visited

setCellValue(cells[["8.3"]], nrow(raw_data_hhci_info_arm_1 %>% filter(hhc_sc_visit_attempt___1=='First' | hhc_sc_visit_attempt___2=='Second' | hhc_sc_visit_attempt___3=='Third') %>% distinct(record_id)))
setCellValue(cells[["9.3"]], nrow(raw_data_hhci_info_arm_1 %>% filter(hhc_sc_visit_attempt___1=='First') %>% distinct(record_id)))
setCellValue(cells[["10.3"]], nrow(raw_data_hhci_info_arm_1 %>% filter(hhc_sc_visit_attempt___2=='Second') %>% distinct(record_id)))
setCellValue(cells[["11.3"]], nrow(raw_data_hhci_info_arm_1 %>% filter(hhc_sc_visit_attempt___3=='Third') %>% distinct(record_id)))

#Households with enrolled HHCs
setCellValue(cells[["13.3"]], nrow(raw_data_hhci_info_arm_1 %>% filter(hhc_sc_verbal_consent=='Yes') %>% distinct(record_id)))

#HHCs listed by IPs
setCellValue(cells[["16.3"]], nrow(subset(raw_data_hhci_info_arm_1, !is.na(raw_data_hhci_info_arm_1$hhc_sc_age_calc))))
setCellValue(cells[["17.3"]], nrow(subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhc_sc_age_calc>=18)))
setCellValue(cells[["18.3"]], nrow(subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhc_sc_age_calc<18)))

#Screened for eligibility
setCellValue(cells[["20.3"]], nrow(subset(raw_data_hhci_info_arm_1, !is.na(raw_data_hhci_info_arm_1$hhc_sc_clinic_visit))))

#Not Eligible
setCellValue(cells[["21.3"]], nrow(subset(raw_data_hhci_info_arm_1, as.integer(raw_data_baseline_arm_1$hhc_sc_age_calc)>=18 |
                                            (raw_data_baseline_arm_1$hhc_sc_on_treatment=='Yes' &
                                            raw_data_hhci_info_arm_1$hhc_sc_weight_loss=='No' &
                                            raw_data_hhci_info_arm_1$hhc_sc_night_sweat=='No' & 
                                            raw_data_hhci_info_arm_1$hhc_sc_coughing=='No' &
                                            raw_data_hhci_info_arm_1$hhc_sc_fever=='No'))))
setCellValue(cells[["22.3"]], nrow(subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhc_sc_clinic_visit=='Yes')))
setCellValue(cells[["23.3"]], nrow(subset(raw_data_hhci_info_arm_1, 
                                          raw_data_hhci_info_arm_1$hhc_sc_weight_loss=='0' & 
                                            raw_data_hhci_info_arm_1$hhc_sc_night_sweat=='0' & 
                                            raw_data_hhci_info_arm_1$hhc_sc_coughing=='0' & 
                                            raw_data_hhci_info_arm_1$hhc_sc_fever=='0')))
setCellValue(cells[["24.3"]], nrow(subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhc_sc_cons_dir_1_hhm=='Proceed')))

#Eligible 
setCellValue(cells[["25.3"]], nrow(subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhc_sc_cons_dir_3=='Proceed')))
setCellValue(cells[["26.3"]], nrow(subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhc_sc_consent_provided=='Yes')))
setCellValue(cells[["27.3"]], nrow(subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhc_sc_consent_provided=='No')))
setCellValue(cells[["28.3"]], nrow(subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhc_sc_competent=='No')))

raw_data_hhci_info_arm_1 <- raw_data_hhci_info_arm_1 %>% mutate(hhc_days_since_referral = difftime(today, as.POSIXct(as.Date(hhc_sc_date_cons, format = '%Y-%m-%d')), units = 'days')) %>% relocate(hhc_days_since_referral, .after = 'hhc_sc_date_cons')
raw_data_hhci_info_arm_1 <- raw_data_hhci_info_arm_1 %>% mutate(hhc_pc_days_to_present = difftime(as.POSIXct(as.Date(hhc_pc_presentation_date, format = '%Y-%m-%d')), as.POSIXct(as.Date(hhc_sc_date_cons, format = '%Y-%m-%d')), units = 'days')) %>% relocate(hhc_pc_days_to_present, .after = 'hhc_sc_date_cons')
raw_data_hhci_info_arm_1 <- raw_data_hhci_info_arm_1 %>% mutate(hhc_pt_days_to_present = difftime(as.POSIXct(as.Date(hhc_pt_return_date, format = '%Y-%m-%d')), as.POSIXct(as.Date(hhc_sc_date_cons, format = '%Y-%m-%d')), units = 'days')) %>% relocate(hhc_pt_days_to_present, .after = 'hhc_sc_date_cons')


#Outcomes 
#Extracted within the 30 Day window
setCellValue(cells[["35.6"]], nrow(raw_data_hhci_info_arm_1 %>% filter(hhc_days_since_referral<=30 & hhc_sc_verbal_consent=='Yes' )))
setCellValue(cells[["36.6"]], nrow(raw_data_hhci_info_arm_1 %>% filter(hhc_days_since_referral<=30 & hhc_pt_intro=='Proceed' & hhc_pt_return_clinic=='Yes')))
setCellValue(cells[["37.6"]], nrow(raw_data_hhci_info_arm_1 %>% filter(hhc_days_since_referral<=30 & hhc_pt_intro=='Proceed' & hhc_pt_collect_sputum=='Yes')))
setCellValue(cells[["38.6"]], nrow(raw_data_hhci_info_arm_1 %>% filter(hhc_days_since_referral<=30 & hhc_sc_verbal_consent=='Yes' & is.na(hhc_pt_return_clinic))))

#Self reported after 30 days
setCellValue(cells[["40.3"]], nrow(raw_data_hhci_info_arm_1 %>% filter(hhc_sc_verbal_consent=='Yes' & hhc_days_since_referral>30)))
setCellValue(cells[["41.3"]], nrow(raw_data_hhci_info_arm_1 %>% filter(hhc_days_since_referral>30 & (hhc_pc_been_facility=='Yes, I remember the date' | hhc_pc_been_facility=='Yes, I don\'t remember the date'))))
setCellValue(cells[["42.3"]], nrow(raw_data_hhci_info_arm_1 %>% filter(hhc_days_since_referral>30 & (hhc_pc_been_facility=='Yes, I remember the date' & hhc_pc_days_to_present<=30))))
setCellValue(cells[["43.3"]], nrow(raw_data_hhci_info_arm_1 %>% filter(hhc_days_since_referral>30 & (hhc_pc_been_facility=='Yes, I remember the date' & hhc_pc_days_to_present>30))))
setCellValue(cells[["44.3"]], nrow(raw_data_hhci_info_arm_1 %>% filter(hhc_days_since_referral>30 & hhc_pc_been_facility=='Yes, I don\'t remember the date')))
setCellValue(cells[["45.3"]], nrow(raw_data_hhci_info_arm_1 %>% filter(hhc_days_since_referral>30 & (hhc_pc_been_facility=='Yes, I remember the date' | hhc_pc_been_facility=='Yes, I don\'t remember the date') & hhc_pc_provide_sputum=='Yes')))
setCellValue(cells[["46.3"]], nrow(raw_data_hhci_info_arm_1 %>% filter(hhc_days_since_referral>30 & hhc_pc_been_facility=='No')))
setCellValue(cells[["47.3"]], nrow(raw_data_hhci_info_arm_1 %>% filter(hhc_sc_verbal_consent=='Yes' & hhc_days_since_referral>30 & is.na(hhc_pc_been_facility))))

#Extracted after 30 days
setCellValue(cells[["40.6"]], nrow(raw_data_hhci_info_arm_1 %>% filter(hhc_sc_verbal_consent=='Yes' & hhc_days_since_referral>30)))
setCellValue(cells[["41.6"]], nrow(raw_data_hhci_info_arm_1 %>% filter(hhc_days_since_referral>30 & (hhc_pt_return_clinic=='Yes, I remember the date' | hhc_pt_return_clinic=='Yes, I don\'t remember the date'))))
setCellValue(cells[["42.6"]], nrow(raw_data_hhci_info_arm_1 %>% filter(hhc_days_since_referral>30 & (hhc_pt_return_clinic=='Yes, I remember the date') & hhc_pt_days_to_present<=30)))
setCellValue(cells[["43.6"]], nrow(raw_data_hhci_info_arm_1 %>% filter(hhc_days_since_referral>30 & (hhc_pt_return_clinic=='Yes, I remember the date') & hhc_pt_days_to_present>30)))
setCellValue(cells[["44.6"]], nrow(raw_data_hhci_info_arm_1 %>% filter(hhc_days_since_referral>30 & (hhc_pt_return_clinic=='Yes, I don\'t remember the date'))))
setCellValue(cells[["45.6"]], nrow(raw_data_hhci_info_arm_1 %>% filter(hhc_days_since_referral>30 & (hhc_pt_return_clinic=='Yes, I remember the date' | hhc_pt_return_clinic=='Yes, I don\'t remember the date') & hhc_pt_collect_sputum=='Yes')))
setCellValue(cells[["46.6"]], nrow(raw_data_hhci_info_arm_1 %>% filter(hhc_days_since_referral>30 & hhc_pt_return_clinic=='No')))
setCellValue(cells[["47.6"]], nrow(raw_data_hhci_info_arm_1 %>% filter(hhc_sc_verbal_consent=='Yes' & hhc_days_since_referral>30 & is.na(hhc_pt_return_clinic))))

xlsx::forceFormulaRefresh(filename_new)
xlsx::saveWorkbook(wb, filename_new)

print("Outcome - End")

