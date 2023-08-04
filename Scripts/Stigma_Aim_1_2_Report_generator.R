library(dplyr)
library(xlsx)
source("Scripts/functions.R")

source("Scripts/dataset_generator_1.R")

options(java.parameters = "- Xmx2048m")

raw_data_hhci_info_arm_1$hhc_pt_return_clinic[raw_data_hhci_info_arm_1$record_id=='24'] <- NA
raw_data_hhci_info_arm_1$hhc_pt_return_clinic[raw_data_hhci_info_arm_1$record_id=='122'] <- NA
raw_data_hhci_info_arm_1$hhc_pt_return_clinic[raw_data_hhci_info_arm_1$record_id=='7'] <- NA

#report_date <- as.POSIXct('2023-07-07 00:00:00',tz="Africa/Johannesburg")

#raw_data_baseline_arm_1 <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date <= report_date)

#raw_data_hhci_info_arm_1 <- subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhcl_date <= report_date)

#raw_data_hhci_visit_info_arm_1 <- subset(raw_data_hhci_visit_info_arm_1, raw_data_hhci_visit_info_arm_1$hhc_date_sched <= report_date)

wb <- xlsx::loadWorkbook("Metadata/TB_Stigma_Aim_1_2_template.xlsx")

works_sheets <- xlsx::getSheets(wb)

tmp_sheet <- works_sheets[["Aim 1 - Index Patients"]]

rows <- getRows(tmp_sheet)

cells <- getCells(rows)

today <- format(Sys.time(), "%Y-%m-%d")

filename_new <- paste("Data/TB_Stigma_Aim_1_2", today, ".xlsx")

#Screening
setCellValue(cells[["4.3"]], nrow(raw_data_baseline_arm_1))
setCellValue(cells[["5.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed')))
setCellValue(cells[["6.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_end_ip=='End' | raw_data_baseline_arm_1$tbip_sc_age <18)))
setCellValue(cells[["7.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_below_age=='Yes' | raw_data_baseline_arm_1$tbip_sc_below_age=='No')))
setCellValue(cells[["8.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q17=='No')))
setCellValue(cells[["9.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q8=='No')))
setCellValue(cells[["10.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_language=='No')))
setCellValue(cells[["11.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q13=='No')))
setCellValue(cells[["12.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_existing=='Yes')))


#Enrolment
setCellValue(cells[["17.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed')))
setCellValue(cells[["18.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes')))
setCellValue(cells[["19.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_fw_note=='No')))
setCellValue(cells[["20.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='No')))
setCellValue(cells[["21.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_refuse___1=='I\'m not interested')))
setCellValue(cells[["22.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_refuse___2=='I am enrolled in another study')))
setCellValue(cells[["23.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_refuse___3=='I do not have time')))
setCellValue(cells[["24.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_refuse___4=='I am tired')))
setCellValue(cells[["25.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_refuse___5=='Other')))


#Caregiver eligibility
setCellValue(cells[["30.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_below_age=='Yes' | raw_data_baseline_arm_1$tbip_sc_below_age=='No')))
setCellValue(cells[["31.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_below_age=='No')))
setCellValue(cells[["32.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_below_age=='Yes')))
setCellValue(cells[["33.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_cgiver_permission=='Yes'Dat)))
setCellValue(cells[["34.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_cgiver_permission=='No')))


#Aim 1 - Participant groups Newly initiated & Aim 1- Study visits / Retention

setCellValue(cells[["41.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed')))
setCellValue(cells[["41.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes')))
setCellValue(cells[["41.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes')))
setCellValue(cells[["42.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed')))
setCellValue(cells[["42.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes')))
setCellValue(cells[["42.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes')))
setCellValue(cells[["43.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_cgiver_permission=='Yes')))

#Clinic screening rate
setCellValue(cells[["49.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Empilweni Gompo CHC')))
setCellValue(cells[["50.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Pefferville Clinic')))
setCellValue(cells[["51.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Duncan Village CHC')))
setCellValue(cells[["52.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo C Jabavu Clinic')))
setCellValue(cells[["53.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Chris Hani Clinic')))
setCellValue(cells[["54.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Luyolo NU 9 Clinic')))
setCellValue(cells[["55.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Alphendale Clinic')))
setCellValue(cells[["56.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='John Dube Clinic')))
setCellValue(cells[["57.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Fezeka NU 3 Clinic')))
setCellValue(cells[["58.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo A Ndende Clinic')))
setCellValue(cells[["59.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Ndevana Clinic')))
setCellValue(cells[["60.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Philani NU 1 Clinic')))
setCellValue(cells[["61.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Aspiranza Clinic')))
setCellValue(cells[["62.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Ginsberg Clinic')))
setCellValue(cells[["63.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Zwelitsha Zone 5 Clinic')))
setCellValue(cells[["64.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Masakhane Clinic (Zwelitsha)')))
setCellValue(cells[["65.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo B Jwayi Clinic')))
setCellValue(cells[["66.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='NU 12 Clinic')))

#Clinic eligible rate
setCellValue(cells[["49.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Empilweni Gompo CHC')))
setCellValue(cells[["50.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Pefferville Clinic')))
setCellValue(cells[["51.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Duncan Village CHC')))
setCellValue(cells[["52.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo C Jabavu Clinic')))
setCellValue(cells[["53.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Chris Hani Clinic')))
setCellValue(cells[["54.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Luyolo NU 9 Clinic')))
setCellValue(cells[["55.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Alphendale Clinic')))
setCellValue(cells[["56.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='John Dube Clinic')))
setCellValue(cells[["57.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Fezeka NU 3 Clinic')))
setCellValue(cells[["58.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo A Ndende Clinic')))
setCellValue(cells[["59.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Ndevana Clinic')))
setCellValue(cells[["60.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Philani NU 1 Clinic')))
setCellValue(cells[["61.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Aspiranza Clinic')))
setCellValue(cells[["62.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Ginsberg Clinic')))
setCellValue(cells[["63.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Zwelitsha Zone 5 Clinic')))
setCellValue(cells[["64.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Masakhane Clinic (Zwelitsha)')))
setCellValue(cells[["65.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo B Jwayi Clinic')))
setCellValue(cells[["66.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='NU 12 Clinic')))


#Clinic enrolled rate
setCellValue(cells[["49.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Empilweni Gompo CHC')))
setCellValue(cells[["50.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Pefferville Clinic')))
setCellValue(cells[["51.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Duncan Village CHC')))
setCellValue(cells[["52.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo C Jabavu Clinic')))
setCellValue(cells[["53.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Chris Hani Clinic')))
setCellValue(cells[["54.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Luyolo NU 9 Clinic')))
setCellValue(cells[["55.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Alphendale Clinic')))
setCellValue(cells[["56.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='John Dube Clinic')))
setCellValue(cells[["57.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Fezeka NU 3 Clinic')))
setCellValue(cells[["58.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo A Ndende Clinic')))
setCellValue(cells[["59.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Ndevana Clinic')))
setCellValue(cells[["60.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Philani NU 1 Clinic')))
setCellValue(cells[["61.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Aspiranza Clinic')))
setCellValue(cells[["62.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Ginsberg Clinic')))
setCellValue(cells[["63.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Zwelitsha Zone 5 Clinic')))
setCellValue(cells[["64.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Masakhane Clinic (Zwelitsha)')))
setCellValue(cells[["65.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo B Jwayi Clinic')))
setCellValue(cells[["66.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='NU 12 Clinic')))

##Participant time
setCellValue(cells[["71.3"]], mean(raw_data_baseline_arm_1$tbip_q3_duration, na.rm = TRUE))
setCellValue(cells[["72.3"]], median(raw_data_baseline_arm_1$tbip_q3_duration, na.rm = TRUE))
setCellValue(cells[["73.3"]], names(sort(-table(raw_data_baseline_arm_1$tbip_q3_duration)))[1])

xlsx::saveWorkbook(wb, filename_new)

xlsx::forceFormulaRefresh(filename_new)

wb <- xlsx::loadWorkbook(filename_new)

works_sheets <- xlsx::getSheets(wb)

tmp_sheet <- works_sheets[["Aim 1 - Newly Initiated IPs"]]

rows <- getRows(tmp_sheet)

cells <- getCells(rows)

#Screening
setCellValue(cells[["4.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14)))
setCellValue(cells[["5.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed')))
setCellValue(cells[["6.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & (raw_data_baseline_arm_1$tbip_sc_end_ip=='End' | raw_data_baseline_arm_1$tbip_sc_age <18))))
setCellValue(cells[["7.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & (raw_data_baseline_arm_1$tbip_sc_below_age=='Yes' | raw_data_baseline_arm_1$tbip_sc_below_age=='No'))))
setCellValue(cells[["8.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_q17=='No')))
setCellValue(cells[["9.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_q8=='No')))
setCellValue(cells[["10.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_language=='No')))
setCellValue(cells[["11.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_q13=='No')))
setCellValue(cells[["12.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_existing=='Yes')))


#Enrolment
setCellValue(cells[["17.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed')))
setCellValue(cells[["18.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes')))
setCellValue(cells[["19.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_fw_note=='No')))
setCellValue(cells[["20.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='No')))
setCellValue(cells[["21.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_refuse___1=='I\'m not interested')))
setCellValue(cells[["22.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_refuse___2=='I am enrolled in another study')))
setCellValue(cells[["23.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_refuse___3=='I do not have time')))
setCellValue(cells[["24.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_refuse___4=='I am tired')))
setCellValue(cells[["25.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_refuse___5=='Other')))


#Aim 1 - Participant groups Newly initiated & Aim 1- Study visits / Retention
setCellValue(cells[["31.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed')))
setCellValue(cells[["31.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes')))
setCellValue(cells[["31.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$index_questionnaire_3_complete=='Complete')))



#setCellValue(cells[["28.7"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed')))
#setCellValue(cells[["28.8"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes')))
#setCellValue(cells[["28.10"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes')))
#setCellValue(cells[["28.11"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_cgiver_permission=='Yes')))


#Clinic screening rate
setCellValue(cells[["36.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_q5=='Empilweni Gompo CHC')))
setCellValue(cells[["37.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_q5=='Pefferville Clinic')))
setCellValue(cells[["38.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_q5=='Duncan Village CHC')))
setCellValue(cells[["39.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo C Jabavu Clinic')))
setCellValue(cells[["40.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_q5=='Chris Hani Clinic')))
setCellValue(cells[["41.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_q5=='Luyolo NU 9 Clinic')))
setCellValue(cells[["42.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_q5=='Alphendale Clinic')))
setCellValue(cells[["43.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_q5=='John Dube Clinic')))
setCellValue(cells[["44.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_q5=='Fezeka NU 3 Clinic')))
setCellValue(cells[["45.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo A Ndende Clinic')))
setCellValue(cells[["46.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_q5=='Ndevana Clinic')))
setCellValue(cells[["47.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_q5=='Philani NU 1 Clinic')))
setCellValue(cells[["48.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_q5=='Aspiranza Clinic')))
setCellValue(cells[["49.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_q5=='Ginsberg Clinic')))
setCellValue(cells[["50.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_q5=='Zwelitsha Zone 5 Clinic')))
setCellValue(cells[["51.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_q5=='Masakhane Clinic (Zwelitsha)')))
setCellValue(cells[["52.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo B Jwayi Clinic')))
setCellValue(cells[["53.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_q5=='NU 12 Clinic')))

#Clinic eligible rate
setCellValue(cells[["36.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Empilweni Gompo CHC')))
setCellValue(cells[["37.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Pefferville Clinic')))
setCellValue(cells[["38.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Duncan Village CHC')))
setCellValue(cells[["39.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo C Jabavu Clinic')))
setCellValue(cells[["40.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Chris Hani Clinic')))
setCellValue(cells[["41.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Luyolo NU 9 Clinic')))
setCellValue(cells[["42.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Alphendale Clinic')))
setCellValue(cells[["43.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='John Dube Clinic')))
setCellValue(cells[["44.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Fezeka NU 3 Clinic')))
setCellValue(cells[["45.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo A Ndende Clinic')))
setCellValue(cells[["46.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Ndevana Clinic')))
setCellValue(cells[["47.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Philani NU 1 Clinic')))
setCellValue(cells[["48.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Aspiranza Clinic')))
setCellValue(cells[["49.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Ginsberg Clinic')))
setCellValue(cells[["50.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Zwelitsha Zone 5 Clinic')))
setCellValue(cells[["51.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Masakhane Clinic (Zwelitsha)')))
setCellValue(cells[["52.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo B Jwayi Clinic')))
setCellValue(cells[["53.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='NU 12 Clinic')))


#Clinic enrolled rate
setCellValue(cells[["36.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Empilweni Gompo CHC')))
setCellValue(cells[["37.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Pefferville Clinic')))
setCellValue(cells[["38.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Duncan Village CHC')))
setCellValue(cells[["39.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo C Jabavu Clinic')))
setCellValue(cells[["40.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Chris Hani Clinic')))
setCellValue(cells[["41.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Luyolo NU 9 Clinic')))
setCellValue(cells[["42.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Alphendale Clinic')))
setCellValue(cells[["43.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='John Dube Clinic')))
setCellValue(cells[["44.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Fezeka NU 3 Clinic')))
setCellValue(cells[["45.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo A Ndende Clinic')))
setCellValue(cells[["46.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Ndevana Clinic')))
setCellValue(cells[["47.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Philani NU 1 Clinic')))
setCellValue(cells[["48.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Aspiranza Clinic')))
setCellValue(cells[["49.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Ginsberg Clinic')))
setCellValue(cells[["50.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Zwelitsha Zone 5 Clinic')))
setCellValue(cells[["51.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Masakhane Clinic (Zwelitsha)')))
setCellValue(cells[["52.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo B Jwayi Clinic')))
setCellValue(cells[["53.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='NU 12 Clinic')))

##Participant time
setCellValue(cells[["58.3"]], mean(raw_data_baseline_arm_1$tbip_q3_duration, na.rm = TRUE))
setCellValue(cells[["59.3"]], median(raw_data_baseline_arm_1$tbip_q3_duration, na.rm = TRUE))
setCellValue(cells[["60.3"]], names(sort(-table(raw_data_baseline_arm_1$tbip_q3_duration)))[1])


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
#setCellValue(cells[["7.3"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(hhc_collection_point=='Clinic') %>% distinct(record_id)))
#setCellValue(cells[["8.3"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(hhc_collection_point=='HH') %>% distinct(record_id)))

#Total households visited

setCellValue(cells[["8.3"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(hhc_sc_visit_attempt___1=='Checked' | hhc_sc_visit_attempt___2=='Checked' | hhc_sc_visit_attempt___3=='Checked') %>% distinct(record_id)))
setCellValue(cells[["9.3"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(hhc_sc_visit_attempt___1=='Checked') %>% distinct(record_id)))
setCellValue(cells[["10.3"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(hhc_sc_visit_attempt___2=='Checked') %>% distinct(record_id)))
setCellValue(cells[["11.3"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(hhc_sc_visit_attempt___3=='Checked') %>% distinct(record_id)))

#Pending HH Visits
setCellValue(cells[["13.3"]], nrow(raw_data_baseline_arm_1 %>% dplyr::filter(hhc_members_visited_1==0) %>% distinct(record_id)))

#Households with enrolled HHCs
setCellValue(cells[["16.3"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(hhc_sc_consent_provided=='Yes') %>% distinct(record_id)))
setCellValue(cells[["17.3"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(hhc_sc_consent_provided=='Yes') %>% distinct(record_id)))

#HHCs listed by IPs
setCellValue(cells[["27.3"]], nrow(subset(raw_data_hhci_info_arm_1, !is.na(raw_data_hhci_info_arm_1$hhcl_member_name))))
setCellValue(cells[["28.3"]], nrow(subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhcl_member_age>=18)))
setCellValue(cells[["29.3"]], nrow(subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhcl_member_age<18)))
setCellValue(cells[["30.3"]], nrow(subset(raw_data_hhci_info_arm_1, is.na(raw_data_hhci_info_arm_1$hhcl_member_age))))

setCellValue(cells[["32.3"]], nrow(subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhc_collection_point=='Clinic')))
setCellValue(cells[["33.3"]], nrow(subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhc_collection_point=='HH')))

#Screened for eligibility
setCellValue(cells[["35.3"]], nrow(subset(raw_data_hhci_info_arm_1, !is.na(raw_data_hhci_info_arm_1$hhc_sc_clinic_visit))))

#Not Eligible
setCellValue(cells[["36.3"]], nrow(subset(raw_data_hhci_info_arm_1, hhc_sc_clinic_visit=='Yes' |
                                          as.integer(hhc_sc_age_calc)<18 |
                                            hhc_sc_on_treatment=='Yes' |
                                            hhc_sc_verbal_consent=='No' |
                                            hhc_sc_language=='No' |
                                            (hhc_sc_weight_loss=='No' &
                                            hhc_sc_night_sweat=='No' & 
                                            hhc_sc_coughing=='No' &
                                            hhc_sc_fever=='No'))))
setCellValue(cells[["37.3"]], nrow(subset(raw_data_hhci_info_arm_1, hhc_sc_on_treatment=='Yes')))
setCellValue(cells[["38.3"]], nrow(subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhc_sc_clinic_visit=='Yes')))
setCellValue(cells[["39.3"]], nrow(subset(raw_data_hhci_info_arm_1, 
                                          hhc_sc_weight_loss=='No' & 
                                          hhc_sc_night_sweat=='No' & 
                                          hhc_sc_coughing=='No' & 
                                          hhc_sc_fever=='No')))
setCellValue(cells[["40.3"]], nrow(subset(raw_data_hhci_info_arm_1, hhc_sc_verbal_consent=='No')))
setCellValue(cells[["41.3"]], nrow(subset(raw_data_hhci_info_arm_1, hhc_sc_language=='No')))
setCellValue(cells[["42.3"]], nrow(subset(raw_data_hhci_info_arm_1, as.integer(hhc_sc_age_calc)<18)))

#Eligible 
setCellValue(cells[["43.3"]], nrow(subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhc_sc_cons_dir_3=='Proceed')))
setCellValue(cells[["44.3"]], nrow(subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhc_sc_consent_provided=='Yes')))
setCellValue(cells[["45.3"]], nrow(subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhc_sc_consent_provided=='No')))
setCellValue(cells[["46.3"]], nrow(subset(raw_data_hhci_info_arm_1, raw_data_hhci_info_arm_1$hhc_sc_competent=='No')))

#Outcomes
setCellValue(cells[["52.3"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(hhc_days_since_referral<=30 & hhc_sc_consent_provided=='Yes')))
setCellValue(cells[["53.3"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(hhc_days_since_referral>30 & hhc_sc_consent_provided=='Yes')))
setCellValue(cells[["54.3"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(hhc_sc_consent_provided=='Yes')))

#Extracted within the 30 Day window
setCellValue(cells[["59.6"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(hhc_days_since_referral<=30 & hhc_sc_verbal_consent=='Yes' )))
setCellValue(cells[["60.6"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(hhc_days_since_referral<=30 & hhc_pt_intro=='Proceed' & hhc_pt_return_clinic=='Yes')))
setCellValue(cells[["61.6"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(hhc_days_since_referral<=30 & hhc_pt_intro=='Proceed' & hhc_pt_collect_sputum=='Yes')))
setCellValue(cells[["62.6"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(hhc_days_since_referral<=30 & hhc_pt_intro=='Proceed' & hhc_pt_collect_sputum=='No')))
setCellValue(cells[["63.6"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(hhc_days_since_referral<=30 & tbr_smear_res_1=='Negative')))
setCellValue(cells[["64.6"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(hhc_days_since_referral<=30 & (tbr_smear_res_1=='1+' | tbr_smear_res_1=='2+' | tbr_smear_res_1=='3+'))))
setCellValue(cells[["65.6"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(hhc_days_since_referral<=30 & hhc_pt_collect_sputum=='Yes' & is.na(hhc_pc_testing_outcome))))

setCellValue(cells[["59.12"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(hhc_days_since_referral<=30 & hhc_sc_verbal_consent=='Yes' & is.na(hhc_pt_intro))))

#Self reported only after 30 days
setCellValue(cells[["68.3"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(is.na(hhc_pt_return_clinic) & hhc_days_since_referral>30 & hhc_pc_been_facility=='No')))
#setCellValue(cells[["55.3"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(hhc_sc_verbal_consent=='Yes' & hhc_days_since_referral>30 & is.na(hhc_pt_intro))))
setCellValue(cells[["69.3"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(is.na(hhc_pt_intro) & hhc_days_since_referral>30 & (hhc_pc_been_facility=='Yes, I remember the date' | hhc_pc_been_facility=='Yes, I don\'t remember the date'))))
setCellValue(cells[["70.3"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(is.na(hhc_pt_intro) & hhc_days_since_referral>30 & (hhc_pc_been_facility=='Yes, I remember the date' & hhc_pc_days_to_present<=30))))
setCellValue(cells[["71.3"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(is.na(hhc_pt_intro) & hhc_days_since_referral>30 & (hhc_pc_been_facility=='Yes, I remember the date' & hhc_pc_days_to_present>30))))
setCellValue(cells[["72.3"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(is.na(hhc_pt_intro) & hhc_days_since_referral>30 & hhc_pc_been_facility=='Yes, I don\'t remember the date')))
setCellValue(cells[["73.3"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(is.na(hhc_pt_intro) & hhc_days_since_referral>30 & (hhc_pc_been_facility=='Yes, I remember the date' | hhc_pc_been_facility=='Yes, I don\'t remember the date') & hhc_pc_provide_sputum=='No')))
setCellValue(cells[["74.3"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(is.na(hhc_pt_intro) & hhc_days_since_referral>30 & (hhc_pc_been_facility=='Yes, I remember the date' | hhc_pc_been_facility=='Yes, I don\'t remember the date') & hhc_pc_provide_sputum=='Yes')))
setCellValue(cells[["75.3"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(is.na(hhc_pt_intro) & hhc_days_since_referral>30 & hhc_pc_testing_outcome=='Negative')))
setCellValue(cells[["76.3"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(is.na(hhc_pt_intro) & hhc_days_since_referral>30 & hhc_pc_testing_outcome=='Positive')))
setCellValue(cells[["77.3"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(is.na(hhc_pt_intro) & hhc_days_since_referral>30 & hhc_pc_provide_sputum=='Yes' & is.na(hhc_pc_testing_outcome))))
#setCellValue(cells[["76.3"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(is.na(hhc_pt_intro) & hhc_sc_verbal_consent=='Yes' & hhc_days_since_referral>30 & is.na(hhc_pc_been_facility))))

#Extracted Only after 30 days
#setCellValue(cells[["55.6"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(hhc_sc_verbal_consent=='Yes' & hhc_days_since_referral>30)))
#setCellValue(cells[["67.6"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='No')))
setCellValue(cells[["69.6"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes' )))
setCellValue(cells[["70.6"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes' & hhc_pt_days_to_present<=30)))
setCellValue(cells[["71.6"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes' & hhc_pt_days_to_present>30)))
#setCellValue(cells[["71.6"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes')))
setCellValue(cells[["73.6"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes' & hhc_pt_collect_sputum=='No')))
setCellValue(cells[["74.6"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes' & hhc_pt_collect_sputum=='Yes')))
setCellValue(cells[["75.6"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & tbr_smear_res_1=='Negative')))
setCellValue(cells[["76.6"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & (tbr_smear_res_1=='1+' | tbr_smear_res_1=='2+' | tbr_smear_res_1=='3+'))))
setCellValue(cells[["77.6"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_collect_sputum=='Yes'& is.na(tbr_smear_res_1))))
#setCellValue(cells[["76.6"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & hhc_sc_verbal_consent=='Yes' & hhc_days_since_referral>30 & is.na(hhc_pt_return_clinic))))

#Self report and Extracted Only after 30 days
#setCellValue(cells[["55.6"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(hhc_sc_verbal_consent=='Yes' & hhc_days_since_referral>30)))
#setCellValue(cells[["67.9"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter((hhc_pc_been_facility=='Yes, I remember the date' | hhc_pc_been_facility=='Yes, I don\'t remember the date') & hhc_days_since_referral>30 & hhc_pt_return_clinic=='No')))
setCellValue(cells[["69.9"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter((hhc_pc_been_facility=='Yes, I remember the date' | hhc_pc_been_facility=='Yes, I don\'t remember the date') & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes' )))
setCellValue(cells[["70.9"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter((hhc_pc_been_facility=='Yes, I remember the date' | hhc_pc_been_facility=='Yes, I don\'t remember the date') & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes' & hhc_pt_days_to_present<=30)))
setCellValue(cells[["71.9"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter((hhc_pc_been_facility=='Yes, I remember the date' | hhc_pc_been_facility=='Yes, I don\'t remember the date') & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes' & hhc_pt_days_to_present>30)))
#setCellValue(cells[["71.9"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter((hhc_pc_been_facility=='Yes, I remember the date' | hhc_pc_been_facility=='Yes, I don\'t remember the date') & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes')))
setCellValue(cells[["73.9"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter((hhc_pc_been_facility=='Yes, I remember the date' | hhc_pc_been_facility=='Yes, I don\'t remember the date') & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes' & hhc_pt_collect_sputum=='No')))
setCellValue(cells[["74.9"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter((hhc_pc_been_facility=='Yes, I remember the date' | hhc_pc_been_facility=='Yes, I don\'t remember the date') & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes' & hhc_pt_collect_sputum=='Yes')))
setCellValue(cells[["75.9"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter((hhc_pc_been_facility=='Yes, I remember the date' | hhc_pc_been_facility=='Yes, I don\'t remember the date') & hhc_days_since_referral>30 & tbr_smear_res_1=='Negative')))
setCellValue(cells[["76.9"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter((hhc_pc_been_facility=='Yes, I remember the date' | hhc_pc_been_facility=='Yes, I don\'t remember the date') & hhc_days_since_referral>30 & (tbr_smear_res_1=='1+' | tbr_smear_res_1=='2+' | tbr_smear_res_1=='3+'))))
setCellValue(cells[["77.9"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter((hhc_pc_been_facility=='Yes, I remember the date' | hhc_pc_been_facility=='Yes, I don\'t remember the date') & hhc_days_since_referral>30 & hhc_pt_collect_sputum=='Yes'& is.na(tbr_smear_res_1))))

setCellValue(cells[["78.9"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(hhc_days_since_referral>30 & ((hhc_pc_been_facility=='Yes, I remember the date' | hhc_pc_been_facility=='Yes, I don\'t remember the date') & hhc_pt_return_clinic=='No') | (hhc_pc_been_facility=='No' & hhc_pt_return_clinic=='Yes'))))
setCellValue(cells[["79.12"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & is.na(hhc_pt_return_clinic) & hhc_days_since_referral>30)))

xlsx::forceFormulaRefresh(filename_new)
xlsx::saveWorkbook(wb, filename_new)
xlsx::saveWorkbook(wb, "Data\TBStigmaWeeklyReport.xlsx")

print("Outcome - End")

gc()