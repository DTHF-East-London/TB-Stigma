library(dplyr)
library(xlsx)
source("Scripts/functions.R")


#source("Scripts/dataset_generator_1.R")

options(java.parameters = "- Xmx2048m")

#raw_data_hhci_info_arm_1$hhc_pt_return_clinic[raw_data_hhci_info_arm_1$record_id=='24'] <- NA
#raw_data_hhci_info_arm_1$hhc_pt_return_clinic[raw_data_hhci_info_arm_1$record_id=='122'] <- NA
#raw_data_hhci_info_arm_1$hhc_pt_return_clinic[raw_data_hhci_info_arm_1$record_id=='7'] <- NA

raw_data_baseline_ni_arm_1 <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14)
raw_data_baseline_ex_arm_1 <- subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14)

raw_data_hhci_info_ni_arm_1 <- raw_data_hhci_info_arm_1 %>% dplyr::filter(record_id %in% raw_data_baseline_ni_arm_1$record_id)
raw_data_hhci_info_ex_arm_1 <- raw_data_hhci_info_arm_1 %>% dplyr::filter(record_id %in% raw_data_baseline_ex_arm_1$record_id)

raw_data_hhci_visit_info_ni_arm_1 <- raw_data_hhci_visit_info_arm_1 %>% dplyr::filter(record_id %in% raw_data_baseline_ni_arm_1$record_id)
raw_data_hhci_visit_info_ex_arm_1 <- raw_data_hhci_visit_info_arm_1 %>% dplyr::filter(record_id %in% raw_data_baseline_ex_arm_1$record_id)

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

################################################################################
#                                                                              #
#                             All Participants                                 #
#                                                                              #
################################################################################

#Screening
setCellValue(cells[["4.3"]], nrow(raw_data_baseline_arm_1))
setCellValue(cells[["5.3"]], nrow(subset(raw_data_baseline_arm_1, !is.na(raw_data_baseline_arm_1$sc_aim))))
setCellValue(cells[["6.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed')))
setCellValue(cells[["7.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_end_ip=='End' | raw_data_baseline_arm_1$tbip_sc_age <18)))
setCellValue(cells[["8.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_below_age=='Yes' | raw_data_baseline_arm_1$tbip_sc_below_age=='No')))
setCellValue(cells[["9.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q17=='No')))
setCellValue(cells[["10.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q8=='No')))
setCellValue(cells[["11.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_language=='No')))
setCellValue(cells[["12.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q13=='No')))
setCellValue(cells[["13.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_existing=='Yes')))


#Enrolment
setCellValue(cells[["18.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed')))
setCellValue(cells[["19.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes')))
setCellValue(cells[["20.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_fw_note=='No')))
setCellValue(cells[["21.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='No')))
setCellValue(cells[["22.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_refuse___1=='I\'m not interested')))
setCellValue(cells[["23.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_refuse___2=='I am enrolled in another study')))
setCellValue(cells[["24.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_refuse___3=='I do not have time')))
setCellValue(cells[["25.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_refuse___4=='I am tired')))
setCellValue(cells[["26.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_refuse___5=='Other')))


#Caregiver eligibility
setCellValue(cells[["31.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_below_age=='Yes' | raw_data_baseline_arm_1$tbip_sc_below_age=='No')))
setCellValue(cells[["32.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_below_age=='No')))
setCellValue(cells[["33.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_below_age=='Yes')))
setCellValue(cells[["34.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_cgiver_permission=='Yes')))
setCellValue(cells[["35.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_cgiver_permission=='No')))

#Aim 1 - Participant groups Newly initiated & Aim 1- Study visits / Retention

setCellValue(cells[["42.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed')))
setCellValue(cells[["42.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes')))
setCellValue(cells[["42.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes')))
setCellValue(cells[["43.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed')))
setCellValue(cells[["43.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes')))
setCellValue(cells[["43.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes')))
setCellValue(cells[["44.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_cgiver_permission=='Yes')))

#Clinic screening rate
setCellValue(cells[["50.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Empilweni Gompo CHC')))
setCellValue(cells[["51.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Pefferville Clinic')))
setCellValue(cells[["52.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Duncan Village CHC')))
setCellValue(cells[["53.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo C Jabavu Clinic')))
setCellValue(cells[["54.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Chris Hani Clinic')))
setCellValue(cells[["55.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Luyolo NU 9 Clinic')))
setCellValue(cells[["56.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Alphendale Clinic')))
setCellValue(cells[["57.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='John Dube Clinic')))
setCellValue(cells[["58.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Fezeka NU 3 Clinic')))
setCellValue(cells[["59.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo A Ndende Clinic')))
setCellValue(cells[["60.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Ndevana Clinic')))
setCellValue(cells[["61.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Philani NU 1 Clinic')))
setCellValue(cells[["62.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Aspiranza Clinic')))
setCellValue(cells[["63.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Ginsberg Clinic')))
setCellValue(cells[["64.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Zwelitsha Zone 5 Clinic')))
setCellValue(cells[["65.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Masakhane Clinic (Zwelitsha)')))
setCellValue(cells[["66.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo B Jwayi Clinic')))
setCellValue(cells[["67.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_q5=='NU 12 Clinic')))

#Clinic eligible rate
setCellValue(cells[["50.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Empilweni Gompo CHC')))
setCellValue(cells[["51.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Pefferville Clinic')))
setCellValue(cells[["52.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Duncan Village CHC')))
setCellValue(cells[["53.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo C Jabavu Clinic')))
setCellValue(cells[["54.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Chris Hani Clinic')))
setCellValue(cells[["55.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Luyolo NU 9 Clinic')))
setCellValue(cells[["56.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Alphendale Clinic')))
setCellValue(cells[["57.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='John Dube Clinic')))
setCellValue(cells[["58.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Fezeka NU 3 Clinic')))
setCellValue(cells[["59.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo A Ndende Clinic')))
setCellValue(cells[["60.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Ndevana Clinic')))
setCellValue(cells[["61.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Philani NU 1 Clinic')))
setCellValue(cells[["62.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Aspiranza Clinic')))
setCellValue(cells[["63.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Ginsberg Clinic')))
setCellValue(cells[["64.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Zwelitsha Zone 5 Clinic')))
setCellValue(cells[["65.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Masakhane Clinic (Zwelitsha)')))
setCellValue(cells[["66.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo B Jwayi Clinic')))
setCellValue(cells[["67.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='NU 12 Clinic')))


#Clinic enrolled rate
setCellValue(cells[["50.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Empilweni Gompo CHC')))
setCellValue(cells[["51.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Pefferville Clinic')))
setCellValue(cells[["52.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Duncan Village CHC')))
setCellValue(cells[["53.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo C Jabavu Clinic')))
setCellValue(cells[["54.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Chris Hani Clinic')))
setCellValue(cells[["55.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Luyolo NU 9 Clinic')))
setCellValue(cells[["56.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Alphendale Clinic')))
setCellValue(cells[["57.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='John Dube Clinic')))
setCellValue(cells[["58.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Fezeka NU 3 Clinic')))
setCellValue(cells[["59.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo A Ndende Clinic')))
setCellValue(cells[["60.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Ndevana Clinic')))
setCellValue(cells[["61.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Philani NU 1 Clinic')))
setCellValue(cells[["62.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Aspiranza Clinic')))
setCellValue(cells[["63.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Ginsberg Clinic')))
setCellValue(cells[["64.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Zwelitsha Zone 5 Clinic')))
setCellValue(cells[["65.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Masakhane Clinic (Zwelitsha)')))
setCellValue(cells[["66.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo B Jwayi Clinic')))
setCellValue(cells[["67.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='NU 12 Clinic')))

##Participant time
setCellValue(cells[["72.3"]], mean(raw_data_baseline_arm_1$tbip_q3_duration, na.rm = TRUE))
setCellValue(cells[["73.3"]], median(raw_data_baseline_arm_1$tbip_q3_duration, na.rm = TRUE))
setCellValue(cells[["74.3"]], names(sort(-table(raw_data_baseline_arm_1$tbip_q3_duration)))[1])


################################################################################
#                                                                              #
#                             Newly Initiated Participants                    #
#                                                                              #
################################################################################

#Screening
setCellValue(cells[["4.11"]], nrow(raw_data_baseline_ni_arm_1))
setCellValue(cells[["5.11"]], nrow(subset(raw_data_baseline_ni_arm_1, !is.na(raw_data_baseline_ni_arm_1$sc_aim))))
setCellValue(cells[["6.11"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_eligible=='Proceed')))
setCellValue(cells[["7.11"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_end_ip=='End' | raw_data_baseline_ni_arm_1$tbip_sc_age <18)))
setCellValue(cells[["8.11"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_below_age=='Yes' | raw_data_baseline_ni_arm_1$tbip_sc_below_age=='No')))
setCellValue(cells[["9.11"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_q17=='No')))
setCellValue(cells[["10.11"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_q8=='No')))
setCellValue(cells[["11.11"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_language=='No')))
setCellValue(cells[["12.11"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_q13=='No')))
setCellValue(cells[["13.11"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_existing=='Yes')))

#Enrolment
setCellValue(cells[["18.11"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_eligible=='Proceed')))
setCellValue(cells[["19.11"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_consent_part=='Yes')))
setCellValue(cells[["20.11"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_fw_note=='No')))
setCellValue(cells[["21.11"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_consent_part=='No')))
setCellValue(cells[["22.11"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_refuse___1=='I\'m not interested')))
setCellValue(cells[["23.11"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_refuse___2=='I am enrolled in another study')))
setCellValue(cells[["24.11"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_refuse___3=='I do not have time')))
setCellValue(cells[["25.11"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_refuse___4=='I am tired')))
setCellValue(cells[["26.11"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_refuse___5=='Other')))


#Caregiver eligibility
setCellValue(cells[["31.12"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_below_age=='Yes' | raw_data_baseline_ni_arm_1$tbip_sc_below_age=='No')))
setCellValue(cells[["32.12"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_below_age=='No')))
setCellValue(cells[["33.12"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_below_age=='Yes')))
setCellValue(cells[["34.12"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_cgiver_permission=='Yes')))
setCellValue(cells[["35.12"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_cgiver_permission=='No')))

#Aim 1 - Participant groups Newly initiated & Aim 1- Study visits / Retention

setCellValue(cells[["42.11"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_eligible=='Proceed')))
setCellValue(cells[["42.12"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_consent_part=='Yes')))
setCellValue(cells[["42.14"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_consent_part=='Yes')))
setCellValue(cells[["43.11"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_eligible=='Proceed')))
setCellValue(cells[["43.12"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_consent_part=='Yes')))
setCellValue(cells[["43.14"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_consent_part=='Yes')))
setCellValue(cells[["44.14"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_cgiver_permission=='Yes')))

#Clinic screening rate
setCellValue(cells[["50.12"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_q5=='Empilweni Gompo CHC')))
setCellValue(cells[["51.12"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_q5=='Pefferville Clinic')))
setCellValue(cells[["52.12"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_q5=='Duncan Village CHC')))
setCellValue(cells[["53.12"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_q5=='Gompo C Jabavu Clinic')))
setCellValue(cells[["54.12"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_q5=='Chris Hani Clinic')))
setCellValue(cells[["55.12"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_q5=='Luyolo NU 9 Clinic')))
setCellValue(cells[["56.12"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_q5=='Alphendale Clinic')))
setCellValue(cells[["57.12"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_q5=='John Dube Clinic')))
setCellValue(cells[["58.12"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_q5=='Fezeka NU 3 Clinic')))
setCellValue(cells[["59.12"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_q5=='Gompo A Ndende Clinic')))
setCellValue(cells[["60.12"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_q5=='Ndevana Clinic')))
setCellValue(cells[["61.12"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_q5=='Philani NU 1 Clinic')))
setCellValue(cells[["62.12"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_q5=='Aspiranza Clinic')))
setCellValue(cells[["63.12"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_q5=='Ginsberg Clinic')))
setCellValue(cells[["64.12"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_q5=='Zwelitsha Zone 5 Clinic')))
setCellValue(cells[["65.12"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_q5=='Masakhane Clinic (Zwelitsha)')))
setCellValue(cells[["66.12"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_q5=='Gompo B Jwayi Clinic')))
setCellValue(cells[["67.12"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_q5=='NU 12 Clinic')))

#Clinic eligible rate
setCellValue(cells[["50.13"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_ni_arm_1$tbip_sc_q5=='Empilweni Gompo CHC')))
setCellValue(cells[["51.13"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_ni_arm_1$tbip_sc_q5=='Pefferville Clinic')))
setCellValue(cells[["52.13"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_ni_arm_1$tbip_sc_q5=='Duncan Village CHC')))
setCellValue(cells[["53.13"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_ni_arm_1$tbip_sc_q5=='Gompo C Jabavu Clinic')))
setCellValue(cells[["54.13"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_ni_arm_1$tbip_sc_q5=='Chris Hani Clinic')))
setCellValue(cells[["55.13"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_ni_arm_1$tbip_sc_q5=='Luyolo NU 9 Clinic')))
setCellValue(cells[["56.13"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_ni_arm_1$tbip_sc_q5=='Alphendale Clinic')))
setCellValue(cells[["57.13"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_ni_arm_1$tbip_sc_q5=='John Dube Clinic')))
setCellValue(cells[["58.13"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_ni_arm_1$tbip_sc_q5=='Fezeka NU 3 Clinic')))
setCellValue(cells[["59.13"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_ni_arm_1$tbip_sc_q5=='Gompo A Ndende Clinic')))
setCellValue(cells[["60.13"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_ni_arm_1$tbip_sc_q5=='Ndevana Clinic')))
setCellValue(cells[["61.13"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_ni_arm_1$tbip_sc_q5=='Philani NU 1 Clinic')))
setCellValue(cells[["62.13"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_ni_arm_1$tbip_sc_q5=='Aspiranza Clinic')))
setCellValue(cells[["63.13"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_ni_arm_1$tbip_sc_q5=='Ginsberg Clinic')))
setCellValue(cells[["64.13"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_ni_arm_1$tbip_sc_q5=='Zwelitsha Zone 5 Clinic')))
setCellValue(cells[["65.13"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_ni_arm_1$tbip_sc_q5=='Masakhane Clinic (Zwelitsha)')))
setCellValue(cells[["66.13"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_ni_arm_1$tbip_sc_q5=='Gompo B Jwayi Clinic')))
setCellValue(cells[["67.13"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_ni_arm_1$tbip_sc_q5=='NU 12 Clinic')))


#Clinic enrolled rate
setCellValue(cells[["50.14"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_ni_arm_1$tbip_sc_q5=='Empilweni Gompo CHC')))
setCellValue(cells[["51.14"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_ni_arm_1$tbip_sc_q5=='Pefferville Clinic')))
setCellValue(cells[["52.14"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_ni_arm_1$tbip_sc_q5=='Duncan Village CHC')))
setCellValue(cells[["53.14"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_ni_arm_1$tbip_sc_q5=='Gompo C Jabavu Clinic')))
setCellValue(cells[["54.14"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_ni_arm_1$tbip_sc_q5=='Chris Hani Clinic')))
setCellValue(cells[["55.14"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_ni_arm_1$tbip_sc_q5=='Luyolo NU 9 Clinic')))
setCellValue(cells[["56.14"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_ni_arm_1$tbip_sc_q5=='Alphendale Clinic')))
setCellValue(cells[["57.14"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_ni_arm_1$tbip_sc_q5=='John Dube Clinic')))
setCellValue(cells[["58.14"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_ni_arm_1$tbip_sc_q5=='Fezeka NU 3 Clinic')))
setCellValue(cells[["59.14"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_ni_arm_1$tbip_sc_q5=='Gompo A Ndende Clinic')))
setCellValue(cells[["60.14"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_ni_arm_1$tbip_sc_q5=='Ndevana Clinic')))
setCellValue(cells[["61.14"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_ni_arm_1$tbip_sc_q5=='Philani NU 1 Clinic')))
setCellValue(cells[["62.14"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_ni_arm_1$tbip_sc_q5=='Aspiranza Clinic')))
setCellValue(cells[["63.14"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_ni_arm_1$tbip_sc_q5=='Ginsberg Clinic')))
setCellValue(cells[["64.14"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_ni_arm_1$tbip_sc_q5=='Zwelitsha Zone 5 Clinic')))
setCellValue(cells[["65.14"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_ni_arm_1$tbip_sc_q5=='Masakhane Clinic (Zwelitsha)')))
setCellValue(cells[["66.14"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_ni_arm_1$tbip_sc_q5=='Gompo B Jwayi Clinic')))
setCellValue(cells[["67.14"]], nrow(subset(raw_data_baseline_ni_arm_1, raw_data_baseline_ni_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_ni_arm_1$tbip_sc_q5=='NU 12 Clinic')))

##Participant time
setCellValue(cells[["72.11"]], mean(raw_data_baseline_ni_arm_1$tbip_q3_duration, na.rm = TRUE))
setCellValue(cells[["73.11"]], median(raw_data_baseline_ni_arm_1$tbip_q3_duration, na.rm = TRUE))
setCellValue(cells[["74.11"]], names(sort(-table(raw_data_baseline_ni_arm_1$tbip_q3_duration)))[1])


################################################################################
#                                                                              #
#                         TB Experienced Participants                          #
#                                                                              #
################################################################################

#Screening
setCellValue(cells[["4.19"]], nrow(raw_data_baseline_ex_arm_1))
setCellValue(cells[["5.19"]], nrow(subset(raw_data_baseline_ex_arm_1, !is.na(raw_data_baseline_ex_arm_1$sc_aim))))
setCellValue(cells[["6.19"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_eligible=='Proceed')))
setCellValue(cells[["7.19"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_end_ip=='End' | raw_data_baseline_ex_arm_1$tbip_sc_age <18)))
setCellValue(cells[["8.19"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_below_age=='Yes' | raw_data_baseline_ex_arm_1$tbip_sc_below_age=='No')))
setCellValue(cells[["9.19"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_q17=='No')))
setCellValue(cells[["10.19"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_q8=='No')))
setCellValue(cells[["11.19"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_language=='No')))
setCellValue(cells[["12.19"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_q13=='No')))
setCellValue(cells[["13.19"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_existing=='Yes')))


#Enrolment
setCellValue(cells[["18.19"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_eligible=='Proceed')))
setCellValue(cells[["19.19"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_consent_part=='Yes')))
setCellValue(cells[["20.19"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_fw_note=='No')))
setCellValue(cells[["21.19"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_consent_part=='No')))
setCellValue(cells[["22.19"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_refuse___1=='I\'m not interested')))
setCellValue(cells[["23.19"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_refuse___2=='I am enrolled in another study')))
setCellValue(cells[["24.19"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_refuse___3=='I do not have time')))
setCellValue(cells[["25.19"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_refuse___4=='I am tired')))
setCellValue(cells[["26.19"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_refuse___5=='Other')))


#Caregiver eligibility
setCellValue(cells[["31.20"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_below_age=='Yes' | raw_data_baseline_ex_arm_1$tbip_sc_below_age=='No')))
setCellValue(cells[["32.20"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_below_age=='No')))
setCellValue(cells[["33.20"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_below_age=='Yes')))
setCellValue(cells[["34.20"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_cgiver_permission=='Yes')))
setCellValue(cells[["35.20"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_cgiver_permission=='No')))


#Aim 1 - Participant groups Newly initiated & Aim 1- Study visits / Retention

setCellValue(cells[["42.19"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_eligible=='Proceed')))
setCellValue(cells[["42.20"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_consent_part=='Yes')))
setCellValue(cells[["42.6"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_consent_part=='Yes')))
setCellValue(cells[["43.19"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_eligible=='Proceed')))
setCellValue(cells[["43.20"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_consent_part=='Yes')))
setCellValue(cells[["43.6"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_consent_part=='Yes')))
setCellValue(cells[["44.6"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_cgiver_permission=='Yes')))

#Clinic screening rate
setCellValue(cells[["50.20"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_q5=='Empilweni Gompo CHC')))
setCellValue(cells[["51.20"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_q5=='Pefferville Clinic')))
setCellValue(cells[["52.20"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_q5=='Duncan Village CHC')))
setCellValue(cells[["53.20"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_q5=='Gompo C Jabavu Clinic')))
setCellValue(cells[["54.20"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_q5=='Chris Hani Clinic')))
setCellValue(cells[["55.20"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_q5=='Luyolo NU 9 Clinic')))
setCellValue(cells[["56.20"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_q5=='Alphendale Clinic')))
setCellValue(cells[["57.20"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_q5=='John Dube Clinic')))
setCellValue(cells[["58.20"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_q5=='Fezeka NU 3 Clinic')))
setCellValue(cells[["59.20"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_q5=='Gompo A Ndende Clinic')))
setCellValue(cells[["60.20"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_q5=='Ndevana Clinic')))
setCellValue(cells[["61.20"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_q5=='Philani NU 1 Clinic')))
setCellValue(cells[["62.20"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_q5=='Aspiranza Clinic')))
setCellValue(cells[["63.20"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_q5=='Ginsberg Clinic')))
setCellValue(cells[["64.20"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_q5=='Zwelitsha Zone 5 Clinic')))
setCellValue(cells[["65.20"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_q5=='Masakhane Clinic (Zwelitsha)')))
setCellValue(cells[["66.20"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_q5=='Gompo B Jwayi Clinic')))
setCellValue(cells[["67.20"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_q5=='NU 12 Clinic')))

#Clinic eligible rate
setCellValue(cells[["50.21"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_ex_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_ex_arm_1$tbip_sc_q5=='Empilweni Gompo CHC')))
setCellValue(cells[["51.21"]], nrow(subset(raw_data_baseline_ex_arm_1, raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Pefferville Clinic')))
setCellValue(cells[["52.21"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Duncan Village CHC')))
setCellValue(cells[["53.21"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo C Jabavu Clinic')))
setCellValue(cells[["54.21"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Chris Hani Clinic')))
setCellValue(cells[["55.21"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Luyolo NU 9 Clinic')))
setCellValue(cells[["56.21"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Alphendale Clinic')))
setCellValue(cells[["57.21"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='John Dube Clinic')))
setCellValue(cells[["58.21"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Fezeka NU 3 Clinic')))
setCellValue(cells[["59.21"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo A Ndende Clinic')))
setCellValue(cells[["60.21"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Ndevana Clinic')))
setCellValue(cells[["61.21"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Philani NU 1 Clinic')))
setCellValue(cells[["62.21"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Aspiranza Clinic')))
setCellValue(cells[["63.21"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Ginsberg Clinic')))
setCellValue(cells[["64.21"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Zwelitsha Zone 5 Clinic')))
setCellValue(cells[["65.21"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Masakhane Clinic (Zwelitsha)')))
setCellValue(cells[["66.21"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo B Jwayi Clinic')))
setCellValue(cells[["67.21"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed' & raw_data_baseline_arm_1$tbip_sc_q5=='NU 12 Clinic')))


#Clinic enrolled rate
setCellValue(cells[["50.22"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Empilweni Gompo CHC')))
setCellValue(cells[["51.22"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Pefferville Clinic')))
setCellValue(cells[["52.22"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Duncan Village CHC')))
setCellValue(cells[["53.22"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo C Jabavu Clinic')))
setCellValue(cells[["54.22"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Chris Hani Clinic')))
setCellValue(cells[["55.22"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Luyolo NU 9 Clinic')))
setCellValue(cells[["56.22"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Alphendale Clinic')))
setCellValue(cells[["57.22"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='John Dube Clinic')))
setCellValue(cells[["58.22"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Fezeka NU 3 Clinic')))
setCellValue(cells[["59.22"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo A Ndende Clinic')))
setCellValue(cells[["60.22"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Ndevana Clinic')))
setCellValue(cells[["61.22"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Philani NU 1 Clinic')))
setCellValue(cells[["62.22"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Aspiranza Clinic')))
setCellValue(cells[["63.22"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Ginsberg Clinic')))
setCellValue(cells[["64.22"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Zwelitsha Zone 5 Clinic')))
setCellValue(cells[["65.22"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Masakhane Clinic (Zwelitsha)')))
setCellValue(cells[["66.22"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='Gompo B Jwayi Clinic')))
setCellValue(cells[["67.22"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes' & raw_data_baseline_arm_1$tbip_sc_q5=='NU 12 Clinic')))

##Participant time
setCellValue(cells[["72.19"]], mean(raw_data_baseline_arm_1$tbip_q3_duration, na.rm = TRUE))
setCellValue(cells[["73.19"]], median(raw_data_baseline_arm_1$tbip_q3_duration, na.rm = TRUE))
setCellValue(cells[["74.19"]], names(sort(-table(raw_data_baseline_arm_1$tbip_q3_duration)))[1])


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
setCellValue(cells[["31.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & (raw_data_baseline_arm_1$index_questionnaire_3_complete=='Complete' | raw_data_baseline_arm_1$index_questionnaire_3_complete=='Unverified'))))
setCellValue(cells[["31.7"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbr_smear_res_1=='Negative')))
setCellValue(cells[["31.8"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & (raw_data_baseline_arm_1$tbr_smear_res_1=='Positive 1+' | raw_data_baseline_arm_1$tbr_smear_1=='Positive 2+' | raw_data_baseline_arm_1$tbr_smear_res_1=='Positive 3+'))))
setCellValue(cells[["31.9"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$index_questionnaire_3_complete=='Complete' & raw_data_baseline_arm_1$tbip_sc_days_since_init>=49)))
setCellValue(cells[["31.11"]], nrow(subset(raw_data_follow_up_1_arm_1, raw_data_follow_up_1_arm_1$tbr_sputum_collected=='Yes')))
setCellValue(cells[["31.13"]], nrow(subset(raw_data_follow_up_1_arm_1, raw_data_follow_up_1_arm_1$tbr_smear_res_1=='Negative')))
setCellValue(cells[["31.15"]], nrow(subset(raw_data_follow_up_1_arm_1, raw_data_follow_up_1_arm_1$index_follow_up_questionnaire_3_complete=='Unverified' | raw_data_follow_up_1_arm_1$index_follow_up_questionnaire_3_complete=="Complete")))



#setCellValue(cells[["28.7"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed')))
#setCellValue(cells[["28.8"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes')))
#setCellValue(cells[["28.10"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes')))
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

################################################################################
#                                                                              #
#                             All Participants                                 #
#                                                                              #
################################################################################

#HHCI: Screening & Enrolment

#Households listed by IPs

setCellValue(cells[["6.3"]], nrow(raw_data_hhci_info_arm_1 %>% distinct(record_id)))
#setCellValue(cells[["7.3"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(hhc_collection_point=='Clinic') %>% distinct(record_id)))
#setCellValue(cells[["8.3"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(hhc_collection_point=='HH') %>% distinct(record_id)))

#Total households visited

setCellValue(cells[["8.3"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(hhc_sc_visit_attempt___1=='Checked' | hhc_sc_visit_attempt___2=='Checked' | hhc_sc_visit_attempt___3=='Checked'| tb_tf_study_time_point_hhc=='1') %>% distinct(record_id)))
setCellValue(cells[["9.3"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(hhc_sc_visit_attempt___1=='Checked') %>% distinct(record_id)))
setCellValue(cells[["10.3"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(hhc_sc_visit_attempt___2=='Checked') %>% distinct(record_id)))
setCellValue(cells[["11.3"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(hhc_sc_visit_attempt___3=='Checked') %>% distinct(record_id)))

#Pending HH Visits
setCellValue(cells[["13.3"]], nrow(raw_data_baseline_arm_1 %>% dplyr::filter(hhc_members_visited_1==0) %>% distinct(record_id)))

pending_hh <- raw_data_baseline_arm_1 %>% dplyr::filter(hhc_members_visited_1==0) %>% distinct(record_id)

write.csv(pending_hh, "Data/pending_hh.csv", row.names = FALSE, na='')

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

tb_asymptomatic <- subset(raw_data_hhci_info_arm_1, 
                          raw_data_hhci_info_arm_1$hhc_sc_weight_loss=='No' & 
                            raw_data_hhci_info_arm_1$hhc_sc_night_sweat=='No' & 
                            raw_data_hhci_info_arm_1$hhc_sc_coughing=='No' & 
                            raw_data_hhci_info_arm_1$hhc_sc_fever=='No')


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
setCellValue(cells[["63.6"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(hhc_days_since_referral<=30 & hhc_pt_testing_outcome=='Negative')))
setCellValue(cells[["64.6"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(hhc_days_since_referral<=30 & hhc_pt_testing_outcome=='Positive')))
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
setCellValue(cells[["68.6"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='No')))
setCellValue(cells[["69.6"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes' )))
setCellValue(cells[["70.6"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes' & hhc_pt_days_to_present<=30)))
setCellValue(cells[["71.6"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes' & hhc_pt_days_to_present>30)))
#setCellValue(cells[["71.6"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes')))
setCellValue(cells[["73.6"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes' & hhc_pt_collect_sputum=='No')))
setCellValue(cells[["74.6"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes' & hhc_pt_collect_sputum=='Yes')))
setCellValue(cells[["75.6"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_testing_outcome=='Negative')))
setCellValue(cells[["76.6"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_testing_outcome=='Positive')))
setCellValue(cells[["77.6"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_collect_sputum=='Yes'& is.na(hhc_pt_testing_outcome))))
#setCellValue(cells[["76.6"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & hhc_sc_verbal_consent=='Yes' & hhc_days_since_referral>30 & is.na(hhc_pt_return_clinic))))

#Self report and Extracted Only after 30 days
#setCellValue(cells[["55.6"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(hhc_sc_verbal_consent=='Yes' & hhc_days_since_referral>30)))
setCellValue(cells[["68.9"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(!is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='No')))
setCellValue(cells[["69.9"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(!is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes' )))
setCellValue(cells[["70.9"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(!is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes' & hhc_pt_days_to_present<=30)))
setCellValue(cells[["71.9"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(!is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes' & hhc_pt_days_to_present>30)))
#setCellValue(cells[["71.9"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter((hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes')))
setCellValue(cells[["73.9"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(!is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes' & hhc_pt_collect_sputum=='No')))
setCellValue(cells[["74.9"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(!is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes' & hhc_pt_collect_sputum=='Yes')))
setCellValue(cells[["75.9"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(!is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_testing_outcome=='Negative')))
setCellValue(cells[["76.9"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(!is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_testing_outcome=='Positive')))
setCellValue(cells[["77.9"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(!is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_collect_sputum=='Yes'& is.na(hhc_pt_testing_outcome))))

setCellValue(cells[["78.9"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(hhc_days_since_referral>30 & ((hhc_pc_been_facility=='Yes, I remember the date' | hhc_pc_been_facility=='Yes, I don\'t remember the date') & hhc_pt_return_clinic=='No') | (hhc_pc_been_facility=='No' & hhc_pt_return_clinic=='Yes'))))
setCellValue(cells[["79.12"]], nrow(raw_data_hhci_info_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & is.na(hhc_pt_return_clinic) & hhc_days_since_referral>30)))


################################################################################
#                                                                              #
#                             Newly Initiated Participants                     #
#                                                                              #
################################################################################

#HHCI: Screening & Enrolment

#Households listed by IPs

setCellValue(cells[["6.16"]], nrow(raw_data_hhci_info_ni_arm_1 %>% distinct(record_id)))
#setCellValue(cells[["7.16"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(hhc_collection_point=='Clinic') %>% distinct(record_id)))
#setCellValue(cells[["8.16"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(hhc_collection_point=='HH') %>% distinct(record_id)))

#Total households visited

setCellValue(cells[["8.16"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(hhc_sc_visit_attempt___1=='Checked' | hhc_sc_visit_attempt___2=='Checked' | hhc_sc_visit_attempt___3=='Checked') %>% distinct(record_id)))
setCellValue(cells[["9.16"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(hhc_sc_visit_attempt___1=='Checked') %>% distinct(record_id)))
setCellValue(cells[["10.16"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(hhc_sc_visit_attempt___2=='Checked') %>% distinct(record_id)))
setCellValue(cells[["11.16"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(hhc_sc_visit_attempt___3=='Checked') %>% distinct(record_id)))

#Pending HH Visits
setCellValue(cells[["13.16"]], nrow(raw_data_baseline_arm_1 %>% dplyr::filter(hhc_members_visited_1==0) %>% distinct(record_id)))

#Households with enrolled HHCs
setCellValue(cells[["16.16"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(hhc_sc_consent_provided=='Yes') %>% distinct(record_id)))
setCellValue(cells[["17.16"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(hhc_sc_consent_provided=='Yes') %>% distinct(record_id)))

#HHCs listed by IPs
setCellValue(cells[["27.16"]], nrow(subset(raw_data_hhci_info_ni_arm_1, !is.na(raw_data_hhci_info_ni_arm_1$hhcl_member_name))))
setCellValue(cells[["28.16"]], nrow(subset(raw_data_hhci_info_ni_arm_1, raw_data_hhci_info_ni_arm_1$hhcl_member_age>=18)))
setCellValue(cells[["29.16"]], nrow(subset(raw_data_hhci_info_ni_arm_1, raw_data_hhci_info_ni_arm_1$hhcl_member_age<18)))
setCellValue(cells[["30.16"]], nrow(subset(raw_data_hhci_info_ni_arm_1, is.na(raw_data_hhci_info_ni_arm_1$hhcl_member_age))))

setCellValue(cells[["32.16"]], nrow(subset(raw_data_hhci_info_ni_arm_1, raw_data_hhci_info_ni_arm_1$hhc_collection_point=='Clinic')))
setCellValue(cells[["33.16"]], nrow(subset(raw_data_hhci_info_ni_arm_1, raw_data_hhci_info_ni_arm_1$hhc_collection_point=='HH')))

#Screened for eligibility
setCellValue(cells[["35.16"]], nrow(subset(raw_data_hhci_info_ni_arm_1, !is.na(raw_data_hhci_info_ni_arm_1$hhc_sc_clinic_visit))))

#Not Eligible
setCellValue(cells[["36.16"]], nrow(subset(raw_data_hhci_info_ni_arm_1, hhc_sc_clinic_visit=='Yes' |
                                            as.integer(hhc_sc_age_calc)<18 |
                                            hhc_sc_on_treatment=='Yes' |
                                            hhc_sc_verbal_consent=='No' |
                                            hhc_sc_language=='No' |
                                            (hhc_sc_weight_loss=='No' &
                                               hhc_sc_night_sweat=='No' & 
                                               hhc_sc_coughing=='No' &
                                               hhc_sc_fever=='No'))))
setCellValue(cells[["37.16"]], nrow(subset(raw_data_hhci_info_ni_arm_1, hhc_sc_on_treatment=='Yes')))
setCellValue(cells[["38.16"]], nrow(subset(raw_data_hhci_info_ni_arm_1, raw_data_hhci_info_ni_arm_1$hhc_sc_clinic_visit=='Yes')))
setCellValue(cells[["39.16"]], nrow(subset(raw_data_hhci_info_ni_arm_1, 
                                          hhc_sc_weight_loss=='No' & 
                                            hhc_sc_night_sweat=='No' & 
                                            hhc_sc_coughing=='No' & 
                                            hhc_sc_fever=='No')))
setCellValue(cells[["40.16"]], nrow(subset(raw_data_hhci_info_ni_arm_1, hhc_sc_verbal_consent=='No')))
setCellValue(cells[["41.16"]], nrow(subset(raw_data_hhci_info_ni_arm_1, hhc_sc_language=='No')))
setCellValue(cells[["42.16"]], nrow(subset(raw_data_hhci_info_ni_arm_1, as.integer(hhc_sc_age_calc)<18)))

#Eligible 
setCellValue(cells[["43.16"]], nrow(subset(raw_data_hhci_info_ni_arm_1, raw_data_hhci_info_ni_arm_1$hhc_sc_cons_dir_3=='Proceed')))
setCellValue(cells[["44.16"]], nrow(subset(raw_data_hhci_info_ni_arm_1, raw_data_hhci_info_ni_arm_1$hhc_sc_consent_provided=='Yes')))
setCellValue(cells[["45.16"]], nrow(subset(raw_data_hhci_info_ni_arm_1, raw_data_hhci_info_ni_arm_1$hhc_sc_consent_provided=='No')))
setCellValue(cells[["46.16"]], nrow(subset(raw_data_hhci_info_ni_arm_1, raw_data_hhci_info_ni_arm_1$hhc_sc_competent=='No')))

#Outcomes
setCellValue(cells[["52.16"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(hhc_days_since_referral<=30 & hhc_sc_consent_provided=='Yes')))
setCellValue(cells[["53.16"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(hhc_days_since_referral>30 & hhc_sc_consent_provided=='Yes')))
setCellValue(cells[["54.16"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(hhc_sc_consent_provided=='Yes')))

#Extracted within the 30 Day window
setCellValue(cells[["59.19"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(hhc_days_since_referral<=30 & hhc_sc_verbal_consent=='Yes' )))
setCellValue(cells[["60.19"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(hhc_days_since_referral<=30 & hhc_pt_intro=='Proceed' & hhc_pt_return_clinic=='Yes')))
setCellValue(cells[["61.19"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(hhc_days_since_referral<=30 & hhc_pt_intro=='Proceed' & hhc_pt_collect_sputum=='Yes')))
setCellValue(cells[["62.19"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(hhc_days_since_referral<=30 & hhc_pt_intro=='Proceed' & hhc_pt_collect_sputum=='No')))
setCellValue(cells[["63.19"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(hhc_days_since_referral<=30 & hhc_pt_testing_outcome=='Negative')))
setCellValue(cells[["64.19"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(hhc_days_since_referral<=30 & hhc_pt_testing_outcome=='Positive')))
setCellValue(cells[["65.19"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(hhc_days_since_referral<=30 & hhc_pt_collect_sputum=='Yes' & is.na(hhc_pc_testing_outcome))))

setCellValue(cells[["59.25"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(hhc_days_since_referral<=30 & hhc_sc_verbal_consent=='Yes' & is.na(hhc_pt_intro))))

#Self reported only after 30 days
setCellValue(cells[["68.16"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(is.na(hhc_pt_return_clinic) & hhc_days_since_referral>30 & hhc_pc_been_facility=='No')))
#setCellValue(cells[["55.16"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(hhc_sc_verbal_consent=='Yes' & hhc_days_since_referral>30 & is.na(hhc_pt_intro))))
setCellValue(cells[["69.16"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(is.na(hhc_pt_intro) & hhc_days_since_referral>30 & (hhc_pc_been_facility=='Yes, I remember the date' | hhc_pc_been_facility=='Yes, I don\'t remember the date'))))
setCellValue(cells[["70.16"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(is.na(hhc_pt_intro) & hhc_days_since_referral>30 & (hhc_pc_been_facility=='Yes, I remember the date' & hhc_pc_days_to_present<=30))))
setCellValue(cells[["71.16"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(is.na(hhc_pt_intro) & hhc_days_since_referral>30 & (hhc_pc_been_facility=='Yes, I remember the date' & hhc_pc_days_to_present>30))))
setCellValue(cells[["72.16"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(is.na(hhc_pt_intro) & hhc_days_since_referral>30 & hhc_pc_been_facility=='Yes, I don\'t remember the date')))
setCellValue(cells[["73.16"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(is.na(hhc_pt_intro) & hhc_days_since_referral>30 & (hhc_pc_been_facility=='Yes, I remember the date' | hhc_pc_been_facility=='Yes, I don\'t remember the date') & hhc_pc_provide_sputum=='No')))
setCellValue(cells[["74.16"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(is.na(hhc_pt_intro) & hhc_days_since_referral>30 & (hhc_pc_been_facility=='Yes, I remember the date' | hhc_pc_been_facility=='Yes, I don\'t remember the date') & hhc_pc_provide_sputum=='Yes')))
setCellValue(cells[["75.16"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(is.na(hhc_pt_intro) & hhc_days_since_referral>30 & hhc_pc_testing_outcome=='Negative')))
setCellValue(cells[["76.16"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(is.na(hhc_pt_intro) & hhc_days_since_referral>30 & hhc_pc_testing_outcome=='Positive')))
setCellValue(cells[["77.16"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(is.na(hhc_pt_intro) & hhc_days_since_referral>30 & hhc_pc_provide_sputum=='Yes' & is.na(hhc_pc_testing_outcome))))
#setCellValue(cells[["76.16"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(is.na(hhc_pt_intro) & hhc_sc_verbal_consent=='Yes' & hhc_days_since_referral>30 & is.na(hhc_pc_been_facility))))

#Extracted Only after 30 days
#setCellValue(cells[["55.19"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(hhc_sc_verbal_consent=='Yes' & hhc_days_since_referral>30)))
setCellValue(cells[["68.19"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='No')))
setCellValue(cells[["69.19"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes' )))
setCellValue(cells[["70.19"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes' & hhc_pt_days_to_present<=30)))
setCellValue(cells[["71.19"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes' & hhc_pt_days_to_present>30)))
#setCellValue(cells[["71.19"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes')))
setCellValue(cells[["73.19"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes' & hhc_pt_collect_sputum=='No')))
setCellValue(cells[["74.19"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes' & hhc_pt_collect_sputum=='Yes')))
setCellValue(cells[["75.19"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_testing_outcome=='Negative')))
setCellValue(cells[["76.19"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_testing_outcome=='Positive')))
setCellValue(cells[["77.19"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_collect_sputum=='Yes'& is.na(hhc_pt_testing_outcome))))
#setCellValue(cells[["76.19"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & hhc_sc_verbal_consent=='Yes' & hhc_days_since_referral>30 & is.na(hhc_pt_return_clinic))))

#Self report and Extracted Only after 30 days
#setCellValue(cells[["55.19"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(hhc_sc_verbal_consent=='Yes' & hhc_days_since_referral>30)))
print(nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(!is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes' )))
setCellValue(cells[["68.22"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(hhc_pc_been_facility=='No' & hhc_days_since_referral>30 & hhc_pt_return_clinic=='No')))
setCellValue(cells[["69.22"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(!is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes' )))
setCellValue(cells[["70.22"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(!is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes' & hhc_pt_days_to_present<=30)))
setCellValue(cells[["71.22"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(!is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes' & hhc_pt_days_to_present>30)))
#setCellValue(cells[["71.22"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(!is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes')))
setCellValue(cells[["73.22"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(!is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes' & hhc_pt_collect_sputum=='No')))
setCellValue(cells[["74.22"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(!is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes' & hhc_pt_collect_sputum=='Yes')))
setCellValue(cells[["75.22"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(!is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_testing_outcome=='Negative')))
setCellValue(cells[["76.22"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(!is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_testing_outcome=='Positive')))
setCellValue(cells[["77.22"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(!is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_collect_sputum=='Yes'& is.na(hhc_pt_testing_outcome))))

setCellValue(cells[["78.22"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(hhc_days_since_referral>30 & ((hhc_pc_been_facility=='Yes, I remember the date' | hhc_pc_been_facility=='Yes, I don\'t remember the date') & hhc_pt_return_clinic=='No') | (hhc_pc_been_facility=='No' & hhc_pt_return_clinic=='Yes'))))
setCellValue(cells[["79.25"]], nrow(raw_data_hhci_info_ni_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & is.na(hhc_pt_return_clinic) & hhc_days_since_referral>30)))


################################################################################
#                                                                              #
#                             TB Experienced Participants                     #
#                                                                              #
################################################################################

#HHCI: Screening & Enrolment

#Households listed by IPs

setCellValue(cells[["6.29"]], nrow(raw_data_hhci_info_ex_arm_1 %>% distinct(record_id)))
#setCellValue(cells[["7.29"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(hhc_collection_point=='Clinic') %>% distinct(record_id)))
#setCellValue(cells[["8.29"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(hhc_collection_point=='HH') %>% distinct(record_id)))

#Total households visited

setCellValue(cells[["8.29"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(hhc_sc_visit_attempt___1=='Checked' | hhc_sc_visit_attempt___2=='Checked' | hhc_sc_visit_attempt___3=='Checked') %>% distinct(record_id)))
setCellValue(cells[["9.29"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(hhc_sc_visit_attempt___1=='Checked') %>% distinct(record_id)))
setCellValue(cells[["10.29"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(hhc_sc_visit_attempt___2=='Checked') %>% distinct(record_id)))
setCellValue(cells[["11.29"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(hhc_sc_visit_attempt___3=='Checked') %>% distinct(record_id)))

#Pending HH Visits
setCellValue(cells[["13.29"]], nrow(raw_data_baseline_arm_1 %>% dplyr::filter(hhc_members_visited_1==0) %>% distinct(record_id)))

#Households with enrolled HHCs
setCellValue(cells[["16.29"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(hhc_sc_consent_provided=='Yes') %>% distinct(record_id)))
setCellValue(cells[["17.29"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(hhc_sc_consent_provided=='Yes') %>% distinct(record_id)))

#HHCs listed by IPs
setCellValue(cells[["27.29"]], nrow(subset(raw_data_hhci_info_ex_arm_1, !is.na(raw_data_hhci_info_ex_arm_1$hhcl_member_name))))
setCellValue(cells[["28.29"]], nrow(subset(raw_data_hhci_info_ex_arm_1, raw_data_hhci_info_ex_arm_1$hhcl_member_age>=18)))
setCellValue(cells[["29.29"]], nrow(subset(raw_data_hhci_info_ex_arm_1, raw_data_hhci_info_ex_arm_1$hhcl_member_age<18)))
setCellValue(cells[["30.29"]], nrow(subset(raw_data_hhci_info_ex_arm_1, is.na(raw_data_hhci_info_ex_arm_1$hhcl_member_age))))

setCellValue(cells[["32.29"]], nrow(subset(raw_data_hhci_info_ex_arm_1, raw_data_hhci_info_ex_arm_1$hhc_collection_point=='Clinic')))
setCellValue(cells[["33.29"]], nrow(subset(raw_data_hhci_info_ex_arm_1, raw_data_hhci_info_ex_arm_1$hhc_collection_point=='HH')))

#Screened for eligibility
setCellValue(cells[["35.29"]], nrow(subset(raw_data_hhci_info_ex_arm_1, !is.na(raw_data_hhci_info_ex_arm_1$hhc_sc_clinic_visit))))

#Not Eligible
setCellValue(cells[["36.29"]], nrow(subset(raw_data_hhci_info_ex_arm_1, hhc_sc_clinic_visit=='Yes' |
                                            as.integer(hhc_sc_age_calc)<18 |
                                            hhc_sc_on_treatment=='Yes' |
                                            hhc_sc_verbal_consent=='No' |
                                            hhc_sc_language=='No' |
                                            (hhc_sc_weight_loss=='No' &
                                               hhc_sc_night_sweat=='No' & 
                                               hhc_sc_coughing=='No' &
                                               hhc_sc_fever=='No'))))
setCellValue(cells[["37.29"]], nrow(subset(raw_data_hhci_info_ex_arm_1, hhc_sc_on_treatment=='Yes')))
setCellValue(cells[["38.29"]], nrow(subset(raw_data_hhci_info_ex_arm_1, raw_data_hhci_info_ex_arm_1$hhc_sc_clinic_visit=='Yes')))
setCellValue(cells[["39.29"]], nrow(subset(raw_data_hhci_info_ex_arm_1, 
                                          hhc_sc_weight_loss=='No' & 
                                            hhc_sc_night_sweat=='No' & 
                                            hhc_sc_coughing=='No' & 
                                            hhc_sc_fever=='No')))
setCellValue(cells[["40.29"]], nrow(subset(raw_data_hhci_info_ex_arm_1, hhc_sc_verbal_consent=='No')))
setCellValue(cells[["41.29"]], nrow(subset(raw_data_hhci_info_ex_arm_1, hhc_sc_language=='No')))
setCellValue(cells[["42.29"]], nrow(subset(raw_data_hhci_info_ex_arm_1, as.integer(hhc_sc_age_calc)<18)))

#Eligible 
setCellValue(cells[["43.29"]], nrow(subset(raw_data_hhci_info_ex_arm_1, raw_data_hhci_info_ex_arm_1$hhc_sc_cons_dir_3=='Proceed')))
setCellValue(cells[["44.29"]], nrow(subset(raw_data_hhci_info_ex_arm_1, raw_data_hhci_info_ex_arm_1$hhc_sc_consent_provided=='Yes')))
setCellValue(cells[["45.29"]], nrow(subset(raw_data_hhci_info_ex_arm_1, raw_data_hhci_info_ex_arm_1$hhc_sc_consent_provided=='No')))
setCellValue(cells[["46.29"]], nrow(subset(raw_data_hhci_info_ex_arm_1, raw_data_hhci_info_ex_arm_1$hhc_sc_competent=='No')))

#Outcomes
setCellValue(cells[["52.29"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(hhc_days_since_referral<=30 & hhc_sc_consent_provided=='Yes')))
setCellValue(cells[["53.29"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(hhc_days_since_referral>30 & hhc_sc_consent_provided=='Yes')))
setCellValue(cells[["54.29"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(hhc_sc_consent_provided=='Yes')))

#Extracted within the 30 Day window
setCellValue(cells[["59.32"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(hhc_days_since_referral<=30 & hhc_sc_verbal_consent=='Yes' )))
setCellValue(cells[["60.32"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(hhc_days_since_referral<=30 & hhc_pt_intro=='Proceed' & hhc_pt_return_clinic=='Yes')))
setCellValue(cells[["61.32"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(hhc_days_since_referral<=30 & hhc_pt_intro=='Proceed' & hhc_pt_collect_sputum=='Yes')))
setCellValue(cells[["62.32"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(hhc_days_since_referral<=30 & hhc_pt_intro=='Proceed' & hhc_pt_collect_sputum=='No')))
setCellValue(cells[["63.32"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(hhc_days_since_referral<=30 & hhc_pt_testing_outcome=='Negative')))
setCellValue(cells[["64.32"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(hhc_days_since_referral<=30 & hhc_pt_testing_outcome=='Positive')))
setCellValue(cells[["65.32"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(hhc_days_since_referral<=30 & hhc_pt_collect_sputum=='Yes' & is.na(hhc_pt_testing_outcome))))

setCellValue(cells[["59.38"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(hhc_days_since_referral<=30 & hhc_sc_verbal_consent=='Yes' & is.na(hhc_pt_intro))))

#Self reported only after 30 days
setCellValue(cells[["68.29"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(is.na(hhc_pt_return_clinic) & hhc_days_since_referral>30 & hhc_pc_been_facility=='No')))
#setCellValue(cells[["55.29"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(hhc_sc_verbal_consent=='Yes' & hhc_days_since_referral>30 & is.na(hhc_pt_intro))))
setCellValue(cells[["69.29"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(is.na(hhc_pt_intro) & hhc_days_since_referral>30 & (hhc_pc_been_facility=='Yes, I remember the date' | hhc_pc_been_facility=='Yes, I don\'t remember the date'))))
setCellValue(cells[["70.29"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(is.na(hhc_pt_intro) & hhc_days_since_referral>30 & (hhc_pc_been_facility=='Yes, I remember the date' & hhc_pc_days_to_present<=30))))
setCellValue(cells[["71.29"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(is.na(hhc_pt_intro) & hhc_days_since_referral>30 & (hhc_pc_been_facility=='Yes, I remember the date' & hhc_pc_days_to_present>30))))
setCellValue(cells[["72.29"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(is.na(hhc_pt_intro) & hhc_days_since_referral>30 & hhc_pc_been_facility=='Yes, I don\'t remember the date')))
setCellValue(cells[["73.29"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(is.na(hhc_pt_intro) & hhc_days_since_referral>30 & (hhc_pc_been_facility=='Yes, I remember the date' | hhc_pc_been_facility=='Yes, I don\'t remember the date') & hhc_pc_provide_sputum=='No')))
setCellValue(cells[["74.29"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(is.na(hhc_pt_intro) & hhc_days_since_referral>30 & (hhc_pc_been_facility=='Yes, I remember the date' | hhc_pc_been_facility=='Yes, I don\'t remember the date') & hhc_pc_provide_sputum=='Yes')))
setCellValue(cells[["75.29"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(is.na(hhc_pt_intro) & hhc_days_since_referral>30 & hhc_pc_testing_outcome=='Negative')))
setCellValue(cells[["76.29"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(is.na(hhc_pt_intro) & hhc_days_since_referral>30 & hhc_pc_testing_outcome=='Positive')))
setCellValue(cells[["77.29"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(is.na(hhc_pt_intro) & hhc_days_since_referral>30 & hhc_pc_provide_sputum=='Yes' & is.na(hhc_pc_testing_outcome))))
#setCellValue(cells[["76.29"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(is.na(hhc_pt_intro) & hhc_sc_verbal_consent=='Yes' & hhc_days_since_referral>30 & is.na(hhc_pc_been_facility))))

#Extracted Only after 30 days
#setCellValue(cells[["55.32"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(hhc_sc_verbal_consent=='Yes' & hhc_days_since_referral>30)))
setCellValue(cells[["68.32"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='No')))
setCellValue(cells[["69.32"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes' )))
setCellValue(cells[["70.32"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes' & hhc_pt_days_to_present<=30)))
setCellValue(cells[["71.32"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes' & hhc_pt_days_to_present>30)))
#setCellValue(cells[["71.32"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes')))
setCellValue(cells[["73.32"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes' & hhc_pt_collect_sputum=='No')))
setCellValue(cells[["74.32"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes' & hhc_pt_collect_sputum=='Yes')))
setCellValue(cells[["75.32"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_testing_outcome=='Negative')))
setCellValue(cells[["76.32"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_testing_outcome=='Positive')))
setCellValue(cells[["77.32"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_collect_sputum=='Yes'& is.na(hhc_pt_testing_outcome))))
#setCellValue(cells[["76.32"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & hhc_sc_verbal_consent=='Yes' & hhc_days_since_referral>30 & is.na(hhc_pt_return_clinic))))

#Self report and Extracted Only after 30 days
#setCellValue(cells[["55.32"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(hhc_sc_verbal_consent=='Yes' & hhc_days_since_referral>30)))
setCellValue(cells[["68.35"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(!is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='No')))
setCellValue(cells[["69.35"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(!is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes' )))
setCellValue(cells[["70.35"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(!is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes' & hhc_pt_days_to_present<=30)))
setCellValue(cells[["71.35"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(!is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes' & hhc_pt_days_to_present>30)))
#setCellValue(cells[["71.35"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(!is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes')))
setCellValue(cells[["73.35"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(!is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes' & hhc_pt_collect_sputum=='No')))
setCellValue(cells[["74.35"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(!is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_return_clinic=='Yes' & hhc_pt_collect_sputum=='Yes')))
setCellValue(cells[["75.35"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(!is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_testing_outcome=='Negative')))
setCellValue(cells[["76.35"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(!is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_testing_outcome=='Positive')))
setCellValue(cells[["77.35"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(!is.na(hhc_pc_been_facility) & hhc_days_since_referral>30 & hhc_pt_collect_sputum=='Yes'& is.na(hhc_pt_testing_outcome))))

setCellValue(cells[["78.35"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(hhc_days_since_referral>30 & ((hhc_pc_been_facility=='Yes, I remember the date' | hhc_pc_been_facility=='Yes, I don\'t remember the date') & hhc_pt_return_clinic=='No') | (hhc_pc_been_facility=='No' & hhc_pt_return_clinic=='Yes'))))
setCellValue(cells[["79.38"]], nrow(raw_data_hhci_info_ex_arm_1 %>% dplyr::filter(is.na(hhc_pc_been_facility) & is.na(hhc_pt_return_clinic) & hhc_days_since_referral>30)))

xlsx::forceFormulaRefresh(filename_new)
xlsx::saveWorkbook(wb, filename_new)
#xlsx::saveWorkbook(wb, "Data/TBStigmaWeeklyReport.xlsx")

print("Outcome - End")

gc()