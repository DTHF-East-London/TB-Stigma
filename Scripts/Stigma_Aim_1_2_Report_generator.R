library(dplyr)
library(xlsx)
source("Scripts/functions.R")

source("Scripts/dataset_generator_1.R")

wb <- xlsx::loadWorkbook("data/TB Stigma_Aim 1 and Aim 2_20230613.xlsx")

works_sheets <- xlsx::getSheets(wb)

tmp_sheet <- works_sheets[["Aim 1 - Index Patients"]]

rows <- getRows(tmp_sheet)

cells <- getCells(rows)

table()

#Screening

setCellValue(cells[["3.3"]], nrow(raw_data_baseline_arm_1))
setCellValue(cells[["4.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed')))
setCellValue(cells[["5.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_end_ip=='End Not-Eligible')))
setCellValue(cells[["6.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_below_age=='yes <18 years old')))
setCellValue(cells[["7.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q17=='No Lives outside study communities')))
setCellValue(cells[["8.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q8=='No Not Pulmonary TB Positive')))
setCellValue(cells[["9.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_language=='No Not fluent in Xhosa or English')))
setCellValue(cells[["10.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q14=='No household contacts')))
setCellValue(cells[["11.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_existing=='Yes Enrolled on other Cohort / HHC')))

#Enrolment

setCellValue(cells[["17.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes Agreed to participate')))
setCellValue(cells[["18.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_end_ip=='No Refused to participate')))
setCellValue(cells[["19.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_refuse==' Not interested')))
setCellValue(cells[["20.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_refuse==' Enrolled in another study')))
setCellValue(cells[["21.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_refuse=='Do not have time')))
setCellValue(cells[["22.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_refuse=='Tired')))
setCellValue(cells[["23.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_refuse=='Other')))


#Caregiver eligibility
setCellValue(cells[["28.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_below_age=='No')))
setCellValue(cells[["29.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_below_age=='Yes')))
setCellValue(cells[["30.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_cgiver_permission=='Yes')))
setCellValue(cells[["31.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_cgiver_permission=='No')))


#Aim 1 - Participant groups Newly initiated & Aim 1- Study visits / Retention

setCellValue(cells[["39.4"]], nrow(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed Eligible'))
setCellValue(cells[["39.8"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes Enrolled')))
setCellValue(cells[["39.9"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_cgiver_permission=='Yes Permission to visit HH')))
setCellValue(cells[["39.10"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_cgiver_permission=='Yes Permission to visit HH')))
setCellValue(cells[["39.11"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_cgiver_permission=='Yes Permission to visit HH')))
setCellValue(cells[["39.12"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_cgiver_permission=='Yes Permission to visit HH')))


##Rx Experienced

setCellValue(cells[["40.4"]], nrow(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed Eligible'))
setCellValue(cells[["40.8"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes Enrolled')))
setCellValue(cells[["40.9"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_cgiver_permission=='Yes Permission to visit HH')))
setCellValue(cells[["40.10"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_cgiver_permission=='Yes Permission to visit HH')))
setCellValue(cells[["40.11"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_cgiver_permission=='Yes Permission to visit HH')))
setCellValue(cells[["40.12"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_cgiver_permission=='Yes Permission to visit HH')))

##Under 18 with caregiver

setCellValue(cells[["41.4"]], nrow(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed Eligible'))
setCellValue(cells[["41.8"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc<14 & raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes Enrolled')))
setCellValue(cells[["41.9"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_cgiver_permission=='Yes Permission to visit HH')))
setCellValue(cells[["41.10"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_cgiver_permission=='Yes Permission to visit HH')))
setCellValue(cells[["41.11"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_cgiver_permission=='Yes Permission to visit HH')))
setCellValue(cells[["41.12"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_ini_days_calc>=14 & raw_data_baseline_arm_1$tbip_sc_cgiver_permission=='Yes Permission to visit HH')))


#Clinic enrolment rate

setCellValue(cells[["47.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q5=='Empilweni Gompo CHC')))
setCellValue(cells[["48.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q5=='Pefferville Clinic')))
setCellValue(cells[["49.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q5=='Duncan Village CHC')))
setCellValue(cells[["50.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q5=='Gompo C Jabavu Clinic')))
setCellValue(cells[["51.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q5=='Chris Hani Clinic')))
setCellValue(cells[["52.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q5=='Luyolo NU 9 Clinic')))
setCellValue(cells[["53.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q5=='Alphendale Clinic')))
setCellValue(cells[["54.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q5=='John Dube Clinic')))
setCellValue(cells[["55.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q5=='Fezeka NU 3 Clinic')))
setCellValue(cells[["56.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q5=='Gompo A Ndende Clinic')))
setCellValue(cells[["57.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q5=='Ndevana Clinic')))
setCellValue(cells[["58.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q5=='Philani NU 1 Clinic')))
setCellValue(cells[["59.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q5=='Aspiranza Clinic')))
setCellValue(cells[["60.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q5=='Ginsberg Clinic')))
setCellValue(cells[["61.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q5=='Zwelitsha Zone 5 Clinic')))
setCellValue(cells[["62.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q5=='Masakhane Clinic (Zwelitsha)')))
setCellValue(cells[["63.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q5=='Gompo B Jwayi Clinic')))
setCellValue(cells[["64.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_q5=='NU 12 Clinic')))


##Participant time

setCellValue(cells[["69.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_q3_duration=='mean')))
setCellValue(cells[["70.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_q3_duration=='Median')))
setCellValue(cells[["71.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_q3_duration=='Mode')))



#HHCI: Screening & Enrolment

#Households listed by IPs

setCellValue(cells[["4.3"]], nrow(raw_data_hhci_visit_arm_1))

#Total households visited

setCellValue(cells[["7.3"]], nrow(subset(raw_data_hhci_visit_arm_1, raw_data_hhci_visit_arm_1$tbip_sc_eligible=='Households visited once')))
setCellValue(cells[["8.3"]], nrow(subset(raw_data_hhci_visit_arm_1, raw_data_hhci_visit_arm_1$tbip_sc_end_ip==' Households visited twice')))
setCellValue(cells[["9.3"]], nrow(subset(raw_data_hhci_visit_arm_1, raw_data_hhci_visit_arm_1$tbip_sc_end_ip=='Households visited thrice')))


#HHCs listed by IPs

setCellValue(cells[["12.3"]], nrow(subset(raw_data_hhci_visit_arm_1, raw_data_hhci_visit_arm_1$hhc_sc_age_calce>=18)))
setCellValue(cells[["13.3"]], nrow(subset(raw_data_hhci_visit_arm_1, raw_data_hhci_visit_arm_1$hhc_sc_age_calc<18)))


#Not Eligible

setCellValue(cells[["18.3"]], nrow(subset(raw_data_hhci_visit_arm_1, raw_data_hhci_visit_arm_1$hhc_sc_clinic_visit=='Yes Visited clinic before HHCI')))
setCellValue(cells[["19.3"]], nrow(subset(raw_data_hhci_visit_arm_1, raw_data_hhci_visit_arm_1$hhc_sc_symptomatic_confirm=='Proceed TB asymptomatic')))
setCellValue(cells[["20.3"]], nrow(subset(raw_data_hhci_visit_arm_1, raw_data_hhci_visit_arm_1$hhc_sc_cons_dir_1_hhm=='Proceed <18 Years Old')))

#Eligible 
setCellValue(cells[["22.3"]], nrow(subset(raw_data_hhci_visit_arm_1, raw_data_hhci_visit_arm_1$hhc_sc_consent_provided=='Yes Agreed to participate')))
setCellValue(cells[["23.3"]], nrow(subset(raw_data_hhci_visit_arm_1, raw_data_hhci_visit_arm_1$hhc_sc_consent_provided=='No Refused to participate')))


#Outcomes 

setCellValue(cells[["29.3"]], nrow(subset(raw_data_hhci_visit_arm_1, raw_data_hhci_visit_arm_1$hhc_sc_clinic_visit=='Yes Presented at clinic')))
setCellValue(cells[["30.3"]], nrow(subset(raw_data_hhci_visit_arm_1, raw_data_hhci_visit_arm_1$hhc_pc_provide_sputum=='Yes Provided sputum')))
setCellValue(cells[["31.3"]], nrow(subset(raw_data_hhci_visit_arm_1, raw_data_hhci_visit_arm_1$hhc_sc_clinic_visit=='No Not presented yet')))


setCellValue(cells[["34.3"]], nrow(subset(raw_data_hhci_visit_arm_1, raw_data_hhci_visit_arm_1$tbip_sc_end_ip=='Presented, more than 30 days')))
setCellValue(cells[["35.3"]], nrow(subset(raw_data_hhci_visit_arm_1, raw_data_hhci_visit_arm_1$hhc_sc_clinic_visit=='No Did not present ')))


xlsx::forceFormulaRefresh(paste("Data/TB Stigma_Aim 1 and Aim 2_20230613", today, ".xlsx"))
xlsx::saveWorkbook(wb, paste("Data/TBStigma_weekly_Report_june_2023", today, ".xlsx"))

print("Outcome - End")

