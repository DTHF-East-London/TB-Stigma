library(dplyr)
library(xlsx)
source("Scripts/functions.R")

wb1 <- xlsx::loadWorkbook("Metadata/Weekly Enrolment Chart Template.xlsx")

works_sheets <- xlsx::getSheets(wb1)

tmp_sheet <- works_sheets[["Week 1"]]

rows <- getRows(tmp_sheet)

cells <- getCells(rows)

today <- format(Sys.time(), "%Y-%m-%d")

filename_new <- paste("Data/Weekly Enrolment Chart",today,".xlsx")

setCellValue(cells[["8.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$approach_date >= Sys.Date() - 7 & 
                                           (raw_data_baseline_arm_1$approach_facility=='Empilweni Gompo CHC'))))


setCellValue(cells[["9.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$approach_date >= Sys.Date() - 7 & 
                                           (raw_data_baseline_arm_1$approach_facility=='Pefferville Clinic'))))


setCellValue(cells[["10.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$approach_date >= Sys.Date() - 7 & 
                                           (raw_data_baseline_arm_1$approach_facility=='Duncan Village CHC'))))


setCellValue(cells[["11.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$approach_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$approach_facility=='Gompo C Jabavu Clinic'))))


setCellValue(cells[["12.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$approach_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$approach_facility=='Chris Hani Clinic'))))


setCellValue(cells[["13.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$approach_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$approach_facility=='Luyolo NU 9 Clinic'))))


setCellValue(cells[["14.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$approach_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$approach_facility=='Alphendale Clinic'))))


setCellValue(cells[["15.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$approach_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$approach_facility=='John Dube Clinic'))))


setCellValue(cells[["16.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$approach_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$approach_facility=='Fezeka NU 3 Clinic'))))


setCellValue(cells[["17.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$approach_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$approach_facility=='Gompo A Ndende Clinic'))))


setCellValue(cells[["18.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$approach_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$approach_facility=='Ndevana Clinic'))))


setCellValue(cells[["19.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$approach_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$approach_facility=='Philani NU 1 Clinic'))))


setCellValue(cells[["20.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$approach_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$approach_facility=='Aspiranza Clinic'))))


setCellValue(cells[["21.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$approach_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$approach_facility=='Ginsberg Clinic'))))


setCellValue(cells[["22.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$approach_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$approach_facility=='Zwelitsha Zone 5 Clinic'))))


setCellValue(cells[["23.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$approach_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$approach_facility=='Masakhane Clinic (Zwelitsha)'))))


setCellValue(cells[["24.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$approach_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$approach_facility=='Gompo B Jwayi Clinic'))))


setCellValue(cells[["25.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$approach_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$approach_facility=='NU 12 Clinici'))))

######################################################################################################
#                                                        Screeened                                   #
######################################################################################################

setCellValue(cells[["8.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                           (raw_data_baseline_arm_1$tbip_sc_q5=='Empilweni Gompo CHC'))))


setCellValue(cells[["9.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                           (raw_data_baseline_arm_1$tbip_sc_q5=='Pefferville Clinic'))))


setCellValue(cells[["10.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$tbip_sc_q5=='Duncan Village CHC'))))


setCellValue(cells[["11.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$tbip_sc_q5=='Gompo C Jabavu Clinic'))))


setCellValue(cells[["12.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$tbip_sc_q5=='Chris Hani Clinic'))))


setCellValue(cells[["13.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$tbip_sc_q5=='Luyolo NU 9 Clinic'))))


setCellValue(cells[["14.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$tbip_sc_q5=='Alphendale Clinic'))))


setCellValue(cells[["15.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$tbip_sc_q5=='John Dube Clinic'))))


setCellValue(cells[["16.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$tbip_sc_q5=='Fezeka NU 3 Clinic'))))


setCellValue(cells[["17.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$tbip_sc_q5=='Gompo A Ndende Clinic'))))


setCellValue(cells[["18.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$tbip_sc_q5=='Ndevana Clinic'))))


setCellValue(cells[["19.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$tbip_sc_q5=='Philani NU 1 Clinic'))))


setCellValue(cells[["20.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$tbip_sc_q5=='Aspiranza Clinic'))))


setCellValue(cells[["21.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$tbip_sc_q5=='Ginsberg Clinic'))))


setCellValue(cells[["22.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$tbip_sc_q5=='Zwelitsha Zone 5 Clinic'))))


setCellValue(cells[["23.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$tbip_sc_q5=='Masakhane Clinic (Zwelitsha)'))))


setCellValue(cells[["24.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$tbip_sc_q5=='Gompo B Jwayi Clinic'))))


setCellValue(cells[["25.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$tbip_sc_q5=='NU 12 Clinici'))))


###############################################Eligible#########################################


setCellValue(cells[["8.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                           (raw_data_baseline_arm_1$tbip_sc_q5=='Empilweni Gompo CHC') & (raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed'))))


setCellValue(cells[["9.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                           (raw_data_baseline_arm_1$tbip_sc_q5=='Pefferville Clinic') & (raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed'))))


setCellValue(cells[["10.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                           (raw_data_baseline_arm_1$tbip_sc_q5=='Duncan Village CHC') & (raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed'))))


setCellValue(cells[["11.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                           (raw_data_baseline_arm_1$tbip_sc_q5=='Gompo C Jabavu Clinic') & (raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed'))))


setCellValue(cells[["12.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                           (raw_data_baseline_arm_1$tbip_sc_q5=='Chris Hani Clinic') & (raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed'))))


setCellValue(cells[["13.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                           (raw_data_baseline_arm_1$tbip_sc_q5=='Luyolo NU 9 Clinic') & (raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed'))))


setCellValue(cells[["14.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                           (raw_data_baseline_arm_1$tbip_sc_q5=='Alphendale Clinic') & (raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed'))))


setCellValue(cells[["15.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                           (raw_data_baseline_arm_1$tbip_sc_q5=='John Dube Clinic') & (raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed'))))


setCellValue(cells[["16.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                           (raw_data_baseline_arm_1$tbip_sc_q5=='Fezeka NU 3 Clinic') & (raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed'))))


setCellValue(cells[["17.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                           (raw_data_baseline_arm_1$tbip_sc_q5=='Gompo A Ndende Clinic') & (raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed'))))


setCellValue(cells[["18.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                           (raw_data_baseline_arm_1$tbip_sc_q5=='Ndevana Clinic') & (raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed'))))


setCellValue(cells[["19.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                           (raw_data_baseline_arm_1$tbip_sc_q5=='Philani NU 1 Clinic') & (raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed'))))


setCellValue(cells[["20.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                           (raw_data_baseline_arm_1$tbip_sc_q5=='Aspiranza Clinic') & (raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed'))))


setCellValue(cells[["21.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                           (raw_data_baseline_arm_1$tbip_sc_q5=='Ginsberg Clinic') & (raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed'))))


setCellValue(cells[["22.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                           (raw_data_baseline_arm_1$tbip_sc_q5=='Zwelitsha Zone 5 Clinic') & (raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed'))))


setCellValue(cells[["23.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                           (raw_data_baseline_arm_1$tbip_sc_q5=='Masakhane Clinic (Zwelitsha)') & (raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed'))))


setCellValue(cells[["24.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                           (raw_data_baseline_arm_1$tbip_sc_q5=='Gompo B Jwayi Clinic') & (raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed'))))


setCellValue(cells[["25.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$tbip_sc_q5=='NU 12 Clinici') & (raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed'))))


######################################Enrolled on Study##########################################

setCellValue(cells[["8.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                           (raw_data_baseline_arm_1$tbip_sc_q5=='Empilweni Gompo CHC') & (raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes'))))


setCellValue(cells[["9.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                           (raw_data_baseline_arm_1$tbip_sc_q5=='Pefferville Clinic') & (raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes'))))


setCellValue(cells[["10.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$tbip_sc_q5=='Duncan Village CHC') & (raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes'))))


setCellValue(cells[["11.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$tbip_sc_q5=='Gompo C Jabavu Clinic') & (raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes'))))


setCellValue(cells[["12.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$tbip_sc_q5=='Chris Hani Clinic') & (raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes'))))


setCellValue(cells[["13.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$tbip_sc_q5=='Luyolo NU 9 Clinic') & (raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes'))))


setCellValue(cells[["14.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$tbip_sc_q5=='Alphendale Clinic') & (raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes'))))


setCellValue(cells[["15.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$tbip_sc_q5=='John Dube Clinic') & (raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes'))))


setCellValue(cells[["16.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$tbip_sc_q5=='Fezeka NU 3 Clinic') & (raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes'))))


setCellValue(cells[["17.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$tbip_sc_q5=='Gompo A Ndende Clinic') & (raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes'))))


setCellValue(cells[["18.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$tbip_sc_q5=='Ndevana Clinic') & (raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes'))))


setCellValue(cells[["19.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$tbip_sc_q5=='Philani NU 1 Clinic') & (raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes'))))


setCellValue(cells[["20.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$tbip_sc_q5=='Aspiranza Clinic') & (raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes'))))


setCellValue(cells[["21.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$tbip_sc_q5=='Ginsberg Clinic') & (raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes'))))


setCellValue(cells[["22.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$tbip_sc_q5=='Zwelitsha Zone 5 Clinic') & (raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes'))))


setCellValue(cells[["23.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$tbip_sc_q5=='Masakhane Clinic (Zwelitsha)') & (raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes'))))


setCellValue(cells[["24.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$tbip_sc_q5=='Gompo B Jwayi Clinic') & (raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes'))))


setCellValue(cells[["25.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$tbip_sc_q5=='NU 12 Clinici') & (raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes'))))


######totals###########
setCellValue(cells[["26.3"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$approach_date >= Sys.Date() - 7 & 
                                           (!is.na(raw_data_baseline_arm_1$approach_facility)))))

setCellValue(cells[["26.4"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$approach_date >= Sys.Date() - 7 & 
                                            (!is.na(raw_data_baseline_arm_1$tbip_sc_q5)))))

setCellValue(cells[["26.5"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$approach_date >= Sys.Date() - 7 &
                                            (raw_data_baseline_arm_1$tbip_sc_eligible=='Proceed'))))


setCellValue(cells[["26.6"]], nrow(subset(raw_data_baseline_arm_1, raw_data_baseline_arm_1$tbip_sc_date >= Sys.Date() - 7 & 
                                            (raw_data_baseline_arm_1$tbip_sc_consent_part=='Yes'))))


#xlsx::forceFormulaRefresh(filename_new)
xlsx::saveWorkbook(wb1, filename_new)



