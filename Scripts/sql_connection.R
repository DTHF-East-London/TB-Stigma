library(DBI)
library(RMySQL)
library(RMariaDB)

db_connection <- dbConnect(
  drv = RMySQL::MySQL(),
  dbname = "stigma",
  host = "localhost",
  username = "root",
  password = "P@55word"
)

print(db_connection)


# Replace 'your_table_name' with the actual table name in your database
#dbWriteTable(conn = db_connection, name = 'your_table_name', value = my_data_frame, append = TRUE)

dbSendQuery(db_connection, "SET GLOBAL local_infile = true;")


dbWriteTable(conn = db_connection, name = "follow_up_data", value = follow_up_data, overwrite = TRUE)



###########################################Back Up#####################################################################
#################Baseline
ip_approach <- raw_data_baseline_arm_1[0:15]

ip_screening_and_consent <- raw_data_baseline_arm_1[16:112]

ip_locator <- raw_data_baseline_arm_1[113:119]

ip_demographics <- raw_data_baseline_arm_1[120:169]

ip_questionnaire_1 <- raw_data_baseline_arm_1[170:449]

ip_questionnaire_2 <- raw_data_baseline_arm_1[450:627]

ip_questionnaire_3 <- raw_data_baseline_arm_1[628:822]

ip_doc_upload <- raw_data_baseline_arm_1[823:833]

ip_tb_results <- raw_data_baseline_arm_1[834:885]

ip_hiv_test_results <- raw_data_baseline_arm_1[886:900]

ip_snack_reimbursement <- raw_data_baseline_arm_1[901:909]

ip_study_note <- raw_data_baseline_arm_1[910:917]

ip_qc <- raw_data_baseline_arm_1[918:931]

ip_tb_adherence <- raw_data_baseline_arm_1[932:1013]






dbWriteTable(conn = db_connection, name = "ip_approach", value = ip_approach, overwrite = TRUE)

dbWriteTable(conn = db_connection, name = "index_screening_and_consent", value = ip_screening_and_consent, overwrite = TRUE)

dbWriteTable(conn = db_connection, name = "index_locator", value = ip_locator, overwrite = TRUE)

dbWriteTable(conn = db_connection, name = "index_demographics", value = ip_demographics, overwrite = TRUE)

dbWriteTable(conn = db_connection, name = "index_questionnaire_1", value = ip_questionnaire_1, overwrite = TRUE)

dbWriteTable(conn = db_connection, name = "index_questionnaire_2", value = ip_questionnaire_2, overwrite = TRUE)

dbWriteTable(conn = db_connection, name = "index_questionnaire_3", value = ip_questionnaire_3, overwrite = TRUE)

dbWriteTable(conn = db_connection, name = "index_doc_upload", value = ip_doc_upload, overwrite = TRUE)

dbWriteTable(conn = db_connection, name = "index_tb_results", value = ip_tb_results, overwrite = TRUE)

dbWriteTable(conn = db_connection, name = "index_hiv_test_results", value = ip_hiv_test_results, overwrite = TRUE)

dbWriteTable(conn = db_connection, name = "index_snack_reimbursement", value = ip_snack_reimbursement, overwrite = TRUE)

dbWriteTable(conn = db_connection, name = "index_study_note", value = ip_study_note, overwrite = TRUE)

dbWriteTable(conn = db_connection, name = "index_qc", value = ip_qc, overwrite = TRUE)

dbWriteTable(conn = db_connection, name = "ip_tb_adherence", value = ip_tb_adherence, overwrite = TRUE)


###########################HHCi Info
hhc_list <- raw_data_hhci_info_arm_1[0:21]

hhc_tb_results <- raw_data_hhci_info_arm_1[22:73]

hhc_screening_and_consent <- raw_data_hhci_info_arm_1[74:139]

hhc_demographics <- raw_data_hhci_info_arm_1[140:189]

hhc_questionnaire_1 <- raw_data_hhci_info_arm_1[190:428]

hhc_questionnaire_2 <- raw_data_hhci_info_arm_1[429:606]

hhc_questionnaire_3 <- raw_data_hhci_info_arm_1[607:787]

hhc_snack_reimbursement <- raw_data_hhci_info_arm_1[788:796]

hhc_study_note <- raw_data_hhci_info_arm_1[797:805]

hhc_qc <-raw_data_hhci_info_arm_1[806:815]

hhc_presentation_calling <- raw_data_hhci_info_arm_1[816:852]

hhc_clinic_data_ex <- raw_data_hhci_info_arm_1[853:867]

hhc_termination_form <- raw_data_hhci_info_arm_1[868:876]


dbWriteTable(conn = db_connection, name = "hhc_list", value = hhc_list, overwrite = TRUE)

dbWriteTable(conn = db_connection, name = "hhc_tb_results", value = hhc_tb_results, overwrite = TRUE)

dbWriteTable(conn = db_connection, name = "hhc_screening_and_consent", value = hhc_screening_and_consent, overwrite = TRUE)

dbWriteTable(conn = db_connection, name = "hhc_questionnaire_1", value = hhc_questionnaire_1, overwrite = TRUE)

dbWriteTable(conn = db_connection, name = "hhc_questionnaire_2", value = hhc_questionnaire_2, overwrite = TRUE)

dbWriteTable(conn = db_connection, name = "hhc_questionnaire_3", value = hhc_questionnaire_3, overwrite = TRUE)

dbWriteTable(conn = db_connection, name = "hhc_snack_reimbursement", value = hhc_snack_reimbursement, overwrite = TRUE)

dbWriteTable(conn = db_connection, name = "hhc_study_note", value = hhc_study_note, overwrite = TRUE)

dbWriteTable(conn = db_connection, name = "hhc_qc", value = hhc_qc, overwrite = TRUE)

dbWriteTable(conn = db_connection, name = "hhc_presentation_calling", value = hhc_presentation_calling, overwrite = TRUE)

dbWriteTable(conn = db_connection, name = "hhc_clinic_data_ex", value = hhc_clinic_data_ex, overwrite = TRUE)

dbWriteTable(conn = db_connection, name = "hhc_termination_form", value = hhc_termination_form, overwrite = TRUE)




####################################################################Follow Up 1
fu_doc_upload <- raw_data_follow_up_1_arm_1[0:15]

fu_tb_results <- raw_data_follow_up_1_arm_1[16:67]

fu_questionnaire_1 <- raw_data_follow_up_1_arm_1[68:164]

fu_questionnaire_2 <- raw_data_follow_up_1_arm_1[165:283]

fu_questionnaire_3 <- raw_data_follow_up_1_arm_1[284:478]

fu_snack_reimbursement <- raw_data_follow_up_1_arm_1[479:487]

fu_qc <- raw_data_follow_up_1_arm_1[488:495]

fu_study_note <- raw_data_follow_up_1_arm_1[496:503]


dbWriteTable(conn = db_connection, name = "fu_doc_upload", value = fu_doc_upload, overwrite = TRUE)

dbWriteTable(conn = db_connection, name = "fu_tb_results", value = fu_tb_results, overwrite = TRUE)

dbWriteTable(conn = db_connection, name = "fu_questionnaire_1", value = fu_questionnaire_1, overwrite = TRUE)

dbWriteTable(conn = db_connection, name = "fu_questionnaire_2", value = fu_questionnaire_2, overwrite = TRUE)

dbWriteTable(conn = db_connection, name = "fu_questionnaire_3", value = fu_questionnaire_3, overwrite = TRUE)

dbWriteTable(conn = db_connection, name = "fu_snack_reimbursement", value = fu_snack_reimbursement, overwrite = TRUE)

dbWriteTable(conn = db_connection, name = "fu_qc", value = fu_qc, overwrite = TRUE)

dbWriteTable(conn = db_connection, name = "fu_study_note", value = fu_study_note, overwrite = TRUE)


#############################################HHCI Visit Info

