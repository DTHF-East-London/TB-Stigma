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

dbWriteTable(conn = db_connection, name = "gxp_returned", value = gxp_returned, append = TRUE)

dbWriteTable(conn = db_connection, name = "raw_data_adhoc_arm_1", value = raw_data_adhoc_arm_1, append = TRUE)

dbWriteTable(conn = db_connection, name = "raw_data_baseline_ni_arm_1", value = raw_data_baseline_ni_arm_1, append = TRUE)

dbWriteTable(conn = db_connection, name = "follow_up_data", value = follow_up_data, append = TRUE)



