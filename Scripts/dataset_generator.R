source("Scripts/sti_functions.R")

#Get REDCap connection
print("getting REDCap connection")
rcon <- getREDCapConnection()
path <- "./Data/"
output_file <- paste0('dataset',format(Sys.time(), '%d_%B_%Y'),'.xlsx')