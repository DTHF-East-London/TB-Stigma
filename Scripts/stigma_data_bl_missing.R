#########

library(redcapAPI)

library(dplyr)
#############

#Data
#source("Scripts/functions.R")

# Check for missing values
missing_values <- sum(is.na(raw_data_baseline_arm_1))
missing_values <- sum(is.na(raw_hhci_info_arm_1))

# Display the number of missing values

print(paste("Number of missing values:", missing_values))
