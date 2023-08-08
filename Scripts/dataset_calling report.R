# Load required libraries
library(redcapAPI)
library(dplyr)
library(ggplot2)

# Create a connection to Redcap
api_url <- 'https:/redcap11.hiv-research.org.za/api/'
api_token <- "E8B4D8574D98C69F6A3CB157F7491527"
redcap_conn <- redcapConnection(url = api_url, token = api_token)

# Retrieve data from Redcap
data <- exportRecords(redcap_conn, fields = c("record_id", "phone_number", "call_status", "call_notes"))

# Filter data or perform any necessary data manipulation
filtered_data <- data %>%
  filter(call_status == "Completed")

# Generate a summary or visualization
call_summary <- filtered_data %>%
  group_by(call_status) %>%
  summarize(call_count = n())

# Create a bar plot of call counts
ggplot(call_summary, aes(x = call_status, y = call_count)) +
  geom_bar(stat = "identity") +
  labs(title = "Calling Report",
       x = "Call Status",
       y = "Call Count")
