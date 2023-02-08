library(dplyr)
library(tidyverse)
require(openxlsx)
library(xlsx)

source("Scripts/functions.R")

insertRow <- function(existingDF, newrow, r) {
  existingDF[seq(r+1,nrow(existingDF)+1),] <- existingDF[seq(r,nrow(existingDF)),]
  existingDF[r,] <- newrow
  existingDF
}

data_dictionary <- getMetadata(NULL, NULL)

data_dictionary <- data_dictionary %>% mutate(.data = ., Observed = NA, Total = NA, Completeness = NA, Mean = NA, Median = NA, Min = NA, Max =NA) %>% relocate(.data = ., c(Observed, Total, Completeness, Mean, Median, Min, Max), .before = field_name)

data_dictionary <- rbind(c("Observed", "Total", "Completeness", "Mean", "Median", "Min", "Max", "Variable / Field Name",	"Form Name",	"Section Header",	"Field Type",	"Field Label",	"Choices, Calculations, OR Slider Labels",	"Field Note",	"Text Validation Type OR Show Slider Number",	"Text Validation Min",	"Text Validation Max",	"Identifier?",	"Branching Logic (Show field only if...)",	"Required Field?",	"Custom Alignment",	"Question Number (surveys only)",	"Matrix Group Name",	"Matrix Ranking?",	"Field Annotation"), data_dictionary)

fields <- data_dictionary$field_name

data <- getREDCapRecords(NULL,NULL,NULL)

#drop identifiers
#data <- subset(data, select = -c(3,4,5,7,17))

#data$redcap_data_access_group <- data$pi_research_assistant

#data$ca_ra <- data$pi_research_assistant

data <- subset(data, select = -c(2))

checkbox <- data %>% select(contains("___"))

checkbox <- as.data.frame(names(checkbox))

names(checkbox) <- "name"

for(field in checkbox$`name`){
  #name <- gsub(,field)
  #which(data_dictionary$field_name==field)
}

field <- "today_s_date"

for(field in fields){
  print(field)
  i <- which(data_dictionary$field_name==field) 
  if(field!="Variable / Field Name" && is.na(data_dictionary[i, 19])){
    
    if(field %in% names(data)){
      field_data <- data %>% select("record_id", field)
      if(field=='record_id'){
        data_dictionary[i, 1] <- nrow(subset(field_data, !is.na(field_data[,1])))
        data_dictionary[i, 2] <- nrow(field_data)
        data_dictionary[i, 3] <- round(nrow(subset(field_data, !is.na(field_data[,1])))/nrow(field_data)*100,1)
      }else{
        data_dictionary[i, 1] <- nrow(subset(field_data, !is.na(field_data[,2])))
        data_dictionary[i, 2] <- nrow(field_data)
        data_dictionary[i, 3] <- round(nrow(subset(field_data, !is.na(field_data[,2])))/nrow(field_data)*100,1)
        if(!is.na(data_dictionary[i,15]) && (data_dictionary[i,15]=='integer' | data_dictionary[i,15]=='number')){
          data_dictionary[i, 4] <- round(mean(field_data[,2],na.rm = TRUE),1)
          data_dictionary[i, 5] <- round(median(field_data[,2],na.rm = TRUE),1)
          data_dictionary[i, 6] <- min(field_data[,2],na.rm = TRUE)
          data_dictionary[i, 7] <- max(field_data[,2],na.rm = TRUE)
        }
      }
    }
  }else if(field!="Variable / Field Name"){
    branching <- data_dictionary[i,19]
    print(branching)
    branching <- gsub('\\[','data$',branching)
    branching <- gsub('\\]',' ',branching)
    branching <- gsub('and','\\&',branching)
    branching <- gsub(' or ','\\ | ',branching)
    branching <- gsub('\\(','\\___',branching)
    branching <- gsub('\\ ___','\\ (',branching)
    branching <- gsub('\\)]',']',branching)
    branching <- gsub(' =', ' == ', branching)
    branching <- gsub('\\______','\\((',branching)
    branching <- gsub('[>]', 'x', branching)
    branching <- gsub('[<]', 'x', branching)
    branching <- gsub('[""]', 'x',branching)
    branching <- gsub('x=', '>=', branching)
    branching <- gsub('  xxxx','x',branching)
    branching <- gsub('\\)  ==  ',' == ',branching)
    branching <- gsub("or\n", " | ", branching)
    branching <- gsub("  x ", " > ", branching)
    branching <- gsub("\\<data$preferred_language_to_be_ux\\>","is.na(data$preferred_language_to_be_u)", branching)
    if(branching=="data$preferred_language_to_be_ux")
      branching <- "is.na(data$preferred_language_to_be_u)"
    
    if(branching=="data$consent_date xxxx")
      branching <- "is.na(data$consent_data)"
    
    print(branching)
    if(field %in% names(data)){
      field_data <- subset(data, eval(parse(text=paste(branching))))
      field_data <- field_data %>% select("record_id", field)
      if(field=='record_id'){
        data_dictionary[i, 1] <- nrow(subset(field_data, !is.na(field_data[,1])))
        data_dictionary[i, 2] <- nrow(field_data)
        data_dictionary[i, 3] <- round(nrow(subset(field_data, !is.na(field_data[,1])))/nrow(field_data)*100,1)
      }else{
        data_dictionary[i, 1] <- nrow(subset(field_data, !is.na(field_data[,2])))
        data_dictionary[i, 2] <- nrow(field_data)
        data_dictionary[i, 3] <- round(nrow(subset(field_data, !is.na(field_data[,2])))/nrow(field_data)*100,1)
        
        if(!is.na(data_dictionary[i,15]) && (data_dictionary[i,15]=='integer' | data_dictionary[i,15]=='number')){
          data_dictionary[i, 4] <- round(mean(field_data[,2],na.rm = TRUE),1)
          data_dictionary[i, 5] <- round(median(field_data[,2],na.rm = TRUE),1)
          data_dictionary[i, 6] <- min(field_data[,2],na.rm = TRUE)
          data_dictionary[i, 7] <- max(field_data[,2],na.rm = TRUE)
        }
      }
    }
  }
}

library(openxlsx)

#field <- list(data_dictionary$field_name)

file_name_data <- paste("Data/TB Stigma: Household_Raw", Sys.Date(),".xlsx")

writeXlsx(data_dictionary, "DataDictionary", file_name_data, TRUE, FALSE)

writeXlsx(data, "Data", file_name_data, FALSE)

write.csv(data_dictionary, "Data/daily_review.csv", sep = ',', row.names = FALSE)

