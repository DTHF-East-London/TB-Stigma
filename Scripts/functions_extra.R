getTBExtractionRecords <- function(dataset){
  temp <- dataset[c(1,58,65:68,752:780,781:793)]
  View(temp)
  temp <- subset(temp, hhc_pt_days_to_present>30)
  return(temp)
}
