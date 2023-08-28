getTBExtractionRecords <- function(dataset){
  temp <- dataset[c(1,58,65:68,752:780,781:793)]
  temp <- subset(temp, !is.na(temp$hhc_pc_days_since_referral) & temp$hhc_pc_days_since_referral>30)
  return(temp)
}
raw_data_hhci_info_arm_1$hhc_sc_pin_calc