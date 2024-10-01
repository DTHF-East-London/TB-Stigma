library(furniture)
library(dplyr)
library(ggplot2)

temp1 <- raw_data_baseline_arm_1[c("record_id", "tbip_sc_date", "tbip_sc_q5", "tbip_sc_consent_part", "tbip_sc_cgiver_permission", "tbip_sc_date_visit_hh1")]
temp2 <- raw_data_hhci_visit_info_arm_1[c("record_id", "hhc_sch_hhi_date_visit")]
temp2 <- subset(temp2, !is.na(temp2$hhc_sch_hhi_date_visit))
temp3 <- raw_data_hhci_info_arm_1[c("record_id", "hhc_sc_pin_calc", "hhc_sc_visit_attempt___1", "hhc_sc_attempt_1_date", "hhc_sc_attempt_1_present", "hhc_sc_visit_attempt___2", "hhc_sc_attempt_2_date", "hhc_sc_attempt_2_present", "hhc_sc_attempt_1_rec_date", "hhc_sc_visit_attempt___3", "hhc_sc_attempt_3_date", "hhc_sc_attempt_3_present", "hhc_sc_attempt_1_rec_date_2", "hhc_sc_consent_provided")]
temp <- left_join(temp1, temp2)
temp <- right_join(temp, temp3)

temp <- temp %>% mutate(scheduled_visit_date = case_when(!is.na(hhc_sch_hhi_date_visit) ~ hhc_sch_hhi_date_visit,
                                                    .default = tbip_sc_date_visit_hh1)) %>% relocate(scheduled_visit_date, .after = hhc_sch_hhi_date_visit)

temp <- temp %>% mutate(days_visit_1 = difftime(scheduled_visit_date, tbip_sc_date, units = "days")) %>% relocate(days_visit_1, .after = tbip_sc_date_visit_hh1)
temp <- temp %>% mutate(days_attempt_1 = difftime(hhc_sc_attempt_1_date, tbip_sc_date, units = "days")) %>% relocate(days_attempt_1, .after = hhc_sc_attempt_1_date)
temp <- temp %>% mutate(days_attempt_2 = difftime(hhc_sc_attempt_2_date, tbip_sc_date, units = "days")) %>% relocate(days_attempt_2, .after = hhc_sc_attempt_2_date)
temp <- temp %>% mutate(days_attempt_3 = difftime(hhc_sc_attempt_3_date, tbip_sc_date, units = "days")) %>% relocate(days_attempt_3, .after = hhc_sc_attempt_3_date)

#Visit adherence
temp <- temp %>% mutate(visit_adherence_1 = case_when(difftime(scheduled_visit_date, hhc_sc_attempt_1_date, units = "days")==0 ~ "Yes",
                                                 .default = "No")) %>% relocate(visit_adherence_1, .after = hhc_sc_attempt_1_date)

temp <- temp %>% mutate(visit_adherence_2 = case_when(difftime(hhc_sc_attempt_1_rec_date, hhc_sc_attempt_2_date, units = "days")==0 ~ "Yes",
                                                      .default = "No")) %>% relocate(visit_adherence_2, .after = hhc_sc_attempt_2_date)

temp <- temp %>% mutate(visit_adherence_3 = case_when(difftime(hhc_sc_attempt_1_rec_date_2, hhc_sc_attempt_3_date, units = "days")==0 ~ "Yes",
                                                      .default = "No")) %>% relocate(visit_adherence_3, .after = hhc_sc_attempt_3_date)

temp <- subset(temp, temp$tbip_sc_consent_part=="Yes" & days_attempt_1>0)

# Histograms of days between enrollment and 1st scheduled HHC visited date + mean, median, mode, range

ggplot(temp, aes(tbip_sc_q5, days_visit_1)) + 
  geom_boxplot() +
  scale_x_discrete(guide = guide_axis(angle = 45))+
  labs(title = "Days between Enrolment Date & Schedule Visit Date \n by clinic", y = "Days", x = "Clinics")+
  theme(plot.title = element_text(hjust = 0.5))

mean(temp$days_visit_1, na.rm = FALSE)

median(temp$days_visit_1, na.rm = FALSE)

find_mode(temp$days_visit_1)

max(temp$days_visit_1)

temp_1 <- subset(temp, temp$days_visit_1>0 & days_visit_1<=15)

temp_1 <- temp_1 %>% distinct(record_id, .keep_all = TRUE)

mean(temp_1$days_visit_1)

ggplot(temp_1, aes(tbip_sc_q5, days_visit_1)) + 
  geom_boxplot() +
  scale_x_discrete(guide = guide_axis(angle = 45))+
  labs(title = "Days between Enrolment Date & Schedule Visit Date \n by clinic", y = "Days", x = "Clinics")+
  theme(plot.title = element_text(hjust = 0.5))

ggplot(temp_1, aes(as.integer(days_visit_1))) + 
  geom_bar(width = .8) +
  geom_text(aes(label=after_stat(count)), 
            stat = 'count', 
            color="blue",
            nudge_y = 6,
            size=3.5) +
  scale_x_continuous(breaks = seq(0,40,by=1)) +
  labs(title = "Days between Enrolment Date & Schedule Visit Date \n by clinic", y = "Number of Households", x = "Days") + 
  theme(plot.title = element_text(hjust = 0.5), panel.grid.major.x = element_blank(), panel.grid.minor.x = element_blank(), panel.grid.major.y = element_blank())
table1::table1(~ record_id | tbip_sc_q5, data=temp)

furniture::table1(temp, days_visit_1)

# Histogram of days between enrolment and actual HHC 1st visit + mean, median,mode, range
temp_1 <- subset(temp, temp$days_attempt_1>0)

mean(temp_1$days_attempt_1, na.rm = FALSE)

median(temp_1$days_attempt_1, na.rm = FALSE)

min(temp_1$days_attempt_1, na.rm = FALSE)

max(temp_1$days_attempt_1, na.rm = FALSE)

find_mode(temp_1$days_attempt_1)

getmode(temp_1$days_attempt_1)

temp_1 <- subset(temp, temp$days_attempt_1>0 & days_attempt_1<=15)

mean(temp_1$days_visit_1)

ggplot(temp, aes(tbip_sc_q5, days_attempt_1)) + 
  geom_boxplot() +
  scale_x_discrete(guide = guide_axis(angle = 45))+
  labs(title = "Days between Enrolment Date & First Visit \n by clinic", y = "Days", x = "Clinics")+
  theme(plot.title = element_text(hjust = 0.5))

temp_1 <- subset(temp, temp$days_attempt_1>0 & temp$days_attempt_1<=15)

ggplot(temp_1, aes(tbip_sc_q5, days_attempt_1)) + 
  geom_boxplot() +
  scale_x_discrete(guide = guide_axis(angle = 45))+
  labs(title = "Days between Enrolment Date & First Visit \n by clinic", y = "Days", x = "Clinics")+
  theme(plot.title = element_text(hjust = 0.5))

temp_1 <- temp_1 %>% distinct(record_id, .keep_all = TRUE)

ggplot(temp_1, aes(as.integer(days_visit_1))) + 
  geom_bar(width = .8) +
  geom_text(aes(label=after_stat(count)), 
            stat = 'count', 
            color="black",
            nudge_y = 4,
            size=2.5) +
  scale_x_continuous(breaks = seq(0,62,by=1)) +
  labs(title = "Days between Enrolment Date & First Visit \n by clinic", y = "Number of Households", x = "Days") + 
  theme(plot.title = element_text(hjust = 0.5), panel.grid.major.x = element_blank(), panel.grid.minor.x = element_blank(), panel.grid.major.y = element_blank())
table1::table1(~ record_id | tbip_sc_q5, data=temp)

ggplot(temp, aes(days_attempt_1)) + geom_histogram(binwidth = 1)

temp_1 <- temp %>% distinct(record_id, .keep_all = TRUE)

#-1st visit adherence rate
furniture::table1(temp_1, hhc_sc_visit_attempt___1, splitby = ~ visit_adherence_1,
                  row_wise = TRUE)
#-2nd visit adherence rate
furniture::table1(temp_1, visit_adherence_2, splitby = ~ hhc_sc_visit_attempt___2,
                  row_wise = TRUE)

#- 3rd visit adherence rate
furniture::table1(temp_1, visit_adherence_3, splitby = ~ hhc_sc_visit_attempt___3,
                  row_wise = TRUE)

#- Breakdown of reasons why not want to participate (n=19)

ggplot(temp, aes(tbip_sc_q5, days_attempt_1)) + geom_boxplot()

household_stats <- as.data.frame(table(temp$record_id))

household_stats$Var1 <- as.numeric(household_stats$Var1)

names(household_stats) <- c("record_id", "num_of_hhc")

ggplot(temp, aes(days_attempt_1)) + geom_histogram(binwidth = 1)

household_with_symptomatic_hhc <- subset(temp, temp$hhc_sc_consent_provided=="Yes")

household_stats <- left_join(household_stats, household_with_symptomatic_hhc)

household_stats <- household_stats[c("record_id", "num_of_hhc", "hhc_sc_consent_provided")]

record_facility <- raw_data_baseline_arm_1[c("record_id", "tbip_sc_q5")]

household_stats <- left_join(household_stats, record_facility)

ggplot(household_stats, aes(num_of_hhc, days_attempt_1)) + geom_col()

household_stats_sympt <- subset(household_stats, household_stats$hhc_sc_consent_provided=="Yes")

ggplot(household_stats_sympt, aes(tbip_sc_q5)) + geom_bar()


household_enrolled_by_clinic <- as.data.frame(table(temp$tbip_sc_q5))

household_sympt_by_clinic <- as.data.frame(table(temp$hhc_sc_consent_provided, temp$tbip_sc_q5))

household_sympt_by_clinic <- subset(household_sympt_by_clinic, household_sympt_by_clinic$hhc_sc_consent_provided=="Yes")

household_sympt_b

