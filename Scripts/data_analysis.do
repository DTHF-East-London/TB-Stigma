clear all

cd "C:\Users\freedomm\OneDrive - foundation.co.za\Documents\Projects\TB Stigma\TB-Stigma"

import delimited "Data\full_dataset.csv", bindquote(strict)

*Add labels
do "Scripts\add_labels.do"

*Create value labels
do "Scripts\create_labels.do"

*Assign value labels
do "Scripts\assign_labels.do"

*Convert strings to dates
do "Scripts\convert_dates.do"


********************************************************************************
*                          Perceived TB Stigma                                 *
********************************************************************************

**1. Evaluate the performance of the stigma items**
*this output evaluates the distribution of response items for all the surveys
tab1 s6_q1 - s6_q10 if did_the_person_consent_to==1 & oversampled==0, missing

*Repeat the above tabulation by language
*sort ques_language
tab2 (s6_q1 - s6_q10) ques_language if did_the_person_consent_to==1& oversampled==0, missing

*By how the survey was completed (participant or research assistant)
tab2 (s6_q1 - s6_q10) instruct_part if  did_the_person_consent_to==0 , missing

*this code generates a variable for the number of missing items
egen tbstigma_miss=rowmiss(s6_q1 - s6_q10)
tab tbstigma_miss if did_the_person_consent_to==1 & oversampled==0, missing

*if there are observations with missing items, repeat the above tabulation by language 
tab tbstigma_miss ques_language if did_the_person_consent_to==1 & oversampled==0, missing

*by how the survey was completed (participant or research assistant)
tab tbstigma_miss instruct_part if did_the_person_consent_to==1 & oversampled==0, missing
*tab tbstigma_miss if did_the_person_consent_to==1 & instruct_part==1, missing nolabel
*tab tbstigma_miss if did_the_person_consent_to==1 & instruct_part==0, missing

*this output evaluates the Cronbach alpha, a measure of the scale performance
alpha s6_q1 - s6_q10 if did_the_person_consent_to==1 & oversampled==0, item

*repeat the above output by language 
alpha s6_q1 - s6_q10 if did_the_person_consent_to==1 & oversampled==0, item
*by how the survey was completed (participant or research assistant)

**2. Create stigma scores **
*Create a sum score for all participants with no missing items
egen tbstigma_sum=rowtotal(s6_q1 - s6_q10) if tbstigma_miss==0

*Create a mean item score for all participants with <25% missing items
egen tbstigma_mean=rowmean(s6_q1 - s6_q10) if tbstigma_miss<=2 

*I do not provide any code here for calculating the community-level stigma score as that involves use of other dataset variables besides just the stigma variables.

bysort area_1 study_community_1: egen community_tb_obs=count(tbstigma_mean)

*e.	Create a community level stigma score 
bysort area_1 study_community_1: egen community_tbstigma_score=mean(tbstigma_mean)

bysort area_1 study_community_1: egen community_tb_sd=sd(tbstigma_mean)

bysort area_1 study_community_1: egen community_tb_min=min(tbstigma_mean)

bysort area_1 study_community_1: egen community_tb_max=max(tbstigma_mean)

*Create a new column called first and set the first row for each community to 1
bysort area_1 study_community_1: gen first=1 if _n==1

*Create community_stats (subset)
frame put area_1  study_community_1 community_tb_obs community_tbstigma_score community_tb_sd community_tb_min community_tb_max if first==1, into(community_stats)

*Switch to community_stats frame
frame change community_stats

*Save data to an excel file
export excel using "C:\Users\freedomm\OneDrive - foundation.co.za\Documents\Projects\TB Stigma\TB-Stigma\Data\Area_Community_TB_Stigma_Scores.xls", sheet("Community TB Stigma Scores") sheetreplace firstrow(variables)

bysort area_1: egen area_obs=count(community_tbstigma_score)

*Create an area level stigma score 
bysort area_1: egen area_tbstigma_score=mean(community_tbstigma_score)

bysort area_1: egen area_sd=sd(community_tbstigma_score)

bysort area_1: egen area_min=min(community_tbstigma_score)

bysort area_1: egen area_max=max(community_tbstigma_score)

*Create a new column called first and set the first row for each community to 1
bysort area_1: gen first=1 if _n==1

*Create community_stats (subset)
frame put area_1 area_obs area_tbstigma_score area_sd area_min area_max if first==1, into(area_stats)

*Switch to area_stats frame
frame change area_stats

export excel using "C:\Users\freedomm\OneDrive - foundation.co.za\Documents\Projects\TB Stigma\TB-Stigma\Data\Area_Community_TB_Stigma_Scores.xls", sheet("Area TB Stigma Scores") sheetreplace firstrow(variables)

********************************************************************************
*                          Perceived HIV Stigma                                *
********************************************************************************

frame change default

*item s5_q6_p needs to be reverse coded on the HIV stigma scale
generate s5_q6_p_rev=(s5_q6_p-3)*(-1)

*double check that if originally coded 3, now 0; if originally coded 2, now 1, etc
tab s5_q6_p s5_q6_p_rev

*all calculations of HIV stigma scores need to use s5_q6_p_rev instead of the original s5_q6_p



**1. Evaluate the performance of the stigma items**
*this output evaluates the distribution of response items for all the surveys
tab1 s5_q1_p - s5_q5_p s5_q6_p_rev s5_q7_p - s5_q14_p if did_the_person_consent_to==1  & oversampled==0, missing

*Repeat the above tabulation by language
sort ques_language
tab2 (s5_q1_p - s5_q5_p s5_q6_p_rev s5_q7_p - s5_q14_p) ques_language if  did_the_person_consent_to==1 & oversampled==0, missing

*By how the survey was completed (participant or research assistant)
tab2 (s5_q1_p - s5_q5_p s5_q6_p_rev s5_q7_p - s5_q14_p) instruct_part if  did_the_person_consent_to==1 & oversampled==0, missing

*this code generates a variable for the number of missing items
egen hivstigma_miss=rowmiss(s5_q1_p - s5_q5_p s5_q6_p_rev s5_q7_p - s5_q14_p)

tab hivstigma_miss if did_the_person_consent_to==1 & oversampled==0, missing

*if there are observations with missing items, repeat the above tabulation by language 
tab hivstigma_miss ques_language if did_the_person_consent_to==1 & oversampled==0, missing

*by how the survey was completed (participant or research assistant)
tab hivstigma_miss instruct_part if did_the_person_consent_to==1 & oversampled==0, missing
tab hivstigma_miss if did_the_person_consent_to==1 & instruct_part==1 & oversampled==0, missing nolabel
tab hivstigma_miss if did_the_person_consent_to==1 & instruct_part==0 & oversampled==0, missing

*this output evaluates the Cronbach alpha, a measure of the scale performance
alpha s5_q1_p - s5_q5_p s5_q6_p_rev s5_q7_p - s5_q14_p if did_the_person_consent_to==1 & oversampled==0, item

*repeat the above output by language 
alpha s5_q1_p - s5_q5_p s5_q6_p_rev s5_q7_p - s5_q14_p if did_the_person_consent_to==1 & oversampled==0, item

*by how the survey was completed (participant or research assistant) gor

**2. Create stigma scores **
*Create a sum score for all participants with no missing items
egen hivstigma_sum=rowtotal(s5_q1_p - s5_q5_p s5_q6_p_rev s5_q7_p - s5_q14_p) if hivstigma_miss==0

*Create a mean item score for all participants with <25% missing items
egen hivstigma_mean=rowmean(s5_q1_p - s5_q5_p s5_q6_p_rev s5_q7_p - s5_q14_p) if hivstigma_miss<=2 

*I do not provide any code here for calculating the community-level stigma score as that involves use of other dataset variables besides just the stigma variables.


*3.	Report stigma scores
*e.	Create a community level stigma score 
bysort area_1 study_community_1: egen community_hiv_obs=count(hivstigma_mean)

bysort area_1 study_community_1: egen community_hivstigma_score=mean(hivstigma_mean)

bysort area_1 study_community_1: egen community_hiv_sd=sd(hivstigma_mean)

bysort area_1 study_community_1: egen community_hiv_min=min(hivstigma_mean)

bysort area_1 study_community_1: egen community_hiv_max=max(hivstigma_mean)

*Create a new column called first and set the first row for each community to 1
*bysort area_1 study_community_1: gen first=1 if _n==1

*drop frame community_stats
frame drop community_stats

*Create community_stats (subset)
frame put area_1  study_community_1 community_hiv_obs community_hivstigma_score community_hiv_sd community_hiv_min community_hiv_max if first==1, into(community_stats)

*Switch to community_stats frame
frame change community_stats

*Save data to an excel file
export excel using "C:\Users\freedomm\OneDrive - foundation.co.za\Documents\Projects\TB Stigma\TB-Stigma\Data\Area_Community_TB_Stigma_Scores.xls", sheet("Community HIV Stigma Scores") sheetreplace firstrow(variables)

bysort area_1: egen area_obs=count(community_hivstigma_score)

*Create an area level stigma score 
bysort area_1: egen area_hivstigma_score=mean(community_hivstigma_score)

bysort area_1: egen area_sd=sd(community_hivstigma_score)

bysort area_1: egen area_min=min(community_hivstigma_score)

bysort area_1: egen area_max=max(community_hivstigma_score)

*Create a new column called first and set the first row for each community to 1
bysort area_1: gen first=1 if _n==1

frame drop area_stats

*Create community_stats (subset)
frame put area_1 area_obs area_hivstigma_score area_sd area_min area_max if first==1, into(area_stats)

*Switch to area_stats frame
frame change area_stats

export excel using "C:\Users\freedomm\OneDrive - foundation.co.za\Documents\Projects\TB Stigma\TB-Stigma\Data\Area_Community_TB_Stigma_Scores.xls", sheet("Area HIV Stigma Scores") sheetreplace firstrow(variables)