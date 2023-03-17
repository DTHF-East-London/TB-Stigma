tostring labelling_date, replace
gen _date_ = date(labelling_date,"YMD")
drop labelling_date
rename _date_ labelling_date
format labelling_date %dM_d,_CY

tostring screening_date_1, replace
gen _date_ = date(screening_date_1,"YMD")
drop screening_date_1
rename _date_ screening_date_1
format screening_date_1 %dM_d,_CY

tostring screening_date_2, replace
gen _date_ = date(screening_date_2,"YMD")
drop screening_date_2
rename _date_ screening_date_2
*format screening_date_2 %dM_d,_CY

tostring screening_date_3, replace
gen _date_ = date(screening_date_3,"YMD")
drop screening_date_3
rename _date_ screening_date_3
format screening_date_3 %dM_d,_CY

tostring consent_date, replace
gen _date_ = date(consent_date,"YMD")
drop consent_date
rename _date_ consent_date
format consent_date %dM_d,_CY

tostring demographics_date, replace
gen _date_ = date(demographics_date,"YMD")
drop demographics_date
rename _date_ demographics_date
format demographics_date %dM_d,_CY

tostring date4, replace
gen _date_ = date(date4,"YMD")
drop date4
rename _date_ date4
format date4 %dM_d,_CY

*tostring end_time, replace
*gen double _temp_ = Clock(end_time,"YMDhm")
*drop end_time
*rename _temp_ end_time
*format end_time %tCMonth_dd,_CCYY_HH:MM

tostring date2, replace
gen _date_ = date(date2,"YMD")
drop date2
rename _date_ date2
format date2 %dM_d,_CY

*tostring date, replace
*gen _date_ = date(date,"YMD")
*drop date
*rename _date_ date
*format date %dM_d,_CY

tostring date3, replace
gen _date_ = date(date3,"YMD")
drop date3
rename _date_ date3
format date3 %dM_d,_CY