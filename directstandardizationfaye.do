/*
Direct standardization lets us remove the distortion caused by the different ${agevar} distributions. 
The adjusted rate is defined as the weighted sum of the crude rates, where the weights are given by the standard distribution. 
This syntax uses the Faye and Feuer (1997) method to calculate CIs for the adjusted rates
*/
************************************************************************************************************************************
clear all
clear results
clear matrix
cap net install distinct.pkg, from(http://fmwww.bc.edu/RePEc/bocode/d)
cap net install distrate.pkg, from(http://fmwww.bc.edu/RePEc/bocode/d)

qui {
capture log close
cap erase quiet_noise.log
log using quiet_noise.log, text replace
************************************************************************************************************************************
/*INPUTS*/
cd "C:/" /*working directory which contains all the required datasets*/

global standardfilename "Population_WEC_ON_CAN_2007_2017" /*file which contains the 2011 standard population data*/
global standardfilesheetname "can_2007-2015_cen2011"

global popestimatesfilename "Population_WEC_ON_CAN_2007_2017" /*file which contains the population estimates data*/
global popestimatessheetname "onphus_2007-2016_cen2011"

global popprojectionsfilename "Population_WEC_ON_CAN_2007_2017" /*file which contains the 2017 population projections data*/
global popprojectionssheetname "onphus_2017_cen2011"

global casefilename "Book5.xlsx"
global casesheetname  "Line List_2"
global casesheetnamecellrange "A4:O399189" /*range in sheet which contains the case data*/
global uniqueid nacrskey /*name of unique identifier in case dataset*/
global hn hn /*name of hn variable (if using intellihealth data and there is a hn in the dataset). otherwise, leave as blank*/
global othercasevartokeep "hn nacrskey date alldx*"

global geovar geo /*name of variable must be consistent across all datasets*/
global yearvar year /*name of variable must be consistent across all datasets*/
global agevar age /*name of variable must be consistent across all datasets*/
global sexvar sex /*name of variable must be consistent across all datasets*/
global popvar pop /*name of variable must be consistent across all datasets*/

global yearstart = 2010 /*start year of interest*/ 
global yearend = 2017 /*last  year of interest*/
global agestart = 1 /*start age of interest*/
global ageend = 151 /*last age of interest*/
global agegroups 0,1,5[5]15,18,20[5]90,151  /*Translates to: 0,1-4,5-9,10-14,15-17,18-19,20-24,25-29...90-151*/

global localphu "WIND" /*First four chars of PHU name in CAPS*/
global rateper 100000 /*report rates per: */
global rateformat %8.1f /*How to display the adjusted rate.*/
global cilevel 95 /*Confidence interval to use*/
************************************************************************************************************************************
noi: cd
noi: di "Standard population file: ${standardfilename}"
noi: di "Sheet: ${standardfilesheetname}"
noi: di ""
noi: di "Population Estimates file: ${popestimatesfilename}"
noi: di "Sheet: ${popestimatessheetname}"
noi: di ""
noi: di "Population Projections file: ${popprojectionsfilename}"
noi: di "Sheet: ${popprojectionssheetname}"
noi: di ""
noi: di "Case data file: ${casefilename};"
noi: di "Sheet: ${casesheetname};	 Cell range: ${casesheetnamecellrange}"
noi: di "Unique identifier variable: ${uniqueid}"
noi: di "Other variabless to keep: ${othercasevartokeep}"
noi: di ""
noi: di "Geographic identifer: ${geovar}; 	Year variable: ${yearvar}"
noi: di "Age variable: ${agevar}; 	Sex variable: ${sexvar}; 	Population variable: ${popvar}"
di
noi: di "Years of interest (inclusive): ${yearstart} to ${yearend}"
noi: di "Age range of interest (inclusive): ${agestart} to ${ageend}"
noi: di "Age groups for standardization: ${agegroups}"
noi: di ""
noi: di "Local geography/phu of interest: ${localphu}" 
noi: di "Comparator geography: Ontario"
noi: di "Rates reported per" %9.0fc ${rateper} " population"
noi: di "Confidence Interval Level: " ${cilevel} "%" "
noi: di ""
************************************************************************************************************************************
capture program drop keepobs
program define keepobs
	args yearstart yearend agestart ageend agevar sexvar
	keep if inrange(${yearvar},${yearstart},${yearend}) 
	keep if inrange(${agevar},${agestart},${ageend}) 
	keep if upper(${sexvar})=="MALE" | upper(${sexvar})=="FEMALE"
end

*A program that defines how the ${agevar} variable should be grouped. age_catcustom is the variable of interest.
capture program drop age_cat
program define age_cat
	args agevar agegroups
	capture drop age_cat*
	egen age_catcustom = cut(${agevar}), at(${agegroups})
	noi: table age_catcustom, contents(min ${agevar} max ${agevar})
end

*A program that collapses the data, summing the population counts by ${yearvar} ${agevar} and ${sexvar}
capture program drop collapsedata
program define collapsedata
	args yearvar sexvar
	collapse (sum) ${popvar}, by(${yearvar} age_catcustom ${sexvar})
end

*A program that fixes the PHU name in the population dataset and the case dataset.
capture program drop fixgeoname
program define fixgeoname
	args geovar
	gen geo2=${geovar}
	replace geo2="NORBAY PARRY SOUND" if strpos(geo2,"NORTH BAY") /*simple work around so that NORTHBAY can be differentiated from NORTHWESTERN once var is trimmed to 4 chars*/
	replace geo2=substr(${geovar},1,4)
end

capture program drop errorvalues
program define errorvalues
cap drop se_gam lb_dob ub_dob
cap drop pev
		gen pev=ub_gam-rateadj
		label var pev "Positive Error Value"
		
		cap drop nev
		gen nev=rateadj-lb_gam
		label var nev "Negative Error Value"
end
************************************************************************************************************************************
*Import and save 2011 Canadian Population data from Census
cap erase pop.dta
import excel "${standardfilename}", sheet("${standardfilesheetname}") firstrow case(lower) clear
	keepobs
	keep if ${yearvar}==2011
	age_cat
	
	li ${yearvar} ${geovar} ${agevar} age_catcustom ${sexvar} ${popvar} in 1/20, sepby(age_catcustom)	
	collapsedata
	sort age_catcustom ${sexvar} /*required for distrate command*/
	li ${yearvar} ${geovar} ${agevar} age_catcustom ${sexvar} ${popvar}, sepby(age_catcustom)	

capture save pop

****************************************************************************************************************************************
*Import and save 2017 Ontario PHU Population Projections from IntelliHealth, and then append it to 2007-2016 Ontario PHU Population Estimates from IntelliHealth
cap erase onpopappended.dta
import excel "${popprojectionsfilename}", sheet("${popprojectionssheetname}") firstrow case(lower) clear
	tempfile onpop
	save `onpop'
	
	import excel "${popestimatesfilename}", sheet("${popestimatessheetname}") firstrow case(lower) clear
	append using `onpop'
	
	fixgeoname
	keepobs
	age_cat
	
	li ${yearvar} geo2 age_catcustom ${sexvar} ${popvar} in 1/20, sepby(${yearvar} geo2) 	
	collapse (sum) ${popvar}, by(${yearvar} geo2 age_catcustom ${sexvar})
	sort ${yearvar} geo2 ${agevar} ${sexvar}
	li ${yearvar} geo2 age_catcustom ${sexvar} ${popvar}  in 1/20, sepby(geo2 ${yearvar})	

capture save onpopappendedv2

****************************************************************************************************************************************
*Import and save NACRS case data from intellihealth
cap erase casedata.dta
import excel "${casefilename}", sheet("${casesheetname}") cellrange("${casesheetnamecellrange}") firstrow case(lower) clear	
	keep ${yearvar} ${agevar} ${sexvar} ${geovar} ${othercasevartokeep}
	
	replace ${geovar} = substr(${geovar},8,.) /*removing first eight characters, the PHU code*/
	fixgeoname 
	keepobs
	age_cat 

	foreach year of numlist $yearstart/$yearend {
		qui: distinct ${uniqueid} if ${yearvar}==`year'
		display `year' "," r(ndistinct)
	}
      
	duplicates report ${uniqueid} ${hn} ${geovar} ${agevar} ${sexvar}
	duplicates tag ${uniqueid} ${hn} ${geovar} ${agevar} ${sexvar}, g(dup)
	
	bysort ${uniqueid}: drop if _n>1 /*dropping duplicate observations*/
	duplicates report ${uniqueid} ${hn} ${geovar} ${agevar} ${sexvar}


	gen cases=1
	
	li ${yearvar} geo2 age_catcustom ${sexvar} cases in 1/50, sepby(geo2)
	collapse (sum) cases, by(${yearvar} geo2 age_catcustom ${sexvar})
	sort ${yearvar} geo2 ${agevar} age_catcustom ${sexvar} 
	li ${yearvar} geo2 age_catcustom ${sexvar} cases, sepby(geo2)
	
	merge 1:1 geo2 ${yearvar} age_catcustom ${sexvar} using onpopappendedv2
	li ${yearvar} geo2 age_catcustom ${sexvar} case, sepby(geo2)
	
	replace cases = 0 if missing(cases)

cap save casedata
****************************************************************************************************************************************
*Perform direct standardization for local PHU
use casedata, clear

	capture erase localrates.dta
	noi: di "Rates for :" "$localphu"
	noi: distrate cases ${popvar} using pop if geo2=="$localphu", standstrata(age_catcustom ${sexvar}) by(${yearvar}) mult(100000) format($rateformat) level(${cilevel}) list(${yearvar} cases N crude rateadj lb_gam ub_gam) saving(localrates, replace)

*Perform direct standardization for Ontario
use casedata, clear

	replace geo2="ONTARIO"

	collapse (sum) cases ${popvar}, by(geo2 ${yearvar} age_catcustom ${sexvar})
	sort geo2 ${yearvar} ${agevar} age_catcustom ${sexvar} 
	li geo2 ${yearvar} age_catcustom ${sexvar} case ${popvar}

	capture erase ontariorates.dta
	noi: di "Rates for Ontario:"
	noi: distrate cases ${popvar} using pop.dta, standstrata(age_catcustom ${sexvar}) by(${yearvar}) mult(${rateper}) format($rateformat) level(${cilevel}) list(${yearvar} cases N crude rateadj lb_gam ub_gam) saving(ontariorates, replace)
****************************************************************************************************************************************
*Export to excel
use localrates, clear
  errorvalues
	export excel using "rates", sheet("rates") cell($localstartcell) firstrow(varlabel) sheetreplace

use ontariorates, clear
  errorvalues
	export excel using "rates", sheet("rates") cell($ontariostartcell) firstrow(varlabel) sheetmodify
****************************************************************************************************************************************
local list pop onpopappended onpopappendedv2 casedata localrates ontariorates 
foreach dataset of local list {
	cap erase "`dataset'.dta"

}
capture log close
cap erase quiet_noise.log

clear all
clear results
clear matrix
}
