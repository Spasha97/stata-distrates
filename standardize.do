/*
Author: Mathew Roy (Created March 27,2018; Updated May 7, 2018)
This syntax uses the Faye and Feuer (1997) method to calculate CIs for the adjusted rates.
*/
****************************************************************************************************************************************
****************************************************************************************************************************************
/* Install packages distrate and distinct*/
// cap net install distrate.pkg, from(http://fmwww.bc.edu/RePEc/bocode/d)
// cap net install distinct.pkg, from(http://fmwww.bc.edu/RePEc/bocode/d) 
qui { 
	clear all
	clear results
	clear matrix

	// The next three lines, along with the qui command above, works to hide any syntax from displaying in the Results window.
	cap log close
	cap erase quiet_noise.log
	log using quiet_noise.log, text replace
	****************************************************************************************************************************************
	****************************************************************************************************************************************
	/*INPUTS*/
	cd "C:\" /* working directory which contains all the required datasets */

	// I chose to assign global macros, so that macro values remain when the code is run in blocks
	global standardfilename "canadianpop" /*Name of file which contains the 2011 standard population data*/
	global standardfilesheetname "canadianpop_linelist"

	global popestimatesfilename "ontariopopestbyphu" /*Name of file which contains the population estimates data*/
	global popestimatessheetname "ontariopopestbyphu_linelist"

	global popprojectionsfilename "" /*Name file which contains the 2017 population projections data; leave blank if N/A*/
	global popprojectionssheetname "" /*leave blank if N/A*/

	global casefilename "hip_fractures_03.xls" /*Filename of Excel spreadsheet with case data*/
	global casesheetname  "linelist" /* Name of worksheet containing the data; leave blank if N/A*/
	global casesheetnamecellrange "A2:N50000" /*range in sheet which contains the case data; leave blank if N/A*/

	global uniqueid dadkey /*name of unique identifier in case dataset*/
	global hn "" /*name of hn variable (if using intellihealth data and there is a hn in the dataset). otherwise, leave as blank*/
	global othercasevartokeep "mrdx* alldx*" /*Variables (besides year, age, geo, uniqueid, hn) to keep when working with the data; leave blank if N/A */

	// Names of the following variables must be consistent across all datasets (case data, canadian standard population data, and ontario population data)
	global geovar geo /*PHU variable*/
	global yearvar year /*Year variable*/
	global agevar age /*Age variable; values should be in in numeric form identifying single age years, or age groups*/
	global sexvar ""  /*Sex variable; Leave blank if you don't want to standardize by sex; value should be in string form "MALE" or  "FEMALE" */
	global popvar pop /*Population count variable*/

	global yearstart = 2010 /*start year of interest*/ 
	global yearend = 2016 /*last  year of interest*/
	global agestart = 0 /*start age of interest*/
	global ageend = 151 /*last age of interest*/
	global agegroups 0[5]90,151  /*Translates to: 0-4,5-9,10-14,15-19,20-24,25-29...90-151*/

	global localphu "WIND" /*First four chars of PHU name in CAPS*/
	global rateper 100000 	/*report rates per: 100,000 population*/
	global rateformat %8.1f /*How to display the adjusted rate.*/
	global cilevel 95 /*Confidence interval to use*/
	global localstartcell "A1" /*Cell in final output Exel spreadhsheet, in which the results of the local analysis should be placed*/
	global ontariostartcell "N1" /*Cell in final output Exel spreadhsheet, in which the results of the provincial analysis should be placed*/
	global finaloutputexcel "rates" /*Name of final output spreadsheet with results*/
	************************************************************************************************************************************
	****************************************************************************************************************************************
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
	
	noi: di "Case data file: ${casefilename}"
	noi: di "Sheet: ${casesheetname}"
	noi: di "Cell range: ${casesheetnamecellrange}"
	noi: di "Unique identifier variable: ${uniqueid}"
	noi: di "Other variabless to keep: ${othercasevartokeep}"
	noi: di ""
	
	noi: di "Geographic identifer: ${geovar}"
	noi: di "Year variable: ${yearvar}"
	noi: di "Age variable: ${agevar}"	
	noi: di "Sex variable: ${sexvar}" 	
	noi: di "Population variable: ${popvar}"
	noi: di ""
	
	noi: di "Years of interest (inclusive): ${yearstart} to ${yearend}"
	noi: di "Age range of interest (inclusive): ${agestart} to ${ageend}"
	noi: di "Age groups for standardization: ${agegroups}"
	noi: di ""
	
	noi: di "Local geography/phu of interest: ${localphu}" 
	noi: di "Comparator geography: Ontario"
	noi: di "Rates reported per" %9.0fc ${rateper} " population"
	noi: di "Confidence Interval Level: " ${cilevel} "%"
	noi: di ""
	************************************************************************************************************************************
	************************************************************************************************************************************
	/*Programs that will be used repeatedly*/

	// Program: Keep only the years groups of interest
	cap program drop keepobs
	program define keepobs
		args yearstart yearend agestart ageend agevar sexvar
		keep if inrange(${yearvar},${yearstart},${yearend}) 
		keep if inrange(${agevar},${agestart},${ageend}) 
		if !missing("${sexvar}") {
			keep if upper(${sexvar})=="MALE" | upper(${sexvar})=="FEMALE"
		}
	end

	// Program: Defines how the age variable should be grouped. age_catcustom is the final grouping variable.
	cap program drop age_cat
	program define age_cat
		args agevar agegroups
		cap drop age_cat*
		egen age_catcustom = cut(${agevar}), at(${agegroups})
		noi: table age_catcustom, contents(min ${agevar} max ${agevar})
	end

	// Program: Changes the PHU name in the population dataset and the case dataset.
	cap program drop fixgeoname
	program define fixgeoname
		args geovar
		replace ${geovar} = substr(${geovar},8,.) if strpos(substr(${geovar},1,1),"(") 
		cap drop geo2
		gen geo2=${geovar}
		replace geo2="NORBAY PARRY SOUND" if strpos(${geovar},"NORTH BAY") /*simple work around so that NORTHBAY can be differentiated from NORTHWESTERN once var is trimmed to 4 chars*/
		replace geo2=substr(geo2,1,4)
	end

	// Program:  Creates positive and negative error values, which will be outputted in the final excel spreadsheet
	cap program drop errorvalues
	program define errorvalues
		cap drop se_gam lb_dob ub_dob
		cap drop pev
		gen pev=ub_gam-rateadj
		label var pev "Positive Error Value"
		
		cap drop nev
		gen nev=rateadj-lb_gam
		label var nev "Negative Error Value"
	end
	****************************************************************************************************************************************
	****************************************************************************************************************************************
	/* Import and save 2011 Canadian Population data from Census */
	cap erase pop.dta
	import excel "${standardfilename}", sheet("${standardfilesheetname}") firstrow case(lower) clear
		
		// Keep only: age groups and/or sex of interest, 2011 Canadian population data, and categorize the data by ages groups of interest
		keepobs
		keep if ${yearvar}==2011
		age_cat
		
		// Aggregate data by the groups of interest
		collapse (sum) ${popvar}, by(${geovar} ${yearvar} age_catcustom ${sexvar}) 
		sort age_catcustom ${sexvar} /*required for distrate command later on*/
		// li ${yearvar} ${geovar} ${agevar} age_catcustom ${sexvar} ${popvar}, sepby(${sexvar})
		// table ${geovar}, c(sum pop) format(%9.0fc)

	save pop
	****************************************************************************************************************************************
	****************************************************************************************************************************************
	/* Import and save Ontario PHU Population Projections and Estimates from IntelliHealth */
	cap erase onpopappended.dta
	cap erase onpopappendedv2.dta
		
		// Import and save Ontario PHU Population Projections from IntelliHealth
		global popprojectionsfilename 
		if !missing("${popprojectionsfilename}") {
			import excel "${popprojectionsfilename}", sheet("${popprojectionssheetname}") firstrow case(lower) clear
			tempfile onpop
			cap save `onpop'
		}
		
		// Append population projection data for 2017 to population estimates data for prior years
		import excel "${popestimatesfilename}", sheet("${popestimatessheetname}") firstrow case(lower) clear
		cap append using `onpop'
		
		// Convert PHU names to 4 characters, keep only observations of interest, and create custom age groups
		fixgeoname
		keepobs
		age_cat

		//Aggregate (sum) population counts in dataset by  year, geography, age, and sex
		collapse (sum) ${popvar}, by(${yearvar} geo2 age_catcustom ${sexvar})
		sort ${yearvar} geo2 ${agevar} ${sexvar}
		// li ${yearvar} geo2 age_catcustom ${sexvar} ${popvar}  in 1/20, sepby(geo2 ${yearvar})	

	save onpopappendedv2 
	****************************************************************************************************************************************
	****************************************************************************************************************************************
	/* Import and save case data*/
	cap erase casedata.dta
	// use "${casefilename}", clear
	import excel "${casefilename}", sheet("${casesheetname}") cellrange("${casesheetnamecellrange}") firstrow case(lower) clear
		format ${uniqueid} ${hn} %12.0g
		
		// Keep only variables of interest
		keep ${yearvar} ${agevar} ${sexvar} ${geovar} ${uniqueid} ${hn} ${othercasevartokeep}

		// Rename geographic regions, Keep only age groups and/or sex of interest,and categorize the data by ages groups of interest
		fixgeoname 
		keepobs
		age_cat 
		
		// Display distinct counts by unique id for Ontario
		di "Region, year, case count"
		foreach year of numlist $yearstart/$yearend {
			qui: distinct ${uniqueid} if ${yearvar}==`year'
			di "Ontario,", `year' "," r(ndistinct)
		}

	     di "Region, year, case count"
	    // Display distinct counts by unique id for localphu
	     foreach year of numlist $yearstart/$yearend {
			qui: distinct ${uniqueid} if ${yearvar}==`year' & geo2=="${localphu}"
			di "${localphu},", `year' "," r(ndistinct)
		}

		// Remove duplicates
		duplicates report ${uniqueid} ${hn} ${geovar} ${agevar} ${sexvar}
		// duplicates list ${uniqueid} ${hn} ${geovar} ${agevar} ${sexvar} in 1/200, sepby(${uniqueid} ${hn})
		bysort ${uniqueid}: drop if _n>1 
		
		// Create a variable, that will be aggregated to form the numerator
		gen cases=1
		
		// Aggregate (sum) population counts in dataset by  year, geography, age, and sex
		collapse (sum) cases, by(${yearvar} geo2 age_catcustom ${sexvar})
		sort ${yearvar} geo2 ${agevar} age_catcustom ${sexvar} 
		//li ${yearvar} geo2 age_catcustom ${sexvar} cases, sepby(geo2)
		
		// Add population (denominator) data
		merge 1:1 geo2 ${yearvar} age_catcustom ${sexvar} using onpopappendedv2
		//li ${yearvar} geo2 age_catcustom ${sexvar} case, sepby(geo2)
		
		// Assign numerator as 0 for groups (age and/or sex) without any cases
		replace cases = 0 if missing(cases)

	save casedata
	****************************************************************************************************************************************
	****************************************************************************************************************************************
	/* Direct standardization */
	
	// display rates per 100,000, show 95% CIs, and (optional) separate by sex
	// output case counts, study population count, crude rate, adjusted rate, lower CI, upper CI
	local mydistrateopts  "mult(100000) format($rateformat) level(${cilevel}) sepby(${sex}) list(cases N crude rateadj lb_gam ub_gam)"

	// Perform direct standardization for local PHU
	use casedata, clear
		cap erase localrates.dta

		// Age standardize if sexvar is missing, otherwise age and sex standardize
		noi: di "Rates for: " "$localphu"
		if !missing("${sexvar}") {
			// standardize rates (nuerator "cases"; denominator: "popvar"; standard population file "geo") for region "localphu"
			// standardize by age and/or sex for every year
			noi: distrate cases ${popvar} using pop if geo2=="$localphu", standstrata(age_catcustom ${sexvar}) by(${yearvar}) `mydistrateopts' saving(localrates)
		}
		else {
			noi: distrate cases ${popvar} using pop if geo2=="$localphu", standstrata(age_catcustom) by(${yearvar}) `mydistrateopts' saving(localrates)	
		}	

	// Perform direct standardization for Ontario
	use casedata, clear
		cap erase ontariorates.dta

		// Aggregate all of the numerator and denominator data for Ontario
		replace geo2="ONTARIO"
		collapse (sum) cases ${popvar}, by(geo2 ${yearvar} age_catcustom ${sexvar})
		sort geo2 ${yearvar} ${agevar} age_catcustom ${sexvar} 
		// li geo2 ${yearvar} age_catcustom ${sexvar} case ${popvar}

		// Age standardize if sexvar is missing, otherwise age and sex standardize
		noi: di "Rates for Ontario:"
		if !missing("${sexvar}") {
			noi: distrate cases ${popvar} using pop.dta, standstrata(age_catcustom ${sexvar}) by(${yearvar}) `mydistrateopts' saving(ontariorates)
		}
		else {
			noi: distrate cases ${popvar} using pop.dta, standstrata(age_catcustom) by(${yearvar}) `mydistrateopts' saving(ontariorates)	
		}
	****************************************************************************************************************************************
	****************************************************************************************************************************************
	/* Export results to excel */
	// Export PHU data to Excel. Columns will start from localstartcell
	use localrates, clear
	 	errorvalues
		export excel using "${finaloutputexcel}", sheet("rates") cell("${localstartcell}") firstrow(varlabel) sheetreplace

	// Export Ontario data to Excel. Columns will start from localstartcell
	use ontariorates, clear
	  	errorvalues
		export excel using "${finaloutputexcel}", sheet("rates") cell("${ontariostartcell}") firstrow(varlabel) sheetmodify
	******************************************************************************************************************************************
	******************************************************************************************************************************************
	/* Erase datasets created during session */
	local list pop onpopappended onpopappendedv2 casedata localrates ontariorates 
	foreach dataset of local list {
		cap erase "`dataset'.dta"
	}

	cap log close
	cap erase quiet_noise.log

	clear all
	clear results
	clear matrix
}
******************************************************************************************************************************************
******************************************************************************************************************************************
