# Direct standardization 
## By age and/or sex using the Faye and Feuer (1997) method <br/>
## Aim:
To calculate directly standardized (age and/or sex standardized) rates for PHUs and Ontario using the 2011 Canadian Population as the standard population.

## Requirements:
1. Standard population: An Excel dataset containing the 2011 Canadian Population by age (numeric) and/or sex (string)
2. PHU population estimates (2010-2016): An Excel dataset containing population estimates by PHU (string), age (numeric) and/or sex (string)
3. (Optional) PHU population projections (2017): An Excel dataset containing PHU population projections by PHU (string), age (numeric) and/or sex (string)
4. Case data: An Excel dataset containing individual or aggregate level case data by PHU (string), age (numeric) and/or sex (string), with ident

## Methods:
### Standard population:
1a. Import and save the standard population data <br/>
1b. Categorize age values into custom age groups


### PHU Population (denominator for crude rates):
2a. Import and save the PHU population estimates <br/>
2b. (If using projections) Append the PHU population projections to the PHU population estimates <br/>
2c. Adjust the PHU names (to allow for merging later on) <br/>
2d. Keep only the observations of interest <br/>
2e. Categorize age values into custom age groups <br/>

### Case data (numerator data for crude dates)
3a. Import and save case data <br/>
3b. Adjust the PHU names (to allow for merging later on) <br/>
3c. Keep only the observations of interest <br/>
3d. Remove any duplicates <br/>
3e. Numerator: Collapse the individual level data by the strata of interest (i.e. phu, year, age, and/or sex). In other words, group the dataset by the strata and sum the number of cases in each strata. <br/>
3f. Denominator: Merge the PHU population data (denominator) with the case data (numerator) <br/>

### Direct standardize
4a. Apply direct standardization command for PHU <br/>
4b. Apply direct standardization command for province <br/>

### Output results
5a. Export results to excel
