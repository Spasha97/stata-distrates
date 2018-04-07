# stata-distrates
## Aim:
To calculate directly standardized (age and sex standardized) rates for PHUs and Ontario using the 2011 Canadian Population as the standard population.

## Requirements:
1. Standard population: An Excel dataset containing the 2011 Canadian Population by age (single year) and sex
2. PHU population estimates (2010-2016): An Excel dataset containing PHU population estimates by age (single year) and sex
3. PHU population projections (2017): An Excel dataset containing PHU population projections by age (single year) and sex
4. Case data: An Excel dataset containing individual or aggregate level case data 

## Methods:
### Standard population:
1a. Import and save the standard population data <br/>
1b. Categorize single year ages into custome age groups


### PHU Population (denominator for crude rates):
2a. Import and save the PHU population estimates <br/>
2b. Append the PHU population projections to the PHU population estimates <br/>
2c. Fix the variable name for the phus (to allow for merging later on) <br/>
2d. Keep only the observations of interest <br/>
2e. Categorize single year ages into custome age groups <br/>

### Case data (numerator data for crude dates)
3a. Import and save case data <br/>
3b. Fix the variable name for the phus (to allow for merging later on)  
3c. Keep only the observations of interest <br/>
3d. Remove any duplicates <br/>
3e. Numerator: Collapse the individual level data by the strata of interest (i.e. phu, year, age, sex). In other words, group the dataset by the strata and sum the number of cases in each strata. <br/>
3f. Denominator: Merge the phu population data (denominator) with the case data (numerator) <br/>

### Direct standardize
4a. Apply direct standardization command for PHU <br/>
4b. Apply direct standardization command for province <br/>

### Output results
5a. Export results to excel
