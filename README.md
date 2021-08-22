# Redact-Excel

This short program has the ability to redact or change fields for a potentially unlimited number of excel (xlsx) spreadsheets in one execution. First you indicate which directory has the excel files you wish to alter, then you either can indicate two key-pair columns for replacing the values, build your own key-pair list, or indicate a single value you wish to use as your redactor based on a column. The program then creates the new, seperate excel files in a location of your choosing.

From what I can tell, there is not a readily accessible program online that works similar to this. As a business app, if you have a lot of files with linked/recurring values that need to be changed, this is an efficient alternate to going through all of them manually.

This program requires pandas and openpyxl. 
