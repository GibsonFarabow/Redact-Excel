# Redact-Excel

This program has the ability to redact or change fields for a potentially unlimited number of excel (xlsx) spreadsheets in one execution. First you indicate which directory has the excel files you wish to alter, then you indicate value-pairs to change or indicate a single value you wish to use as the update based on a column of values. The program then creates the new, seperate excel files in a location of your choosing. Version 1 does so for books with only 1 sheet; version 2 can handle 2 and is editted for better clarity.

From what I can tell, there is not a readily accessible program online that works similar to this. As a business app, if you have a lot of files with linked/recurring values that need to be changed, this is an efficient alternate to going through all of them manually.

This program requires pandas and openpyxl. 
