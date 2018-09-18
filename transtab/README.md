# TransTab    
  
'TransTab' is 'Transform Tabular-data'  
  
This solution transforms the given csv/xls/xlsx file into a desired format.  
  
The format has to be specified as a sequence of operations.  
The solution will execute the operations one-by-one to finally save the result as a .xlsx file  
  
Please note that if you rename a col from 'A' to 'B', 
then the subsequent references to this column should be made by 'B'  
  
Also if you want to delete the very first two rows, you have to do it in the following fashion:  
` delete row 1 `  
` delete row 1 `  
  
## Specifying a format - commands, their meaning and options.  
```  
delete row 1  
  
# You may add comments like this line. Everything after a hash is a comment on that line.  
# The program slice and dice the data using its header row(row that contains the column names)  
# So if you have some non-headers rows at the beginning of the input file,  
# they should be removed before any other operation can be done.  
  
# ie, if there are 3 non-header rows in the beginning,  
# you have to enter 'delete row 1' thrice before any other operations.  
  
  
dates 'Birth Date', 'Hire Date'  
  
# User has to specify which the date columns are using the above command  
# This is mainly because in some formats like .xls, .xlsx the date is simply a number.  
# So we need this declaration to do the number-to-date conversion before other operations.   
# All remaining commands can be specified in any order  
  
  
delete-duplicate-rows  
  
# also possible: delete-duplicate-rows unique <col_name>  
# so that duplicates are identified on the basis of just <col_name>   
  
  
delete 'File Number'  
  
# delete <col_name>  
  
  
rename 'SSN' as 'Employee SSN'  
  
# 'as' is optional, you may add it for readability.  
  
  
clear 'Hourly Rate'  
  
  
copy 'Payroll Frequency' to 'Pay Group'  
cutpaste 'Pay Group' to 'Department Code'  
  
# 'to' is optional.  
  
  
concatenate 'First Name', 'Last Name' and store in 'Name' using ' '  
  
# using 'str': str will be placed in between the strings being concatenated.  
# It is optional, empty string is used as the default join-string.  
# concatenate, copy, cutpaste destination columns will be created if they don't exist already  
  
  
replace 'Employee Status'  
{  
	'Active': 'A',  
	'Terminated': 'T',  
	'Leave of Absence': 'L',  
	'Inactive': 'I'  
} case-insensitive default 'null'  
  
# case-insensitive: optional flag to indicate that case-insensitive matching may be done  
# default  <val>: If the word found in cell has no corresponding entry in the replacement map,  
# the <val> is used. This is an optional setting, not enabled by default.  
# So you get error if there is no matching key in the map for a cell value.  
  
  
sum-col-and-delete-duplicate-rows sum 'Annual Rate' unique 'Employee SSN'  
  
# sum-col-and-delete-duplicate-rows sum <col_name1> unique <col_name2>   
# Sum up <col_name1> for each unique value of <col_name2>,  keep a row with that sum  
# and delete other(duplicate) rows  
  
  
do validate_phone_number on 'Phone number'  
do validate_ssn on 'Employee SSN'  
do validate_number on 'Payroll Frequency'  
  
# validate_phone_number, validate_ssn, validate_number are available globally  
# This kind of custom operations, where the column is specified, can modify only one(specified) column in any row  
# The col, row values are passed to the function.  
# The value returned by the function is stored in the specified column  
  
  
do set_hire_date  
  
# This kind of custom operations, where the column is not specified, can modify any column in any row  
# Entire rows are passed to the function.  
# The row object returned by the function is stored in place of the previous row    
  
  
# The custom operations may be defined at two places:  
# 1) global_operations.py file. Re-install the package from source whenever you update this file    
# 2) <format_name.py> This file is read from the same location as that of the format specification file  
```  
  
The solution will apply each operation in the specified order and save the final state of the file.  
  
## Install, Run and Test  
  
### Install  
`pip install -r requirements.txt`  
`pip install --upgrade dist/transtab-0.0.1.tar.gz`  
  
  
### Run  
From anywhere on the commandline:  
`$ transtab -f <format> -i <input_file>`  
`$ transtab -f riceland.txt -i riceland.xlsx`  
  
`Optionally you can pass a sheetname for the output file with the -s option`  
  
Or from a python program:  
`from transtab import TransTab`  
`TransTab(in_fname = <input filename>, format_fname = <format_filename>, out_sheet=<out_sheetname>).transform()`  

### Test  
`python test_transtab.py`  
  
Note: Some error indications are part of the testing (like error message when validate_number fails as expected.)  
If it says 'OK' at the last line in the output, tests were passed.  
