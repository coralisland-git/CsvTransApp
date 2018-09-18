# Spreadsheet Formatter
This solution lets you create a format specification and apply it to csv/xls/xlsx files.

## Install, Run and Test

### Install Dependencies
`pip install -r requirements.txt`

### How to run
`$ python Formatter.py -f <format> -i <input_file> -s <sheet_name>`  
`$ python Formatter.py -f riceland -i final_data.xlsx -s shop_data`

### Test
`$ cd test`  
`$ python test.py`


## Creating Formats and custom operations
Formats have to be stored as JSON files under the directory **formats**  
Eg: the format for **Seaworld** could be saved as **formats/seaworld.json**

The _custom_ operations have to be saved under the directory **operations**  
Eg: the custom operations for **seaworld** could be saved as **operations/seaworld.py**

There are 4 outer-level properties accepted in the format specification at the moment:
* _has_header_row_ - true/false depending on whether there is a header row
* _header_rows_
* _columns_
* _rows_

**header_rows, columns and rows** can have JSON values.  
Kindly use the **riceland.json** as a reference on how to specify operations for the time being.

## Order of operations
1. row deletions
2. header renaming
3. uniqueness check
4. cut-paste
5. custom operations
6. plain replacements
7. replacements based on columns
8. create new columns
9. drop columns


## Future
#### Limitation
Suppose for this spreadsheet:  

|  A  |  B   |  
| --- | ---- |  
| AP  |      |  
| OR  |      |  


if you you want to replace the value in column A based on a map 
`{'AP': 'Apple', 'OR': 'Orange'}`
and then cut the value in column A and paste it in column B

you use a format:
```
{
   "has_header_row": true,
   "columns": {
      "A": {
         "action": "replace",
         "with": {
            "": "",
            "AP": "Apple",
            "OR": "Orange"
         }
      },
      "B": {
         "action": "cutpaste",
         "from": "A"
      }
   }
}
```

and therefore expect the output to be:  

|  A  |  B     |  
| --- | ------ |  
|     | Apple  |  
|     | Orange |  

But it will not work as expected currently.

It is because the current format specification doesn't accept an order in which the operations are to be done. It _collects_ the instructions in the JSON, _groups_ them depending on the type and _executes_ those in the order specified in the previous section - **Order of operations**

In the example, cut-paste operation will be done **before** replacement and the result would be:  

|  A  |  B   |  
| --- | ---- |  
|     | AP   |  
|     | OR   |  

### Improvement
Use an instruction list instead of JSON.

Sample of such a list:  
`  replace col A map {'AP': 'Apple', 'OR': 'Orange'}`  
`  cut-paste col B from A`  
`   drop col A`

