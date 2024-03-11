# Worksheet Macros
A collection of Excel/VBA macros that was created originally as per request from other people to address their specific requirements which I then made into a more generic macro.

## WorksheetMacros.xslm
- [Index worksheet](#index-worksheet)
- [Extract worksheet](#extract-worksheet)
- [Combine worksheet](#combine-worksheet)
- [Compile worksheet](#compile-worksheet)

# WorksheetMacros.xlsm
Below is the list of all the worksheets in this workbook.

## Index worksheet
Lists all the other Excel sheets in the workbook, their descriptions, and a link to quickly navigate to them.

<b>NOTE:</b> This worksheet is automatically repopulated when saving the workbook.

## Extract worksheet
Copy all Excel files in a folder, retain specific sheets on each one, and place them on another folder.

<b>NOTE:</b> Each file processed will result in its own file bearing the same name.

### Parameters
#### General parameters
|Parameter Name|Description|Remarks|
|:-|:-|:-|
|Find all Excel files in:|The directory containing all the Excel files to be processed.|Only Excel files at the top level of  this folder will be processed. Other files (non-Excel files, hidden/system files, Excel files within sub folders) will be ignored.<br /><br />Will throw an error if the directory does not exist.|
|Extract the following sheets and place them into:|The directory where the output of the worksheet will be written to.|Each file processed will result in its own file bearing the same name on this folder.<br /><br />Will throw an error if the directory does not exist.|

#### Table parameters
These will apply to each file being processed.
|Column Name|Description|Remarks|
|:-|:-|:-|
|Worksheet name/index to extract|The name or index of the worksheet in the processed file to copy details from.|Will throw an error if the worksheet does not exist on the file. <br /><br /><b>NOTE:</b> The index of the first worksheet is 1.|
|Range to extract|The range in the worksheet specified on the previous parameter to copy details from.|The value has to be in a valid Excel range format (i.e.: `A:C`, `A2:B10,D2:H10`, etc).|
|New worksheet name (optional)|The name of the new worksheet in the output where the details will be copied to.|If left blank, the name will be the same as the name of the worksheet where the details were copied from.|

## Combine worksheet
Get specific Excel files in a folder, extract specific sheets on each one, and combine them into one Excel file.

### Parameters
#### General parameters
|Parameter Name|Description|Remarks|
|:-|:-|:-|
|Find the required files in:|The directory containing all the Excel files to be processed.|Will throw an error if the directory does not exist.|
|Combine those files and place them into:|The directory where the output of the worksheet will be written to.|Will throw an error if the directory does not exist.|
|Look for files that starts with:|The prefix of the files to be processed.|Will be expected on each file that will be processed|
|Combine files into:|The name of the output file.||

#### Table parameters
These will apply to each file being processed.
|Column Name|Description|Remarks|
|:-|:-|:-|
|Excel files to combine|The name of the Excel file to be processed.|The prefix parameter and this value will form the full name of the Excel file to be processed.<br /><br />Will throw an error if the file does not exist or is not an Excel file.|
|Worksheet name/index to copy|The name or index of the worksheet in the processed file to copy.|Will throw an error if the worksheet does not exist on the file. <br /><br /><b>NOTE:</b> The index of the first worksheet is 1.|
|Range to copy (optional)|The range in the worksheet specified on the previous parameter to copy details from.|The value has to be in a valid Excel range format (i.e.: `A:C`, `A2:B10,D2:H10`, etc).<br /><br />If left blank, the entire sheet will be copied.|
|New worksheet name (optional)|The name of the new worksheet in the output that is a copy of the worksheet referenced on the previous parameter.|If left blank, the name will be the same as the name of the worksheet where the details were copied from.|

## Compile worksheet
Gets all Excel files in a folder, extract specific sheets on each one, and compile them into one Excel file.

### Parameters
#### General parameters
|Parameter Name|Description|Remarks|
|:-|:-|:-|
|Find all Excel files in:|The directory containing all the Excel files to be processed.|Only Excel files at the top level of  this folder will be processed. Other files (non-Excel files, hidden/system files, Excel files within sub folders) will be ignored.<br /><br />Will throw an error if the directory does not exist.|
|Combine those files and place them into:|The directory where the output of the worksheet will be written to.|Will throw an error if the directory does not exist.|
|Extract the following sheets from each file and compile them into:|The name of the output file.||
|Remove this text from the start of the file to get the prefix:|The text to remove from the start of the name of each file being processed.|Will be used later on to determine the prefix that represents the file being processed.|
|Remove this text from the end of the file to get the prefix:|The text to remove at the end of the name of each file being processed.|Will be used later on to determine the prefix that represents the file being processed.<br /><br /><b>NOTE:</b> The name of the file includes the extension (.xlsx).|

#### Table parameters
These will apply to each file being processed.
|Column Name|Description|Remarks|
|:-|:-|:-|
|Worksheet name/index to compile|The name or index of the worksheet in the processed file to copy.|Will throw an error if the worksheet does not exist on the file. <br /><br /><b>NOTE:</b> The index of the first worksheet is 1.|
|Range to copy (optional)|The range in the worksheet specified on the previous parameter to copy details from.|The value has to be in a valid Excel range format (i.e.: `A:C`, `A2:B10,D2:H10`, etc).<br /><br />If left blank, the entire sheet will be copied.|
|New worksheet name (optional)|The name of the new worksheet in the output that is a copy of the worksheet referenced on the previous parameter.|If left blank, the value will be the same as the name of the worksheet where the details were copied from..<br /><br />The prefix and this value will form the actual worksheet name on the output.<br /><br />The prefix will be what remains on the name of the processed file after removing certain values at the start and end as specified in the corresponding parameters. No errors will be thrown if these values do not exist in the name.<br /><br /><b>NOTE:</b> The prefix is used to ensure that there are only unique worksheet names in the output, as well as provide a reference as to which file the worksheet came from.|