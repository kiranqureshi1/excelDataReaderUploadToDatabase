# excelDataReaderUploadToDatabase-app

Technology Used:
C#
ExcelDataReader
SMO (Server Management Objects)

Basic Functionality:
App goes to folder location and open the folder and the loops through the Excel files
Opening each Excel file one by one, reading file name and then looping through sheets, reading each sheet name and columns names, rows, their dataTypes.
In the next step the app checks for basic validation. 
First is the sheet empty?
It then loops through all the column names to see:
Are there any null column names?
does each column name exceeds 128 characters limit?
Are there any duplicated column names?

Then it loops through rows
Checking:
Is the row data exceeds – characters limit

