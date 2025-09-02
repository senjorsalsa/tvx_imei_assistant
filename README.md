# tvx_imei_assistant
TVX internal IMEI checkdigit and excel formatting tool

Start by downloading the XLS from Partner, follow below steps to do so.

  1. Go to RTFM and search for IMEI and go to the IMEI export Query.
  2. Copy the Query and use it in the Telavox > Web > Data export page (remember to swap the customer id to the customer you are exporting for).
  3. Download the result in Excel format.

When you have downloaded the file, run this program and click on Choose file. Choose the file you just downloaded from Partner.
When you select the file, the program will do the rest, it will calculate the checkdigit and append a new column to the Excel with the new data and hide the original column.
The program will save the new updated file as 'original filename' + '_{random 8-bit hex value}.xlsx'.


Step by step explanation of the code

1. File Selection and Data Reading

  When you click the "Choose File" button, a file dialog window pops up, asking you to select an .xls file.
  After you select a file, the script uses the xlrd library to read the data from the selected .xls file.
  It then creates a new, empty openpyxl workbook, which is the modern .xlsx format.
  The script copies all the data from the original .xls file into this new .xlsx workbook, preserving the sheet name and content.

2. IMEI Validation and Correction

  The script inserts a new, empty column at the 6th position (column F) in the new workbook. It labels the first cell of this new column as "imei".
  It then reads the IMEI numbers from the 5th column (column E) of the original data. The IMEI number is the 15-digit unique number for every mobile phone.
  For each IMEI number it reads, the script performs a Luhn algorithm check. The Luhn algorithm is a simple checksum formula used to validate various identification numbers. This check verifies if the last digit (the check     digit) of the IMEI is correct based on the preceding 14 digits.
  If the check digit is correct, the script keeps the original IMEI number.
  If the check digit is incorrect, the script calculates the correct check digit and replaces the original one, essentially correcting the IMEI number.
  The script then hides the original 5th column (column E) so that only the new, corrected IMEI numbers in column F are visible.

4. Saving the New File

  Finally, the script creates a new filename by adding a unique, random string to the original file's name. For example, if the original file was data.xls, the new file might be named data_1a2b3c4d.xlsx. This ensures the       original file is never overwritten.
  The new workbook, with the corrected IMEI numbers and the hidden column, is then saved to this new file.
