# tvx_imei_assistant
TVX internal IMEI checkdigit and excel formatting tool

Start by downloading the XLS from Partner, follow below steps to do so.

  1. Go to RTFM and search for IMEI and go to the IMEI export Query.
  2. Copy the Query and use it in the Telavox > Web > Data export page (remember to swap the customer id to the customer you are exporting for).
  3. Download the result in Excel format.

When you have downloaded the file, run this program and click on Choose file. Choose the file you just downloaded from Partner.
When you select the file, the program will do the rest, it will calculate the checkdigit and append a new column to the Excel with the new data and hide the original column.
The program will save the new updated file as 'original filename' + '_updated.xlsx'.
