# Dataverse Report

The following project generates an Excel Workbook with statistics pulled from a Dataverse database.

To use this code, edit the `sql_connect.py` file with connection details for your postgres database and run the project from the `init.py` file.

The resulting excel file can be used in the metrics dashboard available at https://github.com/scholarsportal/Dataverse-Web-Report.

Tested with Dataverse Version 4.19

## Dev Environment Set Up

To install python modules:

`pip3 install -r requirements.txt`

To update requirements.txt:

`pip3 freeze > requirements.txt`