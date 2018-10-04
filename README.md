# iTMS-Job-Status-Excel-Formatter
A Python Script to format the 'Current Job Status' excel spreadsheet that is exported from iTMS.

## Getting Started

You will need an Excel Spreadsheet from iTMS for this script to work. In iTMS, open Jobs (5), Current Job Status (5), click on the refresh button, then click on the 'Export the list to Microsoft Excel...' button. This will launch Excel, save this file where the script is located and call the file 'book1.xlsx'

### Prerequisites

This script uses the openpyxl library, this library will need to be installed for this script to work.

https://openpyxl.readthedocs.io/en/stable/

```
$ pip install openpyxl
```

### Installing

Copy the script anywhere you like, preferably its own folder and make the script executable

### Runnings the Script

Once you have exported a spreadsheet into the folder where the script is, you can now run the script


```
Change into the directory where the script is located,
$ python job-status-formatter.py
```
