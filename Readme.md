# Overview of repo and individual scripts:

## datadownload.py

### Basic Functionality
the Python script "datadownload.py" scrapes publicly available data from the following sources:

  - CME Group NYMEX Futures
  - CME Group Henry Hub Natural Gas Future Settlements
  - Nodal Exchange settlement prices
  - CAISO NP15 and SP15 prices

this script also downloads subscription based data from Platts:

  - Gas Daily Prices
  - Megawatt Daily Prices

Note that valid credentials are needed to download the data from Platts

Each data source runs in its own function and saves the output to Excel format

The script takes general inputs from a "settings" file that contain URLs, credentials, email addresses and directory locations. This is so that the most likely changed settings can be modified easily without the need to modify individual lines of code in the script-- in other words, non-programmers are able to modify the functionality of the script, without having to scroll through lines and lines of code

### Other Features

There is built in logic to detect if data is downloaded on a holiday (i.e. repeated data) and renames the files accordingly. 

All activity is saved to log files. If errors occur, alert emails are sent

## databaseupdate.py

### Basic Funcionality
the Python script "databaseupdate.py" updates Microsoft Access databases with the data downloaded from the "datadownload.py" script

  - The data from each source is housed in its own database
  - This script updates the databases three business days prior to the current date. This provides a buffer, should the data not download automatically 
  - Depending on the the day of the week, not all databases will be updated
  
There is built in logic to determine if the data has been downloaded on a holiday. Typically, the same data will be posted to the source website. The script checks the validity of the data and determines if the data is a repeat. If data has been repeated, then the data will not be uploaded to the database. 

This activity is recorded in the daily log files.

## tariffs.py

### Basic Functionality
The Python script "tariffs.py" downloads the electricity and gas tariffs from each of the four Investor Owned Utilities in the state of California. The relevant URLs for where these tariffs are hosted are saved in a separate .txt file, to enable non-programmers to modify the script, should the URL changes.

Only a single URL is provided for each utility. The script searches the webpages for relevant keywords and finds the appropriate subdomains from which to download the tariffs. This prevents the user from having to enter several different URLs for the same utility. (e.g. rather than entering Residential, Commercial, and Agricultural URLs for Pacific, Gas & Electric, the script takes a single URL and then searches the webpage for relevant keywords to find those categories of tariffs)

## forecastfiles.py

### Basic Functionality
The Python script "forecastfiles.py" queries the Access database and extracts the relevant gas prices from the appropriate price databases. This script is designed to work with an external .txt file, so the user can specify  which natural gas hubs are queried out. Querries are made via SQL using the PYODBC library. Consequently, this script only functions in a Windows environment.

## cpuccal.py

### Basic Functionality
This Python script downloads the daily agenda and calendar posted on the California Public Utitlies (CPUC) site  
