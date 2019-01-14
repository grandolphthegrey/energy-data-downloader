#Overview of repo and individual scripts:

##datadownload.py

###Basic Functionality
the Python script "datadownload.py" scrapes publicly available data from the following sources:

- CME Group NYMEX Futures
- CME Group Henry Hub Natural Gas Future Settlements
- Nodal Exchange settlement prices
- CAISO NP15 and SP15 prices

this script also downloads subscription based data from Platts:

- Gas Daily Prices
- Megawatt Daily Prices

Valid credentials are needed to download the data from Platts

Each data source runs in its own function and saves the output to Excel format

The script takes general inputs from a "settings" file that contain URLs, credentials, email addresses and directory locations. This is so that the most likely changed settings can be modified easily without the need to modify individual lines of code in the script-- in other words, non-programmers are able to modify the functionality of the script, without having to scroll through lines and lines of code

###Other Features

There is built in logic to detect if data is downloaded on a holiday (i.e. repeated data) and renames the files accordingly. 

All activity is saved to log files. If errors occur, alert emails are sent
