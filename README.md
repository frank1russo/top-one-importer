# top-one-importer
Reads one or more csv files from Top One Futures and converts it to an excel file. 

Each csv is placed in it's own worksheet with a column that calculates the Total PnL for all 
trades.

The final excel file contains a summary worksheet that includes a sum of all accounts.

To download the csv files log into your accounts at toponefutures.com, click on the account tab on
the left and choose an account. Scroll to the bottom of the account page, click the Trade History 
tab, and click the Download CSV button on the right above the table. 

Copy the csv files into the project's imports folder and then run the command ./run_converter.sh
from the project's root folder.