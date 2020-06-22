# Monthly Instrument Usage Tracking in R

This code extracts data from .csv files obtained from a SAS database containing a log of user usage for an instrument. The script is designed to automatically batch process all .csv files in the data directory.

## Required Packages

This program was written in R and requires the following library packages:

```
rstudioapi
lubridate
data.table
stringr
modules
```

## Expected Output

A compiled usage, user, and payment report will be generated in a single excel workbook file for each month available.

The usage report contains the number of users grouped by supervisor, task performed per user, summed cost per task, and number of hours spent for each task. 
The user report contains the number of users grouped by supervisor, the individual user's name, task performed per user, summed cost per task, and number of hours spent for each task.
The payment report contains the total summed cost for each user.

The repository includes:
* Source code of the Monthly Instrument Usage Tracking written in R.
* Sample data to test 

## Usage

**CSV File Format**:
The input files are named with the format "year-month instrument name". 

Run monthly_instrument_usage_tracking.R file to see the demo script output.
