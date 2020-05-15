# Microscopy Instrument Monthly Usage Tracking in R

# Input

This code will extract data from .csv files containing a log of user usage for a microscopy instrument. Ths script is designed to batch process all .csv files in the working directory.

# Output

A compiled usage and payment report will be generated in an excel file.

The usage report contains the number of users grouped by supervisor, task performed per user, summed cost per task, and number of hours spent for each task. 
The user report contains the number of users grouped by supervisor, the individual user's name, task performed per user, summed cost per task, and number of hours spent for each task.
The payment report contains the total summed cost for each user.

The repository includes:
* Source code of the Microscopy Instrument Usage Tracking written in R.
