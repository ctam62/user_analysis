"
Microscopy instrument usage tracking

"
#--------------------------------------------------------------------------------
# Setup working environment
#--------------------------------------------------------------------------------
# Set working directory
setwd("D:/instrument_usage_tracking")
working_dir <- getwd()

# import required library packages
library(plyr)
library(stringr)
library(openxlsx)

#--------------------------------------------------------------------------------
# Functions
#--------------------------------------------------------------------------------
instrument_usage <- function(data_df, price_list){
  " 
  Arguments:
  data_df -- input dataframe, of shape (num_obs, num_variables)
  price_list -- named atomic vector type numeric, of shape (1,3)
  
  Return:
  usage_report -- dataframe, of shape (num_obs, 4)
  payment_report -- dataframe, of shape (num_obs, 2)
  "
  num_supervisors <- count(data_df, c("Supervisor", "ïZeissEpi"))
  usage_report <- aggregate(Price ~ Supervisor + ïZeissEpi, data=data_df, sum)
  
  # calculate the number of hours spent per task
  # preallocate usage_hours vector
  usage_hours <- c()
  
  for (item in 1:length(price_list)){
    if (names(price_list[item]) %in% usage_report$ïZeissEpi){
      temp <-subset(usage_report, ïZeissEpi == names(price_list)[item])
      num_hours <- temp$Price / price_list[item]
      usage_hours <- append(usage_hours, num_hours, after=length(usage_hours))
      usage_hours[is.nan(usage_hours)] <- 0
    }
  }
  
  # add hours to usage_report dataframe
  usage_report$Hours = usage_hours
  
  # calculate total price for each supervisor
  payment_report <- aggregate(Price ~ Supervisor, data=usage_report, sum)
  # create list to return multiple reports
  return_list <- list(usage_report, payment_report)

  return(return_list)
}

#--------------------------------------------------------------------------------
# Main Script
#--------------------------------------------------------------------------------

# Define the input arguments
#--------------------------------------------------------------------------------
# Microscopy Fees (Update as necessary)
price_list <- c(
  "CHRIM Training" = 50, 
  "CHRIM User" = 10, 
  "Data Analysis" = 0, 
  "CHRIM Full Service" = 50)
#--------------------------------------------------------------------------------

# import excel file
data_df <- read.csv(file='2019-8 Zeiss_Epi.csv', header=TRUE, 
                    stringsAsFactors=FALSE)
# remove nonalphanumeric characters from header names
names(data_df) <- gsub("[^[:alnum:]///' ]", "", names(data_df))
# standardize supervisor names by removing any punctuations
data_df$Supervisor <- gsub("[[:punct:]]", "", data_df$Supervisor)

# retrieve reports from instrument_usage function
accounting_reports <- instrument_usage(data_df, price_list)
# convert individual reports back to dataframes
usage_report <- data.frame(accounting_reports[1])
usage_report[order(usage_report[,1]),] # sort df by supervisor
payment_report <- data.frame(accounting_reports[2])

# export reports
wb <- createWorkbook()
addWorksheet(wb, "Usage Report")
addWorksheet(wb, "Payment Totals")

writeData(wb, "Usage Report", usage_report)
writeData(wb, "Payment Totals", payment_report)

saveWorkbook(wb, file="instrument_usage_reports.xlsx", overwrite=TRUE)

