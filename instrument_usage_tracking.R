"
Microscopy Instrument Usage Tracking

Written by Clara Tam
Copyright (c) 2020
Licensed under the MIT License (see LICENSE for details)

"
#--------------------------------------------------------------------------------
# Setup working environment
#--------------------------------------------------------------------------------
# Set working directory
working_dir <- getwd()
cat("Your working directory is set to:", working_dir, "\n")
setwd(working_dir)

#--------------------------------------------------------------------------------
# Define the input arguments
#--------------------------------------------------------------------------------
# specify date of reports
date <-"2019-11"

# csv files to import
confocal_file <- paste(date, "Confocal.csv")
epicalcium_file <- paste(date, "Epi&Calcium.csv")
zeissepi_file <- paste(date, "Zeiss_Epi.csv")

# output filename
confocal_savename <- paste0(date, "_Confocal", "_usage_report", ".xlsx")
epicalcium_savename <- paste0(date, "_EpiCalcium", "_usage_report", ".xlsx")
zeissepi_savename <- paste0(date, "_ZeissEpi", "_usage_report", ".xlsx")

# Microscopy Fees (Update as necessary)
confocal_price_list <- c(
  "CHRIM Training" = 50, 
  "CHRIM User" = 10, 
  "Data Analysis" = 0, 
  "CHRIM Full Service" = 50)

epicalcium_price_list <- c(
  "CHRIM Training" = 50, 
  "CHRIM User" = 10, 
  "Data Analysis" = 0, 
  "CHRIM Full Service" = 50)

zeissepi_price_list <- c(
  "CHRIM Training" = 50, 
  "CHRIM User" = 10, 
  "Data Analysis" = 0, 
  "CHRIM Full Service" = 50)
#--------------------------------------------------------------------------------

# import required library packages
library(plyr)
library(stringr)
library(openxlsx)

#--------------------------------------------------------------------------------
# Functions
#--------------------------------------------------------------------------------
dataframe_preprocessing <- function(data_df){
  "
  Arguments:
  data_df -- dataframe, of shape (num_obs, num_variables)
  
  Returns:
  data_df -- dataframe, of shape (num_obs, num_variables)
  "
  # remove nonalphanumeric characters from header names
  names(data_df) <- gsub("[^[:alnum:]///' ]", "", names(data_df))
  names(data_df)[1] <- substring(names(data_df)[1],2)
  # standardize supervisor names by removing any punctuations
  data_df$Supervisor <- gsub("[[:punct:]]", "", data_df$Supervisor)
  
  return(data_df)
}


instrument_usage <- function(data_df, price_list){
  " 
  Arguments:
  data_df -- input dataframe, of shape (num_obs, num_variables)
  price_list -- named atomic vector type numeric, of shape (1, 3)
  
  Return:
  usage_report -- dataframe, of shape (num_obs, 4)
  payment_report -- dataframe, of shape (num_obs, 2)
  "
  if(names(data_df)[1] == "Confocal"){
    # Confocal instrument usage report
    usage_report <- aggregate(Price ~ Supervisor + Confocal, data=data_df, sum) #edit instrument name as needed
    
    # calculate the number of hours spent per task
    # preallocate usage_hours vector
    usage_hours <- c()
    
    for (item in 1:length(price_list)){
      if (names(price_list[item]) %in% usage_report$Confocal){ #edit instrument name as needed
        temp <-subset(usage_report, Confocal == names(price_list)[item]) #edit instrument name as needed
        num_hours <- temp$Price / price_list[item]
        usage_hours <- append(usage_hours, num_hours, after=length(usage_hours))
        usage_hours[is.nan(usage_hours)] <- 0
      }# end inner if
    }# end for
  }# end of outer if statement
  else if (names(data_df)[1] == "EpiCalcium"){
    # EpiCalcium instrument usage report
    usage_report <- aggregate(Price ~ Supervisor + EpiCalcium, data=data_df, sum) #edit instrument name as needed
    
    # calculate the number of hours spent per task
    # preallocate usage_hours vector
    usage_hours <- c()
    
    for (item in 1:length(price_list)){
      if (names(price_list[item]) %in% usage_report$EpiCalcium){ #edit instrument name as needed
        temp <-subset(usage_report, EpiCalcium == names(price_list)[item]) #edit instrument name as needed
        num_hours <- temp$Price / price_list[item]
        usage_hours <- append(usage_hours, num_hours, after=length(usage_hours))
        usage_hours[is.nan(usage_hours)] <- 0
      }# end if
    }# end for
  }# end of if statement
  else if(names(data_df)[1] == "ZeissEpi"){
    # ZeissEpi instrument usage report
    usage_report <- aggregate(Price ~ Supervisor + ZeissEpi, data=data_df, sum) #edit instrument name as needed
    
    # calculate the number of hours spent per task
    # preallocate usage_hours vector
    usage_hours <- c()
    
    for (item in 1:length(price_list)){
      if (names(price_list[item]) %in% usage_report$ZeissEpi){ #edit instrument name as needed
        temp <-subset(usage_report, ZeissEpi == names(price_list)[item]) #edit instrument name as needed
        num_hours <- temp$Price / price_list[item]
        usage_hours <- append(usage_hours, num_hours, after=length(usage_hours))
        usage_hours[is.nan(usage_hours)] <- 0
      }# end if
    }# end for
  }# end of else if statement
  
  # add hours to usage_report dataframe
  usage_report$Hours = usage_hours
  
  # calculate total price for each supervisor
  payment_report <- aggregate(Price ~ Supervisor, data=usage_report, sum)
  # create list to return multiple reports
  return_list <- list(usage_report, payment_report)

  return(return_list)
}


export_reports <- function(reports, savename){
  "
  Arguments:
  reports -- list, of shape (2, 1)
  savename -- string, of shape (1, 1)
  "
  # convert individual reports back to dataframes
  usage_report <- data.frame(reports[1])
  usage_report <- usage_report[order(usage_report[,1]),] # sort df by supervisor
  payment_report <- data.frame(reports[2])
  
  # export reports
  wb <- createWorkbook()
  addWorksheet(wb, "Usage Report")
  addWorksheet(wb, "Payment Totals")
  
  writeData(wb, "Usage Report", usage_report)
  writeData(wb, "Payment Totals", payment_report)
  
  saveWorkbook(wb, file=savename, overwrite=TRUE)
  cat("\nFile:", savename, " has been created")
}

#--------------------------------------------------------------------------------
# Main Script
#--------------------------------------------------------------------------------

# check if csv files to import exists
# generate report if files exist

# Conofocal Instrument
if (file_test("-f", confocal_file)){
  # read csv file
  confocal_data_df <- read.csv(file=confocal_file, header=TRUE, stringsAsFactors=FALSE)
  # preprocessing dataframe
  confocal_data_df <- dataframe_preprocessing(confocal_data_df)
  # retrieve reports from instrument_usage function
  confocal_reports <- instrument_usage(confocal_data_df, confocal_price_list)
  # export report
  export_reports(confocal_reports, confocal_savename)
}# end if

# Epi & Calcium Instrument
if (file_test("-f", epicalcium_file)){
  # read csv file
  epicalcium_data_df <- read.csv(file=epicalcium_file, header=TRUE, stringsAsFactors=FALSE)
  # preprocessing dataframe
  epicalcium_data_df <- dataframe_preprocessing(epicalcium_data_df)
  # retrieve reports from instrument_usage function
  epicalcium_reports <- instrument_usage(epicalcium_data_df, epicalcium_price_list)
  # export report
  export_reports(epicalcium_reports, epicalcium_savename)
}# end if

# Zeiss Epi Instrument
if (file_test("-f", zeissepi_file)){
  # read csv file
  zeissepi_data_df <- read.csv(file=zeissepi_file, header=TRUE, stringsAsFactors=FALSE)
  # preprocessing dataframe
  zeissepi_data_df <- dataframe_preprocessing(zeissepi_data_df)
  # retrieve reports from instrument_usage function
  zeissepi_reports <- instrument_usage(zeissepi_data_df, zeissepi_price_list)
  # export report
  export_reports(zeissepi_reports, zeissepi_savename)
}# end if
