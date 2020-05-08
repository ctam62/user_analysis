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
date <-"2020-2"

# csv files to import
confocal_file <- paste(date, "Confocal.csv")
epicalcium_file <- paste(date, "Epi&Calcium.csv")
zeissepi_file <- paste(date, "Zeiss_Epi.csv")

# output filename
output_dir <- file.path(working_dir, paste0(date,"_reports"))
if (!dir.exists(output_dir)){dir.create(output_dir)}

confocal_savename <- paste0(output_dir,"/", date, "_Confocal", "_usage_report", ".xlsx")
epicalcium_savename <- paste0(output_dir, "/", date, "_EpiCalcium", "_usage_report", ".xlsx")
zeissepi_savename <- paste0(output_dir, "/", date, "_ZeissEpi", "_usage_report", ".xlsx")

# Microscopy Fees (Update as necessary)
confocal_price_list <- c(
  "CHRIM Training" = 75, 
  "CHRIM User" = 50, 
  "Data Analysis" = 0, 
  "CHRIM Full Service" = 50,
  "UofM User" = 75,
  "External" = 150)

epicalcium_price_list <- c(
  "CHRIM Training" = 75, 
  "CHRIM User" = 25, 
  "Data Analysis" = 0, 
  "CHRIM Full Service" = 50,
  "UofM User" = 30,
  "External" = 60)

zeissepi_price_list <- c(
  "CHRIM Training" = 75, 
  "CHRIM User" = 25, 
  "Data Analysis" = 0, 
  "CHRIM Full Service" = 50,
  "UofM User" = 30,
  "External" = 60)
#--------------------------------------------------------------------------------

# import required library packages
library(plyr)
library(stringr)
library(openxlsx)
library(lubridate)

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
  # names(data_df)[1] <- substring(names(data_df)[1],2)
  # standardize supervisor names by removing any punctuations
  data_df$Supervisor <- gsub("[[:punct:]]", "", data_df$Supervisor)
  
  return(data_df)
}


compute_timespent <-function(data_df){
  "
  Arguments:
  data_df -- dataframe, of shape (num_obs, num_variables)
  
  Return:
  num_hours -- numeric value
  
  "
  # reformat Starttime and Finishtime
  data_df$Starttime <- sub(".+? ", "", data_df$Starttime)
  data_df$Finishtime <- sub(".+? ", "", data_df$Finishtime)
  # extract time values
  starttime <- as.numeric(hm(data_df$Starttime), units="hours")
  endtime <- as.numeric(hm(data_df$Finishtime), units="hours")
  # compute the difference
  num_hours <- abs(endtime - starttime)
  
  return(num_hours)
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
    
    # extract supversior and instrument columns  
    temp_report <- subset(data_df, select=c("Supervisor", "Fullname", "Confocal", "Price"))
    # calculate the number of hours spent per task
    # preallocate usage_hours vector
    usage_hours <- c()
      
    for (item in 1:length(price_list)){
      if (names(price_list[item]) %in% data_df$Confocal){
        # extract data for Confocal
        temp <-subset(data_df, Confocal == names(price_list)[item])
        num_hours <- compute_timespent(temp)
        usage_hours <- append(usage_hours, num_hours, after=length(usage_hours))
      }# end if
    }# end for
    
    # create usage_report dataframe with hours
    temp_report$Hours <- usage_hours
    usage_report <- aggregate(cbind(Price, Hours) ~ Supervisor + Confocal, data=temp_report, sum)
    
    # create user_report
    user_report <- aggregate(cbind(Price, Hours) ~ Supervisor + Fullname + Confocal, data=temp_report, sum)
  }# end if 
  else if (names(data_df)[1] == "EpiCalcium"){
    # EpiCalcium instrument usage report
    
    # extract supversior and instrument columns  
    temp_report <- subset(data_df, select=c("Supervisor", "Fullname", "EpiCalcium", "Price"))  
    # calculate the number of hours spent per task
    # preallocate usage_hours vector
    usage_hours <- c()
      
    for (item in 1:length(price_list)){
      if (names(price_list[item]) %in% data_df$EpiCalcium){
        # extract data for EpiCalcium
        temp <-subset(data_df, EpiCalcium == names(price_list)[item])
        num_hours <- compute_timespent(temp)
        usage_hours <- append(usage_hours, num_hours, after=length(usage_hours))
      }# end if
    }# end for
    
    # create usage_report dataframe with hours
    temp_report$Hours <- usage_hours
    usage_report <- aggregate(cbind(Price, Hours) ~ Supervisor + EpiCalcium, data=temp_report, sum)
    
    # create user_report
    user_report <- aggregate(cbind(Price, Hours) ~ Supervisor + Fullname + EpiCalcium, data=temp_report, sum)
  }# end else if
  else if(names(data_df)[1] == "ZeissEpi"){
    # ZeissEpi instrument usage report
      
    # extract supversior and instrument columns  
    temp_report <- subset(data_df, select=c("Supervisor", "Fullname", "ZeissEpi", "Price")) 
    # calculate the number of hours spent per task
    # preallocate usage_hours vector
    usage_hours <- c()
      
    for (item in 1:length(price_list)){
      if (names(price_list[item]) %in% data_df$ZeissEpi){
        # extract data for ZeissEpi
        temp <-subset(data_df, ZeissEpi == names(price_list)[item])
        num_hours <- compute_timespent(temp)
        usage_hours <- append(usage_hours, num_hours, after=length(usage_hours))
      }# end if
    }# end for
    
    # create usage_report dataframe with hours
    temp_report$Hours <- usage_hours
    usage_report <- aggregate(cbind(Price, Hours) ~ Supervisor + ZeissEpi, data=temp_report, sum)
    
    # create user_report
    user_report <- aggregate(cbind(Price, Hours) ~ Supervisor + Fullname + ZeissEpi, data=temp_report, sum)
  }# end else if
    
  # calculate total price for each supervisor
  payment_report <- aggregate(Price ~ Supervisor, data=data_df, sum)
  
  # create list to return multiple reports
  return_list <- list(usage_report, user_report, payment_report)

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
  user_report <- data.frame(reports[2])
  user_report <- user_report[order(user_report[,1]),] # sort df by supervisor
  payment_report <- data.frame(reports[3])
  
  # export reports
  wb <- createWorkbook()
  addWorksheet(wb, "Instrument Usage Report")
  addWorksheet(wb, "User Report")
  addWorksheet(wb, "Payment Totals")
  
  writeData(wb, "Instrument Usage Report", usage_report)
  writeData(wb, "User Report", user_report)
  writeData(wb, "Payment Totals", payment_report)
  
  saveWorkbook(wb, file=savename, overwrite=TRUE)
  cat("\nFile:", savename, " has been created\n")
}

#--------------------------------------------------------------------------------
# Main Script
#--------------------------------------------------------------------------------

# check if csv files to import exists
# generate report if files exist

# Conofocal Instrument
if (file_test("-f", confocal_file)){
  # read csv file
  confocal_df <- read.csv(file=confocal_file, header=TRUE, stringsAsFactors=FALSE)
  # preprocessing dataframe
  confocal_df <- dataframe_preprocessing(confocal_df)
  # check if dataframe is empty
  if (empty(confocal_df)){cat("\n", names(confocal_df)[1], 
                              " has no data. Usage report was not created\n")
    } else{
    # retrieve reports from instrument_usage function
    confocal_reports <- instrument_usage(confocal_df, confocal_price_list)
    # export report
    export_reports(confocal_reports, confocal_savename)
  }# end else
}# end if

# Epi & Calcium Instrument
if (file_test("-f", epicalcium_file)){
  # read csv file
  epicalcium_df <- read.csv(file=epicalcium_file, header=TRUE, stringsAsFactors=FALSE)
  # preprocessing dataframe
  epicalcium_df <- dataframe_preprocessing(epicalcium_df)
  # check if dataframe is empty
  if (empty(epicalcium_df)){cat("\n", names(epicalcium_df)[1], 
                                " has no data. Usage report was not created.\n")
  } else{
    # retrieve reports from instrument_usage function
    epicalcium_reports <- instrument_usage(epicalcium_df, epicalcium_price_list)
    # export report
    export_reports(epicalcium_reports, epicalcium_savename)
  }# end else
}# end if

# Zeiss Epi Instrument
if (file_test("-f", zeissepi_file)){
  # read csv file
  zeissepi_df <- read.csv(file=zeissepi_file, header=TRUE, stringsAsFactors=FALSE)
  # preprocessing dataframe
  zeissepi_df <- dataframe_preprocessing(zeissepi_df)
  # check if dataframe is empty
  if (empty(zeissepi_df)){cat("\n", names(zeissepi_df)[1], 
                              " has no data. Usage report was not created.\n")
  } else{
    # retrieve reports from instrument_usage function
    zeissepi_reports <- instrument_usage(zeissepi_df, zeissepi_price_list)
    # export report
    export_reports(zeissepi_reports, zeissepi_savename)
  }# end else
}# end if
