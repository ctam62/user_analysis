"
Monthly Instrument Usage Tracking

Written by Clara Tam
Copyright (c) 2020
Licensed under the MIT License (see LICENSE for details)

"
#--------------------------------------------------------------------------------
# Import required library packages
#--------------------------------------------------------------------------------
library(rstudioapi)
library(plyr)
library(stringr)
library(openxlsx)
library(data.table)
library(lubridate)

#--------------------------------------------------------------------------------
# Setup working environment
#--------------------------------------------------------------------------------
# Set working directory
setwd(dirname(rstudioapi::getActiveDocumentContext()$path))
working_dir <- getwd()
cat("Your working directory is set to:", working_dir, "\n")

#--------------------------------------------------------------------------------
# Import files
#--------------------------------------------------------------------------------
# batch import csv files
csv_files = list.files(pattern="*.csv")

#--------------------------------------------------------------------------------
# Instrument Fees (Update as necessary)
#--------------------------------------------------------------------------------
instrumentA_price_list <- c(
  "Training" = 75, 
  "User" = 50, 
  "Data Analysis" = 0, 
  "Full Service" = 50,
  "External Training" = 150,
  "External User" = 150,
  "External Full Service" = 75
  )

instrumentB_price_list <- c(
  "Training" = 75, 
  "User" = 25, 
  "Data Analysis" = 0, 
  "Full Service" = 50,
  "External Training" = 150,
  "External User" = 60,
  "External Full Service" = 75
  )

instrumentC_price_list <- c(
  "Training" = 75, 
  "User" = 25, 
  "Data Analysis" = 0, 
  "Full Service" = 50,
  "External Training" = 150,
  "External User" = 60,
  "External Full Service" = 75
  )

price_list <- list("Instrument_A" = instrumentA_price_list, 
                   "Instrument_B" = instrumentB_price_list,
                   "Instrument_C" = instrumentC_price_list)

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
  
  if(names(data_df)[1] == "InstrumentA"){
    # Confocal instrument usage report
    df_first_column <- data_df$InstrumentA
    
    # calculate the number of hours spent per task
    # preallocate usage_hours vector
    usage_hours <- c()
    for (item in 1:length(price_list)){
      if (names(price_list[item]) %in% df_first_column){
        # extract data for Instrument
        temp <- subset(data_df, InstrumentA == names(price_list)[item])
        num_hours <- compute_timespent(temp)
        usage_hours <- append(usage_hours, num_hours, after=length(usage_hours))
      }# end if
    }# end for
    
    usage_formula <- cbind(Price, Hours) ~ Supervisor + InstrumentA
    user_formula <- cbind(Price, Hours) ~ Supervisor + Fullname + InstrumentA
    report_col_names <- c("Supervisor", "Fullname", "InstrumentA", "Price")
  
  }else if (names(data_df)[1] == "InstrumentB"){
    # EpiCalcium instrument usage report  
    df_first_column <- data_df$InstrumentB
    
    # calculate the number of hours spent per task
    # preallocate usage_hours vector
    usage_hours <- c()
    for (item in 1:length(price_list)){
      if (names(price_list[item]) %in% df_first_column){
        # extract data for Instrument
        temp <- subset(data_df, InstrumentB == names(price_list)[item])
        num_hours <- compute_timespent(temp)
        usage_hours <- append(usage_hours, num_hours, after=length(usage_hours))
      }# end if
    }# end for
    
    usage_formula <- cbind(Price, Hours) ~ Supervisor + InstrumentB
    user_formula <- cbind(Price, Hours) ~ Supervisor + Fullname + InstrumentB
    report_col_names <- c("Supervisor", "Fullname", "InstrumentB", "Price")
    
  }else if(names(data_df)[1] == "InstrumentC"){
    # ZeissEpi instrument usage report
    df_first_column <- data_df$InstrumentC
    
    # calculate the number of hours spent per task
    # preallocate usage_hours vector
    usage_hours <- c()
    for (item in 1:length(price_list)){
      if (names(price_list[item]) %in% df_first_column){
        # extract data for Instrument
        temp <- subset(data_df, InstrumentC == names(price_list)[item])
        num_hours <- compute_timespent(temp)
        usage_hours <- append(usage_hours, num_hours, after=length(usage_hours))
      }# end if
    }# end for
    
    usage_formula <- cbind(Price, Hours) ~ Supervisor + InstrumentC
    user_formula <- cbind(Price, Hours) ~ Supervisor + Fullname + InstrumentC
    report_col_names <- c("Supervisor", "Fullname", "InstrumentC", "Price")
  }
  
  # extract supversior and instrument columns  
  temp_report <- subset(data_df, select=report_col_names)
    
  # create usage_report dataframe with hours
  temp_report$Hours <- usage_hours
  usage_report <- aggregate(usage_formula, data=temp_report, sum)
    
  # create user_report
  user_report <- aggregate(user_formula, data=temp_report, sum)
    
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
  usage_report <- data.frame(reports[1], stringsAsFactors=FALSE)
  usage_report <- usage_report[order(usage_report[,1]),] # sort df by supervisor
  user_report <- data.frame(reports[2], stringsAsFactors=FALSE)
  user_report <- user_report[order(user_report[,1]),] # sort df by supervisor
  payment_report <- data.frame(reports[3], stringsAsFactors=FALSE)
  
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
# batch read csv files
csv_data = lapply(csv_files,fread)

# obtain dates and instrument names
dates <- gsub("[A-z.& ]", "", csv_files)
instruments <- gsub("[0-9 -]|+.csv", "", csv_files)

for(item in 1:length(csv_files)){
  # batch generate output directory and filenames
  output_dir <- file.path(working_dir, paste0(dates[item],"_reports"))
  if (!dir.exists(output_dir)){dir.create(output_dir)}
  savename <- paste0(output_dir, "/", dates[item], "_", instruments[item],"_usage_report.xlsx")
  
  # batch process and generate reports
  imported_df <- data.frame(csv_data[item], stringsAsFactors=FALSE)
  preprocessed_df <- dataframe_preprocessing(imported_df)
  # check if dataframe is empty
  if (empty(preprocessed_df)){cat("\n", names(preprocessed_df)[1], 
                              " has no data. Usage report was not created.\n")
  } else{
    # retrieve reports from instrument_usage function
    index <- match(instruments[item], names(price_list))
    instrument_reports <- instrument_usage(preprocessed_df, price_list[[index]])
    # export report
    export_reports(instrument_reports, savename)
  }# end else
}# end for
