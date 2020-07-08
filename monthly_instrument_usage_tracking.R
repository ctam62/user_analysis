"
Monthly Instrument Usage Tracking

Written by Clara Tam
Copyright (c) 2020
Licensed under the MIT License (see LICENSE for details)

"
#--------------------------------------------------------------------------------
# Import required library packages
#--------------------------------------------------------------------------------
import::from(rstudioapi, getActiveDocumentContext)
import::from(data.table, fread)
library(plyr)
library(stringr)
library(modules)

#--------------------------------------------------------------------------------
# Setup working environment
#--------------------------------------------------------------------------------
# Set working directory
setwd(dirname(getActiveDocumentContext()$path))
working_dir <- getwd()
cat("Your working directory is set to:", working_dir, "\n")

#--------------------------------------------------------------------------------
# Import files
#--------------------------------------------------------------------------------
# batch import csv files from data directory
csv_files <- Sys.glob(paste0(working_dir,'/data/*.csv'))
csv_filenames <- gsub(".*/", "", csv_files)

#--------------------------------------------------------------------------------
# Instrument Fees (Update as necessary)
#--------------------------------------------------------------------------------
# define the price list for each instrument
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

# store instrument specific price lists in price_list
price_list <- list("Instrument_A" = instrumentA_price_list, 
                   "Instrument_B" = instrumentB_price_list,
                   "Instrument_C" = instrumentC_price_list)

#--------------------------------------------------------------------------------
# Main Script
#--------------------------------------------------------------------------------
# import helper functions from utils module
utils <- use("utils.R")

# batch read csv files
csv_data = lapply(csv_files, fread)

# obtain dates and instrument names from csv filenames
dates <- gsub("[A-z.&/: ]", "", csv_filenames)
instruments <- gsub("[0-9 -]|+.csv", "", csv_filenames)

for(item in 1:length(csv_files)){
  # generate output directory and filenames
  cat("Generating reports for", instruments[item],"....")
  output_dir <- file.path(working_dir, paste0(dates[item],"_reports"))
  if(!dir.exists(output_dir)){dir.create(output_dir)}
  savename <- paste0(output_dir, "/", dates[item], "_", instruments[item],"_usage_report.xlsx")
  
  # generate reports
  imported_df <- data.frame(csv_data[item], stringsAsFactors=FALSE)
  preprocessed_df <- utils$dataframe_preprocessing(imported_df)
  # check if data frame is empty
  if (empty(preprocessed_df)){cat("\n", names(preprocessed_df)[1], 
                              " has no data. Usage report was not created.\n")
  }else{
    # retrieve reports from instrument_usage function
    index <- match(instruments[item], names(price_list))
    instrument_reports <- utils$instrument_usage(preprocessed_df, price_list[[index]])
    # export reports
    utils$export_reports(instrument_reports, savename)
  }# end else
}# end for
