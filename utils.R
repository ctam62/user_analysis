import("stats", "aggregate")
import("openxlsx")
import("lubridate", "hm")


dataframe_preprocessing <- function(data_df){
  "
  Remove unwanted columns and whitespaces in column header names
  
  Arguments:
  data_df -- data frame, of shape (num_obs, num_variables)
  
  Returns:
  data_df -- data frame, of shape (num_obs, num_variables)
  "
  # remove nonalphanumeric characters from header names
  names(data_df) <- gsub("[^[:alnum:]///' ]", "", names(data_df))
  # standardize supervisor names by removing any punctuations
  data_df$Supervisor <- gsub("[[:punct:]]", "", data_df$Supervisor)
  
  return(data_df)
}

compute_timespent <-function(data_df){
  "
  Comput the number of hours spent per user
  
  Arguments:
  data_df -- data frame, of shape (num_obs, num_variables)
  
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
  Compiles usage, user, and payment reports
  
  Arguments:
  data_df -- input data frame, of shape (num_obs, num_variables)
  price_list -- named atomic vector type numeric, of shape (1, num_instruments)
  
  Return:
  usage_report -- data frame, of shape (num_obs, 4)
  payment_report -- data frame, of shape (num_obs, 2)
  "
  
  if(names(data_df)[1] == "InstrumentA"){
    # InstrumentA instrument usage report
    df_first_column <- data_df$InstrumentA
    
    # calculate the number of hours spent per task
    # preallocate usage_hours vector
    usage_hours <- c()
    for(item in 1:length(price_list)){
      if(names(price_list[item]) %in% df_first_column){
        # extract data for Instrument
        temp <- subset(data_df, InstrumentA == names(price_list)[item])
        num_hours <- compute_timespent(temp)
        usage_hours <- append(usage_hours, num_hours, after=length(usage_hours))
      }# end if
    }# end for
    
    usage_formula <- cbind(Price, Hours) ~ Supervisor + InstrumentA
    user_formula <- cbind(Price, Hours) ~ Supervisor + Fullname + InstrumentA
    report_col_names <- c("Supervisor", "Fullname", "InstrumentA", "Price")
    
  }else if(names(data_df)[1] == "InstrumentB"){
    # InstrumentB instrument usage report  
    df_first_column <- data_df$InstrumentB
    
    # calculate the number of hours spent per task
    # preallocate usage_hours vector
    usage_hours <- c()
    for(item in 1:length(price_list)){
      if(names(price_list[item]) %in% df_first_column){
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
    # InstrumentC instrument usage report
    df_first_column <- data_df$InstrumentC
    
    # calculate the number of hours spent per task
    # preallocate usage_hours vector
    usage_hours <- c()
    for(item in 1:length(price_list)){
      if(names(price_list[item]) %in% df_first_column){
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
  
  # extract supervisor and instrument columns  
  temp_report <- subset(data_df, select=report_col_names)
  
  # create usage_report data frame with hours
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
  Export reports to an excel workbook
  
  Arguments:
  reports -- list, of shape (3, 1)
  savename -- string, of shape (1, 1)
  "
  # convert individual reports back to data frames
  usage_report <- data.frame(reports[1], stringsAsFactors=FALSE)
  usage_report <- usage_report[order(usage_report[,1]),] # sort df by supervisor
  user_report <- data.frame(reports[2], stringsAsFactors=FALSE)
  user_report <- user_report[order(user_report[,1]),] # sort df by supervisor
  payment_report <- data.frame(reports[3], stringsAsFactors=FALSE)
  
  # create a workbook object and add worksheets
  wb <- createWorkbook()
  addWorksheet(wb, "Instrument Usage Report")
  addWorksheet(wb, "User Report")
  addWorksheet(wb, "Payment Totals")
  
  # write data into the workbook object
  writeData(wb, "Instrument Usage Report", usage_report)
  writeData(wb, "User Report", user_report)
  writeData(wb, "Payment Totals", payment_report)
  
  # save the workbook
  saveWorkbook(wb, file=savename, overwrite=TRUE)
  cat("\nFile:", savename, " has been created\n")
}
