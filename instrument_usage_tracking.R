"
Microscopy instrument usage tracking

"

# Set working directory
setwd("D:/instrument_usage_tracking")

# import required library packages
library(plyr)
library(dplyr)

#---------------------------------------
# Define the input arguments
#---------------------------------------
# Microscopy Fees (Update as necessary)
price_list <- c(data_analysis=0, user=10, training=50)

#---------------------------------------

# import excel file
data_df <- read.csv(file='2019-8 Zeiss_Epi.csv', header=TRUE, stringsAsFactors=FALSE)
supervisors <- data_df$Supervisor
price <- data_df$Price
instrument <- data_df$ï..Zeiss.Epi

num_supervisors <- count(data_df, c("Supervisor"))
usage_report <- aggregate(Price ~ Supervisor + ï..Zeiss.Epi, data=data_df, sum)

# usage_report <- instrument_usage(price_list){


# }


