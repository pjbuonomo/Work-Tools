
library(dplyr)
library(DBI)
library(RODBC)
library(readr)
library(readxl)

db_conn <- odbcConnect("ILS")

#################################################################################################################

#SPREADSHEET NEEDS TO BE MANUALLY CLEANED BEFORE THIS SCRIPT IS RUN
table <- read_excel("S:/Touchstone/Catrader/Boston/Database/CatBond/Shiny/BH_Weekly/BH20231229Clean.xlsx", col_names=T, skip=1) %>%
		na.omit() %>%
		data.frame()

#Add date to front of table (Needs to be manually edited)
Date <- rep(as.Date("2023-12-29"), times=nrow(table))
table <-cbind(Date, table)

ColNames <- sqlColumns(db_conn, "BH") %>% 
			select('COLUMN_NAME')

result <- table %>%
		select(-15)

colnames(result) <- ColNames$COLUMN_NAME

#Limit fields to 2 decimal places
result$EL <- formatC(c(result$EL), digits = 2, format = 'f')
result$Margin <- formatC(c(result$Margin), digits = 2, format = 'f')
result$Coupon <- formatC(c(result$Coupon), digits = 2, format = 'f')
result$BidPrice <- formatC(c(result$BidPrice), digits = 2, format = 'f')
result$BidDiscountMargin <- formatC(c(result$BidDiscountMargin), digits = 2, format = 'f')
result$OfferPrice <- formatC(c(result$OfferPrice), digits = 2, format = 'f')
result$OfferDiscountMargin <- formatC(c(result$OfferDiscountMargin), digits = 2, format = 'f')
result$Size <- formatC(c(result$Size), digits = 2, format = 'f')

sqlSave(db_conn, result, tablename="BH", rownames=F, append=T, verbose = TRUE)

##################################################################################################################################

odbcClose(db_conn)

##################################################################################################################################
