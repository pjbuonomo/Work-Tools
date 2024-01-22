
library(dplyr)
library(DBI)
library(RODBC)
library(readr)
library(readxl)

db_conn <- odbcConnect("ILS")

#################################################################################################################

table <- read_excel("S:/Touchstone/Catrader/Boston/Database/CatBond/Shiny/GCS_Weekly/GCS20231229.xlsx", sheet = 2, col_names=F, skip=6) %>%
		data.frame()

#Add date to front of table (Needs to be manually edited)
Date <- rep(as.Date("2023-12-29"), times=nrow(table))
table <-cbind(Date, table)

ColNames <- sqlColumns(db_conn, "GCS") %>% 
			select('COLUMN_NAME')

colnames(table) <- ColNames$COLUMN_NAME

#Limit BidSpread and OfferSpread to 2 decimal places and convert any N/A values to 0
table$BidSpread[is.na(table$BidSpread)] <- 0
table$OfferSpread[is.na(table$OfferSpread)] <- 0
table$BidSpread <- formatC(c(table$BidSpread), digits = 2, format = 'f')
table$OfferSpread <- formatC(c(table$OfferSpread), digits = 2, format = 'f')

sqlSave(db_conn, table, tablename="GCS", rownames=F, append=T, fast=F)

##################################################################################################################################

odbcClose(db_conn)

##################################################################################################################################
