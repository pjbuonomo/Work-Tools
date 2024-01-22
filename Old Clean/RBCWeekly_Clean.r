
library(dplyr)
library(DBI)
library(RODBC)
library(readr)
library(readxl)

db_conn <- odbcConnect("ILS")

#################################################################################################################

table <- read_excel("S:/Touchstone/Catrader/Boston/Database/CatBond/Shiny/RBC_Weekly/RBC20231215.xlsx", col_names=T) %>%
		na.omit() %>%
		data.frame()

#Add date to front of table (Needs to be manually edited)
Date <- rep(as.Date("2023-12-15"), times=nrow(table))
table <-cbind(Date, table)

ColNames <- sqlColumns(db_conn, "RBC") %>% 
			select('COLUMN_NAME')

colnames(table) <- ColNames$COLUMN_NAME

sqlSave(db_conn, table, tablename="RBC", rownames=F, append=T, fast=F, verbose = T)

##################################################################################################################################

odbcClose(db_conn)

##################################################################################################################################
