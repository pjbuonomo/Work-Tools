
library(dplyr)
library(DBI)
library(RODBC)
library(readr)
library(readxl)

db_conn <- odbcConnect("ILS")

#################################################################################################################

#FILE NEEDS TO BE CLEANED BEFORE USE
table <- read_excel("S:/Touchstone/Catrader/Boston/Database/CatBond/Shiny/Aon_Weekly/Aon20231229Clean.xlsx", sheet = 'RLS', col_names=T, skip=1) %>%
		data.frame()

#DATE NEEDS TO BE MANUALLY CHANGED
QDate <- rep(as.Date("2023-12-29"), times=nrow(table))
table <-cbind(QDate, table)

table <- table %>%
		select(-4,-13,-14,-15)

ColNames <- sqlColumns(db_conn, "Aon") %>% 
			select('COLUMN_NAME')

colnames(table) <- ColNames$COLUMN_NAME

table$LongTermAsk[is.na(table$LongTermAsk)] <- 0
table$LongTermEL[is.na(table$LongTermEL)] <- 0
table$NearTermAsk[is.na(table$NearTermAsk)] <- 0
table$NearTermEL[is.na(table$NearTermEL)] <- 0
table$BidSpread[table$BidSpread == 'n/a'] <- 0
table$OfferSpread[table$OfferSpread == 'n/a'] <- 0

table$Coupon <- table$Coupon %>% parse_number() %>% as.integer()
table$BidPrice <- table$BidPrice %>% parse_number()
table$OfferPrice <- table$OfferPrice %>% parse_number()

table$Size <- formatC(c(table$Size), digits = 2, format = 'f')
table$LongTermAsk <- formatC(c(table$LongTermAsk), digits = 2, format = 'f')
table$LongTermEL <- formatC(c(table$LongTermEL), digits = 2, format = 'f')
table$NearTermAsk <- formatC(c(table$NearTermAsk), digits = 2, format = 'f')
table$NearTermEL <- formatC(c(table$NearTermEL), digits = 2, format = 'f')
table$Coupon <- formatC(c(table$Coupon), digits = 2, format = 'f')
table$BidPrice <- formatC(c(table$BidPrice), digits = 2, format = 'f')
table$OfferPrice <- formatC(c(table$OfferPrice), digits = 2, format = 'f')

sqlSave(db_conn, table, tablename="Aon", rownames=F, append=T, fast=F)

##################################################################################################################################

odbcClose(db_conn)

##################################################################################################################################
