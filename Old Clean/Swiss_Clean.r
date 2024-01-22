
library(dplyr)
library(DBI)
library(RODBC)
library(readr)
library(readxl)

db_conn <- odbcConnect("ILS")

#################################################################################################################

table <- read_excel("S:/Touchstone/Catrader/Boston/Database/CatBond/Shiny/Weekly_Pricing/ILS pricing data 20231229.xlsx", col_names=T, skip=12) %>%
		na.omit() %>%
		data.frame()

ColNames <- sqlColumns(db_conn, "SwissRe") %>% 
			select('COLUMN_NAME')

result <- table %>%
		select(-3, -4, -6, -8, -11, -12)

colnames(result) <- ColNames$COLUMN_NAME

for (i in 1:nrow(result)) {
	if (nchar(result$CUSIP[i]) > 9) {
		result$CUSIP[i] <- substr(result$CUSIP[i], 3, 11)
	}
}

result$EL <- result$EL %>% parse_number() %>% as.integer()
result$Coupon <- result$Coupon %>% parse_number() %>% as.integer()
sqlSave(db_conn, result, tablename="SwissRe", rownames=F, append=T, fast=F, verbose = TRUE)

##################################################################################################################################

odbcClose(db_conn)

##################################################################################################################################
