library(tidyr)
library(stringr)
library(dplyr)

options(scipen=999)


#######################################################################

#RBC <- read.csv(file.choose(), stringsAsFactor = FALSE, header=FALSE, skip=1)[,2]
	
#RBC_clean(RBC)

RBC_clean <- function(RBC) {

	RBC <- do.call(cbind, strsplit(RBC, split="<br>"))

	# find first occurance of space from left to right, the first column is ask or bid
	Action <- rep(NA, nrow(RBC))
	rest <- rep(NA, nrow(RBC))
	trader <- gsub(" .*$","",RBC)
	
	for (i in 1:nrow(RBC)) {
		if (trader[i,1] == "Seller") {
			Action[i] <- "offer"
		} else if (trader[i,1] == "Buyer") {
			Action[i] <- "bid"
		} else if (trader[i,1] == "Market"){
			Action[i] <- "market"
		} else {
			Action[i] <- ""
		}
		rest[i] <- substr(RBC[i], nchar(trader[i,1])+2, nchar(RBC[i]))
	}


	# then find first occurance of space from right to left, 1st column = transfer to size, pay attenton on "+"
	Number <- rep(0, nrow(RBC))
	last_word <- function(string) {
		result <- substr(string, tail(unlist(gregexpr(" ", string)), n=1)+1, nchar(string))
		return (result)
	}

	for (i in 1:nrow(RBC)) {

		if (substr(rest[i], nchar(rest[i]), nchar(rest[i])) == " ") {
			rest[i] <- substr(rest[i], 1, nchar(rest[i])-1)
		}
		
		last <- last_word(rest[i])
		
		if ((length(grep("k|K|m|M", last_word(last))) == 0) || (nchar(last) > 7)) {
			Number[i] = 0
			next
		}

		Number[i] <- last_word(rest[i]) %>%
					str_replace("K", "k") %>%
					str_replace("M", "m")

		check <- str_sub(Number[i], -1)

		# Delete "+" or other character after k and mm
		if (check != "m" && check != "k") {
			Number[i] <- str_sub(Number[i], 1, -2)
		}

		rest[i] <- substr(rest[i], 1, nchar(rest[i])-nchar(Number[i])-1)
		
		while (last_word(rest[i]) == "") {
			rest[i] <- str_sub(rest[i], 1, -2)
		}
	}

	Number[grepl("k$", Number)] <- as.numeric(sub("k$", "", Number[grepl("k$", Number)]))*10^3
	Number[grepl("mm$", Number)] <- as.numeric(sub("mm$", "", Number[grepl("mm$", Number)]))*10^6


	# then r-l next first space = price
	Price <- rep(NA, nrow(RBC))
	Name <- rep(NA, nrow(RBC))
	ISIN_CUSIP <- rep(NA, nrow(RBC))
	check <- rep(NA, nrow(RBC))

	for (i in 1:nrow(RBC)) {
		check[i] <- !is.na(as.numeric(last_word(rest[i])))

		if (check[i] == TRUE) {

			# Last element is price, second last is ISIN
			Price[i] <- as.numeric(last_word(rest[i]))
			temp <- substr(rest[i], 1, nchar(rest[i])-nchar(Price[i])-1)

			while (last_word(temp) == "") {
				temp <- str_sub(temp, 1, -2)
			}

			if (nchar(last_word(temp))<8) {
				ISIN_CUSIP[i] <- ""
				Name[i] <- temp
			} else {
				ISIN_CUSIP[i] <- last_word(temp)
				if (nchar(ISIN_CUSIP[i]) > 9) {
					ISIN_CUSIP[i] <- substr(ISIN_CUSIP[i], 3, 11)
				}
				Name[i] <- substr(temp, 1, nchar(temp)-nchar(ISIN_CUSIP[i])-1)
			}

		} else {

			# Last element is ISIN, second last is price
			ISIN_CUSIP[i] <- last_word(rest[i])
		
			if (nchar(ISIN_CUSIP[i]) > 9) {
				ISIN_CUSIP[i] <- substr(ISIN_CUSIP[i], 3, 11)
			}
			
			temp <- substr(rest[i], 1, nchar(rest[i])-nchar(ISIN_CUSIP[i])-1)

			Price[i] <- as.numeric(last_word(temp))
			Name[i] <- substr(temp, 1, nchar(temp)-nchar(Price[i])-1)
		}
	}


	# then rest is the name of bond

	Date <- rep(Sys.Date(), nrow(RBC))
	Broker <- rep("RBC", nrow(RBC))

	Acorn <- data.frame(Date, Broker, ISIN_CUSIP, Name, Size=Number, Action, Price)
	Acorn <- na.omit(Acorn)
	
	return (Acorn)
	
}

