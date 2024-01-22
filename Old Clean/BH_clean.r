library(tidyr)
library(stringr)

#######################################################################

# Add select button, long(CUSIP) format

#BH2 <- read.csv(file.choose(), stringsAsFactor = FALSE, header=FALSE, skip=1)[,2]
#BH_long(BH2)


BH_long <- function(BH2) {

	# Seperate each line of message
    BH2 <- do.call(cbind, strsplit(BH2, split="<br>"))

	# Avoid descrption of message
	BH2 <- as.data.frame(BH2[!str_detect(BH2[,1], '\\*'),])
	colnames(BH2) <- "Message"

	# Detect the 9-digits CUSIP
	message <- strsplit(BH2$Message, split="\\ ")
	CUSIP <- rep("", nrow(BH2))
	
	for (i in 1:length(message)) {
		CUSIP[i] <- message[[i]][1]
		if (nchar(CUSIP[i]) > 9) {
			CUSIP[i] <- substr(CUSIP[i], 3, 11) 
		}
	}

	# Split by @ to get the price and the rest of message
	temp <- strsplit(as.character(BH2[['Message']]), split="@")

	columns <- c("CUSIP", "Name", "Size", "Action", "Price")
	BH2_Clean <- data.frame(matrix(nrow = 0, ncol = length(columns)))
	colnames(BH2_Clean) <- columns
	
	last_word <- function(string) {
		result <- substr(string, tail(unlist(gregexpr(" ", string)), n=1)+1, nchar(string))
		return (result)
	}

	for (i in 1:length(temp)) {
		rest <- temp[[i]][1]
		Action <- last_word(substr(rest,1,nchar(rest)-1))
		Name <- substr(rest, nchar(CUSIP[i])+5, nchar(rest)-nchar(Action)-1)
		Price <- as.numeric(temp[[i]][2])
		row <- cbind(CUSIP[i], Name, "0", Action, Price)
		BH2_Clean[i,] <- row
	}

	Date <- rep(Sys.Date(), nrow(BH2_Clean))
	Broker <- rep("BeechHill", nrow(BH2_Clean))
	BH2_Clean <- cbind(Broker, BH2_Clean)
	BH2_Clean <- cbind(Date, BH2_Clean)

	return (BH2_Clean)


}


#######################################################################

# Add select button, long(Name) format

#BH2 <- read.csv(file.choose(), stringsAsFactor = FALSE, header=FALSE, skip=1)[,2]
#BH_long2(BH2)


BH_long2 <- function(BH2) {

    # Seperate each line of message
	BH2 <- do.call(cbind, strsplit(BH2, split="<br>"))

	BH2 <- as.data.frame(BH2[!str_detect(BH2[,1], '\\*'),])
	colnames(BH2) <- "Message"

	temp <- strsplit(BH2[['Message']], split="@")

	columns <- c("CUSIP", "Name", "Size", "Action", "Price")
	BH2_Clean <- data.frame(matrix(nrow = 0, ncol = length(columns)))
	colnames(BH2_Clean) <- columns

	for (i in 1:length(temp)) {
		rest <- temp[[i]][1]
		rest_split <- unlist(strsplit(rest, split="\\ "))

		Price <- as.numeric(temp[[i]][2])
		Action <- tail(rest_split,1)
		CUSIP <- tail(rest_split,2)[1]
		
		if (nchar(CUSIP) > 9) {
			CUSIP <- substr(CUSIP, 3, 11) 
		}

		len <- length(rest_split)-2
		curr <- rest_split[1]

		for (j in 2:len){
			Name <- paste(curr,rest_split[j])
			curr <- Name
		}

		row <- cbind(CUSIP, Name, "0", Action, Price)
		BH2_Clean[i,] <- row
	}

	Date <- rep(Sys.Date(), nrow(BH2_Clean))
	Broker <- rep("BeechHill", nrow(BH2_Clean))
	BH2_Clean <- cbind(Broker, BH2_Clean)
	BH2_Clean <- cbind(Date, BH2_Clean)

	return (BH2_Clean)


}


#######################################################################

# Add select button, size format

#BH <- read.csv(file.choose(), stringsAsFactor = FALSE, header=FALSE, skip=1)[,2]

#BH_short(BH)

BH_short <- function(BH) {
	
	# Seperate each line of message
	BH <- do.call(cbind, strsplit(BH, split="<br>"))

	Number <- gsub(" .*$","",BH[,1])

	# Convert k and mm to numeric value
	Number[grepl("k$", Number)] <- as.numeric(sub("k$", "", Number[grepl("k$", Number)]))*10^3
	Number[grepl("mm$", Number)] <- as.numeric(sub("mm$", "", Number[grepl("mm$", Number)]))*10^6

	# remove price from whole string
	rest <- rep(NA, nrow(BH))

	for (i in 1:nrow(BH)){
		rest[i] <- substr(BH[[i, 1]], unlist(gregexpr('k|mm', BH[[i, 1]]))+2, nchar(BH[[i, 1]]))
		if (unlist(gregexpr(" ", rest[i]))[1] == 1) {
			rest[i] <- substr(rest[i], 2, nchar(rest[i]))
		}
	}

	# Extract name
	Name <- do.call(rbind, strsplit(rest, split=" \\("))[,1]
	rest <- do.call(rbind, strsplit(rest, split=" \\("))[,2]

	# Extract ISIN
	CUSIP <- do.call(rbind, strsplit(rest, split="\\) "))[,1]
	rest <- do.call(rbind, strsplit(rest, split="\\) "))[,2]

	# Extract price and action
	rest <- as.data.frame(rest)
	columns <- c("CUSIP", "Name", "Size", "Action", "Price")
	BH_Clean <- data.frame(matrix(nrow = 0, ncol = length(columns)))
	colnames(BH_Clean) <- columns

	for (i in 1:nrow(rest)) {

		if (str_detect(rest[i,], "\\*")) {
			rest[i,] <- do.call(rbind, strsplit(rest[i,], split="\\*"))[,1]
		}

		if (nchar(CUSIP[i]) > 9) {
			CUSIP[i] <- substr(CUSIP[i], 3, 11) 
		}

		# Whether the message contains only one quote with @ or bid and offer with /
		if (str_detect(rest[i,], "@")) {
			Action <- unlist(strsplit(rest[i,], split="\\ @ "))[1]
			if (Action == "offered") {
				Action = "offer"
			}
			Price <- as.numeric(unlist(strsplit(rest[i,], split="\\ @ "))[2])
			row <- cbind(CUSIP[i], Name[i], Number[i], Action, Price)
			BH_Clean[nrow(BH_Clean)+1,] <- row
		} else {
			# If there are bid and offer in one message
			temp <- unlist(strsplit(rest[i,], split="\\ "))
			temp <- temp[(temp != "-") & (temp != "/")]
			Action <- temp[2]
			if (Action == "offered") {
				Action = "offer"
			}
			Price <- as.numeric(temp[1])
			row <- cbind(CUSIP[i], Name[i], Number[i], Action, Price)
			BH_Clean[nrow(BH_Clean)+1,] <- row
			# Add second row to the data frame
			Action <- temp[4]
			if (Action == "offered") {
				Action = "offer"
			}
			Price <- as.numeric(temp[3])
			row <- cbind(CUSIP[i], Name[i], Number[i], Action, Price)
			BH_Clean[nrow(BH_Clean)+1,] <- row
		}
	}


	Date <- rep(Sys.Date(), nrow(BH_Clean))
	Broker <- rep("BeechHill", nrow(BH_Clean))
	BH_Clean <- cbind(Broker, BH_Clean)
	BH_Clean <- cbind(Date, BH_Clean)

	return (BH_Clean)


}
