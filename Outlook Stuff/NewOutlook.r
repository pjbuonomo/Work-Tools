library(RDCOMClient)
library(tools) # Load the 'tools' library for file operations

# Create an Outlook application object
Outlook <- COMCreate("Outlook.Application")
myNameSpace <- Outlook$GetNameSpace("MAPI")

# Access the Inbox and then the specific subfolder
inboxFolderIndex <- 1 # Adjust based on your Outlook setup
inbox <- myNameSpace$Folders(inboxFolderIndex)$Folders("Inbox")
bhCatBondFolder <- inbox$Folders("BH Cat Bond")

# Create a directory to store the text files
outputDir <- "S:/Touchstone/Catrader/Boston/Database/UnreadEmails"
dir.create(outputDir, showWarnings = FALSE)

# ... [previous code remains the same]

# Loop to process unread messages

# Loop to process unread messages
# Loop to process unread messages
for (i in 1:bhCatBondFolder$Items()$Count()) {
    message <- bhCatBondFolder$Items()$Item(i)
    
    # Process only if the message is unread
    if (message$UnRead() == TRUE) {
        # Create a unique filename for the text file
        timestamp <- format(Sys.time(), "%Y%m%d%H%M%S")
        filename <- file.path(outputDir, paste("email_", timestamp, ".txt", sep = ""))
        
        # Create a text stream and save the email as a text file
        textStream <- message$GetInspector()$WordEditor()$Content$Text
        cat(textStream, file = filename)
        
        # Mark the message as read (optional)
        message$UnRead(FALSE)
        message$Save()
    }
}


# Process the saved text files and extract their content
emailFiles <- list.files(path = outputDir, pattern = "*.txt", full.names = TRUE)

# Initialize a data frame to store email details
emails_df <- data.frame(Timestamp = character(),
                        Subject = character(),
                        Content = character(),
                        stringsAsFactors = FALSE)

for (emailFile in emailFiles) {
    # Read the text from the file
    emailContent <- readLines(emailFile, warn = FALSE)
    
    # Retrieve and format the ReceivedTime
    receivedTime <- file.info(emailFile)$mtime
    formattedTime <- format(as.POSIXct(receivedTime), "%Y-%m-%d %H:%M:%S")
    
    # Add email details to the data frame
    emails_df <- rbind(emails_df, data.frame(Timestamp = formattedTime,
                                             Subject = "Email Subject", # Modify as needed
                                             Content = paste(emailContent, collapse = "\n")))
    
    # Delete the text file
    unlink(emailFile)
}

# Write the data frame to a CSV file
write.csv(emails_df, file = "S:/Touchstone/Catrader/Boston/Database/UnreadDatabaseEntryEmails.csv", row.names = FALSE)
