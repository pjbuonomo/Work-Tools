library(RDCOMClient)
library(tools) # Load the 'tools' library for file operations

# Create an Outlook application object
Outlook <- COMCreate("Outlook.Application")
myNameSpace <- Outlook$GetNameSpace("MAPI")

# Access the Inbox and then the specific subfolder
inboxFolderIndex <- 1 # Adjust based on your Outlook setup
inbox <- myNameSpace$Folders(inboxFolderIndex)$Folders("Inbox")
bhCatBondFolder <- inbox$Folders("BH Cat Bond")

# Create a directory to store the email files
outputDir <- "S:/Touchstone/Catrader/Boston/Database/UnreadEmails"
dir.create(outputDir, showWarnings = FALSE)

# Loop to process unread messages
for (i in 1:bhCatBondFolder$Items()$Count()) {
    message <- bhCatBondFolder$Items()$Item(i)
    
    # Process only if the message is unread
    if (message$UnRead() == TRUE) {
        # Create a unique filename for the email file
        timestamp <- format(Sys.time(), "%Y%m%d%H%M%S")
        filename <- file.path(outputDir, paste("email_", timestamp, ".msg", sep = ""))
        
        # Save the entire email as a .msg file
        message$SaveAs(filename, 3) # Use 3 for olMSG format
        
        # Mark the message as read (optional)
        message$UnRead(FALSE)
        message$Save()
    }
}
