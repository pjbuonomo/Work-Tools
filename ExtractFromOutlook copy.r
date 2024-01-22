library(RDCOMClient)

# Create an Outlook application object
Outlook <- COMCreate("Outlook.Application")
myNameSpace <- Outlook$GetNameSpace("MAPI")

# Access the Inbox and then the specific subfolder
inboxFolderIndex <- 1 # Adjust based on your Outlook setup
inbox <- myNameSpace$Folders(inboxFolderIndex)$Folders("Inbox")
bhCatBondFolder <- inbox$Folders("BH Cat Bond")

# Get all messages in the "BH Cat Bond" folder
messages <- bhCatBondFolder$Items()

# Initialize a data frame to store email details
emails_df <- data.frame(Timestamp = character(),
                        Subject = character(),
                        Content = character(),
                        stringsAsFactors = FALSE)

# Get the number of messages in the folder
num_messages <- messages$Count()

# Loop to read messages
for (i in 1:num_messages) {
    message <- messages$Item(i)
    
    # Process only if the message is unread
    if (message$UnRead() == TRUE) {
        # Retrieve email content
        # First try to get the Body (plain text)
        emailContent <- message$Body()

        # If Body is empty or null, fall back to HTMLBody
        if (is.null(emailContent) || emailContent == "") {
            emailContent <- message$HTMLBody()
        }

        # Retrieve and format the ReceivedTime
        receivedTime <- message$ReceivedTime()
        formattedTime <- format(as.POSIXct(receivedTime, origin = "1970-01-01"), "%Y-%m-%d %H:%M:%S")

        # Add email details to the data frame
        emails_df <- rbind(emails_df, data.frame(Timestamp = formattedTime,
                                                 Subject = message$Subject(),
                                                 Content = emailContent))

        # Mark the message as read (optional)
        # message$UnRead(FALSE)
        # message$Save()
    }
}


# Write the data frame to a CSV file
write.csv(emails_df, file = "S:/Touchstone/Catrader/Boston/Database/UnreadDatabaseEntryEmails.csv", row.names = FALSE)
cat("Email processing completed. Data written to UnreadDatabaseEntryEmails.csv\n")
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"><title></title></head><body><!-- rte-version 0.2 9947551637294008b77bce25eb683dac --><div class="rte-style-maintainer rte-pre-wrap" data-color="global-default" bbg-color="default" data-bb-font-size="medium" bbg-font-size="medium" style="font-family: Arial, &quot;BB.Proportional&quot;; white-space: pre-wrap; font-size: small;" style="font-family: Arial, &quot;BB.Proportional&quot;; white-space: pre-wrap; font-size: small; color: rgb(0, 0, 0);"><div class="rte-style-maintainer" data-color="global-default" bbg-color="default" data-bb-font-size="medium" bbg-font-size="medium" style="font-size: small; font-family: &quot;Courier New&quot;, Courier, &quot;BB.FixedWidth&quot;;" style="font-size: small; font-family: &quot;Courier New&quot;, Courier, &quot;BB.FixedWidth&quot;; color: rgb(0, 0, 0);"><div class="rte-style-maintainer rte-pre-wrap" data-color="global-default" bbg-color="default" data-bb-font-size="medium" bbg-font-size="medium" style="font-family: Arial, &quot;BB.Proportional&quot;; white-space: pre-wrap; font-size: small;" style="font-family: Arial, &quot;BB.Proportional&quot;; white-space: pre-wrap; font-size: small; color: rgb(0, 0, 0);"><div class="rte-style-maintainer" data-color="global-default" bbg-color="default" data-bb-font-size="medium" bbg-font-size="medium" style="font-size: small; font-family: &quot;Courier New&quot;, Courier, &quot;BB.FixedWidth&quot;;" style="font-size: small; font-family: &quot;Courier New&quot;, Courier, &quot;BB.FixedWidth&quot;; color: rgb(0, 0, 0);"><div class="rte-style-maintainer rte-pre-wrap" data-color="global-default" bbg-color="default" data-bb-font-size="medium" bbg-font-size="medium" style="font-family: Arial, &quot;BB.Proportional&quot;; white-space: pre-wrap; font-size: small;" style="font-family: Arial, &quot;BB.Proportional&quot;; white-space: pre-wrap; font-size: small; color: rgb(0, 0, 0);"><span class="bbScopedStyle2037222850725069">5mm Nakama 2021-1 1 (62984JAA6) offered @ 99.20</span></div><div class="rte-style-msg-personal-disclaimer" style="border-top:1px dotted #383838; margin-top:12px; padding-top: 12px;"><span class="bbScopedStyle2037222850725069"><span data-bb-font-size="medium" bbg-font-size="medium" style="font-family: Arial, &quot;BB.Proportional&quot;; font-size: small;"><span style="color: rgb(191, 191, 191);">Craig Bonder</span><br style="color: rgb(191, 191, 191);"><span style="color: rgb(191, 191, 191);">BH | Beech Hill Securities, Inc.</span><br style="color: rgb(191, 191, 191);"><span style="color: rgb(191, 191, 191);">Managing Director</span><br style="color: rgb(191, 191, 191);"><span style="color: rgb(191, 191, 191);">880 Third Avenue, 16th floor</span><br style="color: rgb(191, 191, 191);"><span style="color: rgb(191, 191, 191);">New York, NY 10022</span><br style="color: rgb(191, 191, 191);"><span style="color: rgb(191, 191, 191);">Office: 212.257.4475</span><br style="color: rgb(191, 191, 191);"><span style="color: rgb(191, 191, 191);">Cell: 917.930.6363</span><br style="color: rgb(191, 191, 191);"><span style="color: rgb(191, 191, 191);">Email: </span><a spellcheck="false" bbg-destination="mailto:rte:bind" data-destination="mailto:rte:bind" href="mailto:cbonder@bh-secs.com">cbonder@bh-secs.com</a><br style="color: rgb(191, 191, 191);"><br style="color: rgb(191, 191, 191);"><a bbg-destination="rte:bind" spellcheck="false" data-destination="rte:bind" href="http://disclaimerbhs.bh-secs.com/">http://disclaimerbhs.bh-secs.com/</a><br style="color: rgb(191, 191, 191);"><br style="color: rgb(191, 191, 191);"><span style="color: rgb(191, 191, 191);">This email message from Beech Hill Securities, Inc. (“Beech Hill”), including any attachments, (a “Communication”) is for the sole use of its intended recipients and may not be duplicated, re-used, redistributed, or forwarded in whole or in part by any means to any other party. If you are not an intended recipient, please notify the sender, delete it and do not act upon, print, disclose, copy, retain or redistribute it in any manner. Communications are for informational purposes only and not an offer or solicitation to buy or sell any product or service. They may contain privileged or confidential information, or may otherwise be protected by work product immunity or other legal rules, and no such confidentiality or privilege is waived or lost by any error in transmission. No Communication is intended for distribution to, or use by, any person or entity in any location where its distribution or use is contrary to law or regulation, or pursuant to which Beech Hill would be subject to any registration requirement. As no Communication can be guaranteed secure, error-free, uncorrupted, complete or virus free and any Communication may be lost, misdelivered, destroyed, delayed, or intercepted by others, please do not send sensitive or personal data electronically. Beech Hill disclaims all liability in connection with the aforementioned risks associated with electronic communications. All Communications are subject to surveillance, archiving and potential production to regulators and in litigation. These may occur in countries other than the country in which you are located, and may be treated legally differently than in your locale. No Communication is intended to supplant your own evaluation of the matters referenced therein. Prior to any investment decision, investors should obtain sufficient information to ascertain legal, financial, tax and regulatory consequences necessary for such decision. Beech Hill is not a fiduciary or advisor and does not provide advice relating to products or strategies. Past performance is not indicative of future performance. Beech Hill is a registered U.S. broker-dealer, Member FINRA and SIPC.</span></span><br><br></span></div></div></div></div></div></body></html>
