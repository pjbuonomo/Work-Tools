library(RDCOMClient)
library(rvest)

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
# Loop to read messages
# Loop to read messages
for (i in 1:num_messages) {
    message <- messages$Item(i)
    
    # Process only if the message is unread
    if (message$UnRead() == TRUE) {
        # Retrieve HTML email content
        htmlContent <- message$HTMLBody()

        # Parse and extract content within the <html> tags
        if (!is.null(htmlContent) && htmlContent != "") {
            parsedHtml <- read_html(htmlContent)
            extractedContent <- html_text(parsedHtml)
        } else {
            # Use a placeholder if HTML content is not available
            extractedContent <- "No HTML content available"
        }

        # Retrieve and format the ReceivedTime
        receivedTime <- message$ReceivedTime()
        formattedTime <- format(as.POSIXct(receivedTime, origin = "1970-01-01"), "%Y-%m-%d %H:%M:%S")

        # Add email details to the data frame
        emails_df <- rbind(emails_df, data.frame(Timestamp = formattedTime,
                                                 Subject = message$Subject(),
                                                 Content = extractedContent))

        # Mark the message as read (optional)
        # message$UnRead(FALSE)
        # message$Save()
    }
}


# Write the data frame to a CSV file
write.csv(emails_df, file = "S:/Touchstone/Catrader/Boston/Database/UnreadDatabaseEntryEmails.csv", row.names = FALSE)



Working
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"><title></title></head><body><!-- rte-version 0.2 9947551637294008b77bce25eb683dac --><div class="rte-style-maintainer rte-pre-wrap" data-color="global-default" bbg-color="default" data-bb-font-size="medium" bbg-font-size="medium" style="font-family: Arial, &quot;BB.Proportional&quot;; white-space: pre-wrap; font-size: small;" style="font-family: Arial, &quot;BB.Proportional&quot;; white-space: pre-wrap; font-size: small; color: rgb(0, 0, 0);"><div class="rte-style-maintainer" data-color="global-default" bbg-color="default" data-bb-font-size="medium" bbg-font-size="medium" style="font-size: small; font-family: &quot;Courier New&quot;, Courier, &quot;BB.FixedWidth&quot;;" style="font-size: small; font-family: &quot;Courier New&quot;, Courier, &quot;BB.FixedWidth&quot;; color: rgb(0, 0, 0);"><div class="rte-style-maintainer rte-pre-wrap" data-color="global-default" bbg-color="default" data-bb-font-size="medium" bbg-font-size="medium" style="font-family: Arial, &quot;BB.Proportional&quot;; white-space: pre-wrap; font-size: small;" style="font-family: Arial, &quot;BB.Proportional&quot;; white-space: pre-wrap; font-size: small; color: rgb(0, 0, 0);"><div class="rte-style-maintainer" data-color="global-default" bbg-color="default" data-bb-font-size="medium" bbg-font-size="medium" style="font-size: small; font-family: &quot;Courier New&quot;, Courier, &quot;BB.FixedWidth&quot;;" style="font-size: small; font-family: &quot;Courier New&quot;, Courier, &quot;BB.FixedWidth&quot;; color: rgb(0, 0, 0);"><div class="rte-style-maintainer rte-pre-wrap" data-color="global-default" bbg-color="default" data-bb-font-size="medium" bbg-font-size="medium" style="font-family: Arial, &quot;BB.Proportional&quot;; white-space: pre-wrap; font-size: small;" style="font-family: Arial, &quot;BB.Proportional&quot;; white-space: pre-wrap; font-size: small; color: rgb(0, 0, 0);"><span class="bbScopedStyle8017738174112958">Please show in all offerings. Many thanks for the focus.</span></div><div class="rte-style-maintainer rte-pre-wrap" data-color="global-default" bbg-color="default" data-bb-font-size="medium" bbg-font-size="medium" style="font-family: Arial, &quot;BB.Proportional&quot;; white-space: pre-wrap; font-size: small;" style="font-family: Arial, &quot;BB.Proportional&quot;; white-space: pre-wrap; font-size: small; color: rgb(0, 0, 0);"><div><span class="bbScopedStyle8017738174112958">Alamo 2023-1 A (011395AJ9) bid at 102.50</span></div><div><span class="bbScopedStyle8017738174112958">Blue Sky 2023-1 (XS2728630596) bid at 100.15</span></div><div><span class="bbScopedStyle8017738174112958">Bonanza 2022-1 A (09785EAJ0) bid at 90.00</span></div><div><span class="bbScopedStyle8017738174112958">Bonanza 2023-1 A (09785EAK7) bid at 99.90</span></div><div><span class="bbScopedStyle8017738174112958">Citrus 2023-1 B (177510AM6) bid at 102.40</span></div><div><span class="bbScopedStyle8017738174112958">Easton 2024-1 A (27777AAA9) bid at 100.25<br>First Coast 2021-1 (31971CAA1) bid at 96.15</span></div><div><span class="bbScopedStyle8017738174112958">First Coast 2023-1 (31969UAA5) bid at 101.10</span></div><div><span class="bbScopedStyle8017738174112958">Galileo 2023-1 B (36354TAP7) bid at 100.25</span></div><div><span class="bbScopedStyle8017738174112958">Galileo 2023-1 A (36354TAN2) bid at 100.25</span></div><div><span class="bbScopedStyle8017738174112958">Hypatia 2023-1 A (44914CAC0) bid at 104.35</span></div><div><span class="bbScopedStyle8017738174112958">Hexagon 2023-1 A (428270AA0) bid at 100.50</span></div><div><span class="bbScopedStyle8017738174112958">Lightning 2023-1 A (532242AA2) bid at 106.30</span></div><div><span class="bbScopedStyle8017738174112958">Matterhorn 2022-I B (577092AQ2) bid at 98.50</span></div><div><span class="bbScopedStyle8017738174112958">Merna 2022-2A (59013MAF9) bid at 98.65 </span></div><div><span class="bbScopedStyle8017738174112958">Merna 2023-2 A (59013MAJ1) bid at 104.35</span></div><div><span class="bbScopedStyle8017738174112958">Mona Lisa 2023-1 B (608800AG3) bid at 107.75</span></div><div><span class="bbScopedStyle8017738174112958">Montoya 2022-2 (613752AB0) bid at 108.60</span></div><div><span class="bbScopedStyle8017738174112958">Montoya 2024-1 A (613752AC8) bid at 101.00</span></div><div><span class="bbScopedStyle8017738174112958">Ocelot 2023-1 A (675951AA5) bid at 100.30</span></div><div><span class="bbScopedStyle8017738174112958">Residential Re 2023-2 5 (76090WAC4) bid at 100.45</span></div><div><span class="bbScopedStyle8017738174112958">Tailwind 2022-1 B (87403TAE6) bid at 96.90</span></div><div><span class="bbScopedStyle8017738174112958">Tailwind 2022-1 C (87403TAE) bid at 98.15</span></div><div><span class="bbScopedStyle8017738174112958">Titania 2021-1 A (888329AA7) bid at 100.40</span></div><div><span class="bbScopedStyle8017738174112958">Titania 2021-2 A (888329AB5) bid at 97.50</span></div><div><span class="bbScopedStyle8017738174112958">Titania 2023-1 A (888329AC3) bid at 108.50</span></div><div><span class="bbScopedStyle8017738174112958">Ursa 2023-1 C (90323WAM2) bid at 100.45</span></div><div><span class="bbScopedStyle8017738174112958">Ursa 2023-3 D (90323WAQ3) bid at 100.35</span></div></div></div></div></div></div></body></html>

Not Working
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"><title></title></head><body><!-- rte-version 0.2 9947551637294008b77bce25eb683dac --><div class="rte-style-maintainer rte-pre-wrap" data-color="global-default" bbg-color="default" data-bb-font-size="medium" bbg-font-size="medium" style="font-family: Arial, &quot;BB.Proportional&quot;; white-space: pre-wrap; font-size: small;" style="font-family: Arial, &quot;BB.Proportional&quot;; white-space: pre-wrap; font-size: small; color: rgb(0, 0, 0);"><div>2.25mm Gateway 2023-3 A (36779CAF3) offered @ 107.10 </div><div><div>3mm Tailwind 2022-1 C (87403TAF3) offered @ 100.10</div><div><br></div></div></div><div class="rte-style-msg-personal-disclaimer" style="border-top:1px dotted #383838; margin-top:12px; padding-top: 12px;"><span data-bb-font-size="medium" style="font-family: Arial, 'BB.Proportional'; font-size: small; color:#383838;">Craig Bonder<br>BH | Beech Hill Securities, Inc.<br>Managing Director<br>880 Third Avenue, 16th floor<br>New York, NY 10022<br>Office: 212.257.4475<br>Cell: 917.930.6363<br>Email: <a spellcheck="false" href="mailto:cbonder@bh-secs.com">cbonder@bh-secs.com</a><br><br><a bbg-destination="rte:bind" spellcheck="false" href="http://disclaimerbhs.bh-secs.com/">http://disclaimerbhs.bh-secs.com/</a><br><br>This email message from Beech Hill Securities, Inc. (“Beech Hill”), including any attachments, (a “Communication”) is for the sole use of its intended recipients and may not be duplicated, re-used, redistributed, or forwarded in whole or in part by any means to any other party. If you are not an intended recipient, please notify the sender, delete it and do not act upon, print, disclose, copy, retain or redistribute it in any manner. Communications are for informational purposes only and not an offer or solicitation to buy or sell any product or service. They may contain privileged or confidential information, or may otherwise be protected by work product immunity or other legal rules, and no such confidentiality or privilege is waived or lost by any error in transmission. No Communication is intended for distribution to, or use by, any person or entity in any location where its distribution or use is contrary to law or regulation, or pursuant to which Beech Hill would be subject to any registration requirement. As no Communication can be guaranteed secure, error-free, uncorrupted, complete or virus free and any Communication may be lost, misdelivered, destroyed, delayed, or intercepted by others, please do not send sensitive or personal data electronically. Beech Hill disclaims all liability in connection with the aforementioned risks associated with electronic communications. All Communications are subject to surveillance, archiving and potential production to regulators and in litigation. These may occur in countries other than the country in which you are located, and may be treated legally differently than in your locale. No Communication is intended to supplant your own evaluation of the matters referenced therein. Prior to any investment decision, investors should obtain sufficient information to ascertain legal, financial, tax and regulatory consequences necessary for such decision. Beech Hill is not a fiduciary or advisor and does not provide advice relating to products or strategies. Past performance is not indicative of future performance. Beech Hill is a registered U.S. broker-dealer, Member FINRA and SIPC.</span><br><br></div></body></html>
