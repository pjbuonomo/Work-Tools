library(OutlookR)

# Connect to Outlook
outlook <- ol_connect()

# Find the "BH Cat Bond" folder
folders <- ol_get_folders(outlook)
inbox_folder <- ol_get_folder(folders, "BH Cat Bond")

# Retrieve unread emails
unread_emails <- ol_get_items(inbox_folder, filter = "IsUnread = TRUE")

# Loop through unread emails and extract content
for (email in unread_emails) {
  email_subject <- email$GetSubject()
  email_body <- email$GetBodyHTML()
  
  # Do something with the email subject and body, e.g., print or save to a file
  print(paste("Subject:", email_subject))
  print(paste("Body:", email_body))
  
  # Mark the email as read (optional)
  email$MarkAsRead()
}

# Disconnect from Outlook
ol_disconnect(outlook)
repos = "https://cran.r-project.org/bin/windows/contrib/your_R_version/"
devtools::install_github("jameshuynh/OutlookR")

