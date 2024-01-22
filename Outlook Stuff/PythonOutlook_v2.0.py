import win32com.client
import pandas as pd
from datetime import datetime

# Create an Outlook application object
Outlook = win32com.client.Dispatch("Outlook.Application")
namespace = Outlook.GetNamespace("MAPI")

# Access the Inbox and then the specific subfolder
inbox = namespace.GetDefaultFolder(6)  # 6 refers to the inbox
bhCatBondFolder = inbox.Folders["BH Cat Bond"]

print("Accessed Outlook folder successfully.")

# List to store email details
emails = []

# Add a test row
emails.append({
    "Timestamp": "Test Time",
    "Subject": "Test Subject",
    "Content": "Test Description"
})

# Loop to process unread messages
for message in bhCatBondFolder.Items:
    if message.UnRead:
        # Extract email details
        subject = message.Subject
        body = message.Body
        received_time = message.ReceivedTime

        # Format the date
        formatted_time = received_time.strftime('%Y-%m-%d %H:%M:%S')

        # Append to the list
        emails.append({
            "Timestamp": formatted_time,
            "Subject": subject,
            "Content": body
        })

        # Mark the message as read (optional)
        # message.UnRead = False
        # message.Save()

print(f"Processed {len(emails) - 1} emails.")

# Create a DataFrame
emails_df = pd.DataFrame(emails)

# Define the output CSV file path
output_csv = "//ad-its.credit-agricole.fr/Amundi_Boston/Homedirs/buonomo/@Config/Desktop/Outlook Scanner/UnreadDatabaseEntryEmails.csv"

# Save the DataFrame to CSV
emails_df.to_csv(output_csv, index=False)

print("Emails saved to CSV file successfully.")
