import win32com.client
import pandas as pd
from datetime import datetime

# Function to extract content up to a specific marker
def extract_content_up_to_marker(body, marker):
    marker_position = body.find(marker)
    if marker_position != -1:
        return body[:marker_position]
    else:
        return body



# Create an Outlook application object
Outlook = win32com.client.Dispatch("Outlook.Application")
namespace = Outlook.GetNamespace("MAPI")

# Access the Inbox and then the specific subfolder
inbox = namespace.GetDefaultFolder(6)  # 6 refers to the inbox
bhCatBondFolder = inbox.Folders["BH Cat Bond"]

emails = []

# Loop to process unread messages
for message in bhCatBondFolder.Items:
    if message.UnRead:
        subject = message.Subject
        body = message.Body
        received_time = message.ReceivedTime

        # Format the date
        formatted_time = received_time.strftime('%Y-%m-%d %H:%M:%S')

        # Extract content up to the marker
        extracted_body = extract_content_up_to_marker(body, "Craig Bonder")

        emails.append({
            "Timestamp": formatted_time,
            "Subject": subject,
            "Content": extracted_body
        })

        # Mark the message as read (optional)
        # message.UnRead = False
        # message.Save()

# Create a DataFrame
emails_df = pd.DataFrame(emails)

# Define the output CSV file path
output_csv = "//ad-its.credit-agricole.fr/Amundi_Boston/Homedirs/buonomo/@Config/Desktop/Outlook Scanner/UnreadDatabaseEntryEmails.csv"

# Save the DataFrame to CSV
emails_df.to_csv(output_csv, index=False)

print("Emails processed and saved to CSV file.")
