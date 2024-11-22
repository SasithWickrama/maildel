import win32com.client
import datetime

# Create an instance of the Outlook application
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Access the inbox folder
inbox = outlook.GetDefaultFolder(6)  # 6 corresponds to the inbox

# Get all the messages in the inbox
messages = inbox.Items

# Define the criteria for deletion
sender_email = "sasith.wickrama@gmail.com"  # Replace with the sender's email address
subject = "test"  # Replace with the subject of the email
sent_date = datetime.datetime(2024, 8, 1)  # Replace with the date the email was sent

# Convert sent_date to Outlook's date format
outlook_date_format = sent_date.strftime("%m/%d/%Y %H:%M %p")

# Loop through the messages and delete those that match the criteria
for message in messages:
    if message.SenderEmailAddress == sender_email and message.Subject == subject and message.SentOn.strftime("%m/%d/%Y %H:%M %p") == outlook_date_format:
        message.Delete()
        print(f"Deleted: {message.Subject} from {message.SenderEmailAddress} sent on {message.SentOn}")

print("Finished deleting emails.")
