import os
import win32com.client
from datetime import datetime, timedelta

# Set the target folder where emails will be saved
base_target_folder = "D:\\Daily Health Check"

# Get today's date and the start of the day
today_date = datetime.now().date()
start_of_day = datetime.combine(today_date, datetime.min.time())

# Create the new folder if it doesn't exist
target_folder = os.path.join(base_target_folder, today_date.strftime("%Y-%m-%d"))
os.makedirs(target_folder, exist_ok=True)

# Initialize Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Get the desired Outlook folder (e.g., Inbox) and subfolder
inbox = outlook.GetDefaultFolder(6)  # 6 corresponds to the Inbox folder
subfolder_name = "Health_Check"  # Replace with the name of your subfolder

# Locate the subfolder within the Inbox
subfolder = None
for folder in inbox.Folders:
    if folder.Name == subfolder_name:
        subfolder = folder
        break

if subfolder:
    # Get all emails in the subfolder received today
    emails = subfolder.Items.Restrict("[ReceivedTime] >= '" + start_of_day.strftime('%m/%d/%Y %H:%M %p') + "'")

    # Iterate through the emails
    for email in emails:
        # Check if the email has any attachments
        if email.Attachments.Count > 0:
            has_excel_attachment = False

            # Check if any attachment is an Excel file (.xlsx or .xlsm or xlsb or csv)
            for attachment in email.Attachments:
                if attachment.FileName.lower().endswith(('.xlsx', '.xlsm', '.xlsb', '.csv')):
                    has_excel_attachment = True
                    break

            if has_excel_attachment:
                # Save the entire email (including content) as a .msg file in the target folder
                email.SaveAs(os.path.join(target_folder, f"{email.Subject}.msg"))
                print(f"Email '{email.Subject}' with Excel attachment saved.")
else:
    print(f"Subfolder '{subfolder_name}' not found.")

print("Emails with Excel attachments received today from subfolder saved successfully!")


