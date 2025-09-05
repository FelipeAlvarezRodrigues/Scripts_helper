"""
🔧 Configuration:
- Update `downloads_folder` with the path where your report files are stored.
- Customize the `recipients` list with placeholder email addresses, messages, and filenames.
- Ensure Microsoft Outlook is installed and configured on your system.

💡 Requirements:
- Python 3.x
- pywin32 (install via `pip install pywin32`)

⚠️ Note:
This version uses placeholder data and is safe to share publicly.
Replace placeholders with actual values when deploying in a real environment.
"""  

import os
import win32com.client

# Configuration: Adjust these values as needed
downloads_folder = r"C:\Path\To\Your\Downloads"

# List of recipients, each with email, message, and specific file to send
recipients = [
    {
        "email": "recipient1@example.com",
        "message": "Hello,\n\nPlease find attached the weekly report.\n\nBest regards,\nYour Name",
        "file": "report1.txt",
    },
    {
        "email": "recipient2@example.com",
        "message": "Hello,\n\nPlease find attached the weekly report.\n\nBest regards,\nYour Name",
        "file": "report2.txt",
    },
    {
        "email": "recipient3@example.com",
        "message": "Hello,\n\nPlease find attached the weekly report.\n\nBest regards,\nYour Name",
        "file": "report3.txt",
    },
]
subject = "Weekly Report: Blocked Emails"

# Function to send an email with attachment using Outlook
def send_email(recipient_email, subject, body, attachment_path):
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = recipient_email
    mail.Subject = subject
    mail.Body = body
    mail.Attachments.Add(attachment_path)
    mail.Send()

# Main function
def main():
    for recipient in recipients:
        file_path = os.path.join(downloads_folder, recipient["file"])
        if not os.path.exists(file_path):
            print(f"File not found: {file_path}")
            continue
        send_email(recipient["email"], subject, recipient["message"], file_path)
        print(f"Email sent to {recipient['email']} with attachment {recipient['file']}.")

# Run the script
if __name__ == "__main__":
    main()
