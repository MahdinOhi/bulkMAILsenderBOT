Email Sender BOT

This Python script allows you to send emails with attachments to multiple recipients using Gmail's SMTP server. Recipient email addresses are read from an Excel file, and attachments are attached to emails based on the order of recipients and files in a specified directory.

Requirements
Python 3.x
openpyxl library (install using pip install openpyxl)
Usage
Clone this repository or download the email_sender_with_attachment.py file.
Install the required dependencies by running pip install openpyxl.
Update the hardcoded sender's email and password in the script.
Prepare an Excel file with recipient email addresses in the first column.
Prepare the directory containing the attachments to be sent.
Run the script and follow the prompts:
Enter the path of the Excel file containing recipient email addresses.
Enter the path of the directory containing attachments.
Enter the subject of the email.
The script will send emails to the recipients listed in the Excel file, attaching corresponding files from the specified directory.
Notes
Make sure to enable "Less Secure Apps" in your Gmail settings or use an App Password if you have 2-Step Verification enabled for your Google account.
Ensure that the Excel file format is .xlsx, and recipient email addresses are listed in the first column.
Attachments are sent sequentially to each recipient based on their order in the Excel file and the order of files in the specified directory.