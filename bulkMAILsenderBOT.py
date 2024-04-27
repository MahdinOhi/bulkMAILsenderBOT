import smtplib
import openpyxl
import os
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

sender_email = ""
sender_password = ""
message = "This is a test email message."


def read_recipients_from_excel(file_path):
    recipients = []
    try:
        wb = openpyxl.load_workbook(file_path)
        sheet = wb.active
        for row in sheet.iter_rows(values_only=True):
            email = row[0]  # Assuming email addresses are in the first column
            recipients.append(email.strip())
    except Exception as e:
        print(f"Error reading recipients from Excel file: {e}")
    return recipients


def get_attachments_in_order(attachment_dir):
    attachments = []
    if os.path.isdir(attachment_dir):
        attachments = sorted([os.path.join(attachment_dir, f) for f in os.listdir(
            attachment_dir) if os.path.isfile(os.path.join(attachment_dir, f))])
    return attachments


def get_email_subject_from_user():
    return input("Enter the subject of the email: ").strip()


def Email_send_function(recipients, subject, sender_email, sender_password, attachments=None):
    s = smtplib.SMTP("smtp.gmail.com", 587)
    s.starttls()
    s.login(sender_email, sender_password)

    if attachments:
        if len(recipients) != len(attachments):
            print("Number of recipients does not match number of attachments.")
            return

        for recipient, attachment in zip(recipients, attachments):
            if os.path.isfile(attachment):  # If file exists
                with open(attachment, "rb") as f:
                    attachment_data = f.read()
                filename = os.path.basename(attachment)
                msg = create_message_with_attachment(
                    subject, sender_email, recipient, message, attachment_data, filename)
                s.sendmail(sender_email, recipient, msg.as_string())
            else:
                print(f"Attachment not found: {attachment}")

    else:
        for recipient in recipients:
            msg = "Subject: {}\n\n{}".format(subject, message)
            s.sendmail(sender_email, recipient, msg)

    s.close()


def create_message_with_attachment(subject, sender_email, recipient, message, attachment_data, filename):
    msg = MIMEMultipart()
    msg["From"] = sender_email
    msg["To"] = recipient
    msg["Subject"] = subject

    part = MIMEBase("application", "octet-stream")
    part.set_payload(attachment_data)
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f"attachment; filename= {filename}")
    msg.attach(part)

    msg.attach(MIMEText(message, "plain"))

    return msg


def main():
    recipients_file_path = input(
        "Enter the path of the Excel file containing recipient email addresses: ").strip()
    attachment_dir = input(
        "Enter the path of the directory containing attachments: ").strip()
    subject = get_email_subject_from_user()

    recipients = read_recipients_from_excel(recipients_file_path)
    attachments = get_attachments_in_order(attachment_dir)

    if recipients:
        Email_send_function(recipients, subject, sender_email,
                            sender_password, attachments)
    else:
        print("No recipients found. Exiting...")


if __name__ == "__main__":
    main()
