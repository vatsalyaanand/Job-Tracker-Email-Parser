import imaplib
import email
from email.header import decode_header
from openpyxl import Workbook
from datetime import datetime

# Function to extract job details from email body
def extract_job_details(body):
    lines = body.split('\n')
    job_role = lines[2].strip()
    company_name = lines[3].strip()
    location = lines[4].strip()
    return job_role, company_name, location


# IMAP settings
IMAP_SERVER = 'imap.gmail.com'
EMAIL_ADDRESS = 'xx@gmail.com'
EMAIL_PASSWORD = 'xx'

# Connect to the IMAP server
imap = imaplib.IMAP4_SSL(IMAP_SERVER)

# Log in to your Gmail account
imap.login(EMAIL_ADDRESS, EMAIL_PASSWORD)

# Select the mailbox (inbox)
status, messages = imap.select('INBOX')

# Search for emails from LinkedIn with subject "your application was sent"
status, messages = imap.search(None, '(FROM "jobs-noreply@linkedin.com" SUBJECT "your application was sent to")')

# Fetch the most recent 10 emails
latest_10_messages = messages[0].split()[-20:]
# Create a new Excel workbook
wb = Workbook()
ws = wb.active

# Add headers to the Excel sheet
ws.append(["Company", "Position", "Location","Date"])

# Iterate through each email
for msg_id in latest_10_messages:
    # Fetch the email
    status, msg_data = imap.fetch(msg_id, '(RFC822)')

    # Parse the email
    email_message = email.message_from_bytes(msg_data[0][1])

    # Extract email details
    received_date_str = email_message.get("Date")
    received_date_str = received_date_str.split('(')[0].strip()
    received_date = datetime.strptime(received_date_str, "%a, %d %b %Y %H:%M:%S %z")
    received_date_formatted = received_date.strftime("%m-%d-%Y")

    body = ""

    # If the email is multipart
    if email_message.is_multipart():
        # Iterate through email parts
        for part in email_message.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition"))

            # Ignore any text/plain (plaintext) attachments
            if "attachment" not in content_disposition:
                # Decode email part to Unicode if it's in bytes
                payload = part.get_payload(decode=True)
                if payload is not None:
                    charset = part.get_content_charset()
                    if charset:
                        body += payload.decode(charset, errors="ignore")
                    else:
                        body += payload
    else:
        # Email is not multipart
        body = email_message.get_payload(decode=True).decode()

    # Extract job details from email body
    job_role, company_name, location = extract_job_details(body)
    ws.append([company_name, job_role, location, received_date_formatted])
# Save the Excel workbook
wb.save("job_tracker.xlsx")

# Close the connection
imap.close()
imap.logout()