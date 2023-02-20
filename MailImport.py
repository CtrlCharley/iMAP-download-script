import imaplib
import email
import os
import datetime

# IMAP server login details
IMAP_SERVER = 'server'
USERNAME = 'login'
PASSWORD = 'password'

# Mailbox and search criteria
MAILBOX = 'INBOX'
SEARCH_CRITERIA = 'ALL'

# Connect to the IMAP server and select the mailbox
imap = imaplib.IMAP4_SSL(IMAP_SERVER)
imap.login(USERNAME, PASSWORD)
imap.select(MAILBOX)

try:
    # Search for the last email
    status, messages = imap.search(None, SEARCH_CRITERIA)
    last_message = messages[0].split()[-1]

    # Fetch the last email message
    status, data = imap.fetch(last_message, '(RFC822)')
    email_message = email.message_from_bytes(data[0][1])

    # Loop through the email parts
    for part in email_message.walk():
        # Check if the part is an XLSX file attachment
        if part.get_content_type() == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
            # Save the attachment to the current directory with the name "any name"
            today_date = datetime.date.today().strftime('%Y%m%d')
            filename = f'MY NAMMED FILE _ CHANGE IT)'
            with open(filename, 'wb') as f:
                f.write(part.get_payload(decode=True))
            print(f'Saved attachment {filename} to current directory')

            # Log the activity with a timestamp
            with open('log.txt', 'a') as log_file:
                timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                log_file.write(f'{timestamp}: Saved attachment {filename} to current directory\n')

    # Delete all emails from the mailbox
    imap.store('1:*', '+FLAGS', '\\Deleted')
    imap.expunge()
    print('Inbox cleared')

    # Log the activity with a timestamp
    with open('log.txt', 'a') as log_file:
        timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        log_file.write(f'{timestamp}: Inbox cleared\n')

except Exception as e:
    # Log the error with a timestamp
    with open('log.txt', 'a') as log_file:
        timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        log_file.write(f'{timestamp}: Error occurred: {e}\n')

finally:
    # Close the connection to the IMAP server
    imap.close()
    imap.logout()
