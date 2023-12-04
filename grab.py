import imaplib
import email
import os
import logging
from email.header import decode_header
from email.utils import parsedate_to_datetime
import datetime
import argparse
import glob
from tqdm import tqdm
import concurrent.futures
import sentry_sdk
from sentry_sdk.integrations.logging import LoggingIntegration

# Increase the buffer size to 100,000,000 bytes (approximately 100 MB)
imaplib._MAXLINE = 100000000

# Set up main logging
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)

# Create file handler which logs even debug messages
fh = logging.FileHandler('imap_download.log')
fh.setLevel(logging.DEBUG)

# Create console handler with a higher log level
ch = logging.StreamHandler()
ch.setLevel(logging.ERROR)

# Create formatter and add it to the handlers
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
ch.setFormatter(formatter)
fh.setFormatter(formatter)

sentry_logging = LoggingIntegration(
    level=logging.INFO,
    event_level=logging.ERROR
)

# Set up decoding failure logging
fail_logger = logging.getLogger('decode_fail')
fail_logger.setLevel(logging.DEBUG)

# Create file handler for decoding failure log
fail_fh = logging.FileHandler('decode_failures.log')
fail_fh.setLevel(logging.DEBUG)

fail_fh.setFormatter(formatter)

# Load environment variables from 'servers.env' if it exists
env_file = 'servers.env'
servers = {}
variables = {}  # New dictionary for variables
if os.path.exists(env_file):
    with open(env_file) as f:
        for line in f:
            if line.startswith('#') or not line.strip():
                continue
            key, value = line.strip().split('=', 1)
            if key.startswith('_'):  # Check if the line is a variable
                variables[key] = value
            else:
                server_name = key.split('_')[0]
                if server_name not in servers:
                    servers[server_name] = {}
                servers[server_name][key] = value

# Fetch the variables from the new 'variables' dictionary
log_to_sentry = variables.get('_LOG_TO_SENTRY', 'false').lower() == 'true'  
sentry_dsn = variables.get('_SENTRY_DSN', '') 

# Fetch the _SENTRY_LOG_LEVEL variable
sentry_log_level = variables.get('_SENTRY_LOG_LEVEL', 'ERROR')

# Convert the log level to the corresponding logging constant
sentry_log_level = getattr(logging, sentry_log_level)

sentry_logging = LoggingIntegration(
    level=sentry_log_level,  # Use the fetched log level
    event_level=logging.ERROR
)

if log_to_sentry:
    sentry_sdk.init(
        dsn=sentry_dsn,
        integrations=[sentry_logging]
    )
    logger.addHandler(sentry_sdk.integrations.logging.EventHandler())
    fail_logger.addHandler(sentry_sdk.integrations.logging.EventHandler())
else:
    # Add the handlers to the logger
    logger.addHandler(ch)
    logger.addHandler(fh)
    # Add the handler to the fail_logger
    fail_logger.addHandler(fail_fh)

# Allowed file extensions
allowed_extensions = ('.pdf', '.doc', '.docx', '.xls', '.xlsx')

# Parse command-line arguments
parser = argparse.ArgumentParser(description='Download attachments from an IMAP server.')
parser.add_argument('host', nargs='*', help='the hostnames of the IMAP servers')
args = parser.parse_args()

# If no servers are specified, download from all servers
if not args.host:
    args.host = list(servers.keys())

def download_attachments(server_name):
    # IMAP server details
    imap_url = servers[server_name].get(f'{server_name}_imap_url', 'default_imap_url')
    username = servers[server_name].get(f'{server_name}_username', 'default_username')
    password = servers[server_name].get(f'{server_name}_password', 'default_password')
    port = int(servers[server_name].get(f'{server_name}_port', 993))  # Default IMAP over SSL port
    use_ssl = bool(servers[server_name].get(f'{server_name}_use_ssl', True))  # Change to False if not using SSL

    try:
        logger.info(f"Starting to connect to the server {server_name}.")
        # connect to the server
        if use_ssl:
            mail = imaplib.IMAP4_SSL(imap_url, port=port)
        else:
            mail = imaplib.IMAP4(imap_url, port=port)
        logger.info(f"Connected to the server {server_name}.")

        logger.info(f"Starting to authenticate for the server {server_name}.")
        # authenticate
        mail.login(username, password)
        logger.info(f"Authenticated successfully for the server {server_name}.")

        # select the mailbox you want to delete in
        # if you want SPAM, use "INBOX.SPAM"
        mailbox = "INBOX"
        logger.info(f"Starting to select the mailbox {mailbox} for the server {server_name}.")
        mail.select(mailbox)
        logger.info(f"Mailbox {mailbox} selected for the server {server_name}.")

        # Get the list of all files in the directory
        files = glob.glob(os.path.join(os.getcwd(), server_name, '*/*'))

        # Extract the dates from the filenames and convert to datetime.date objects
        dates = [datetime.datetime.strptime(os.path.basename(f).split('_')[0], '%Y-%m-%d').date() for f in files]

        # If there are no files, start_date is None (start from the beginning)
        # Otherwise, start from the day after the latest date in the files
        start_date = max(dates) + datetime.timedelta(days=1) if dates else None

        # get uids
        if start_date is not None:
            # Format the start_date in the format required by the IMAP protocol
            start_date_str = start_date.strftime("%d-%b-%Y")
            result, data = mail.uid('search', None, f'(SINCE "{start_date_str}")')
        else:
            result, data = mail.uid('search', None, "ALL")
        uid_list = data[0].split()

        # Calculate the total number of emails
        total_emails = len(uid_list)

        # Create a progress bar with the server name as the prefix
        pbar = tqdm(total=total_emails, desc=server_name)

        # loop through all uids
        for uid in uid_list:
            # fetch the email (header and body)
            result, message_data = mail.uid('fetch', uid, '(BODY.PEEK[])')
            raw_email = message_data[0][1]
            if isinstance(raw_email, bytes):
                try:
                    raw_email = raw_email.decode("utf-8")
                except UnicodeDecodeError:
                    try:
                        raw_email = raw_email.decode("iso-8859-1")
                    except UnicodeDecodeError:
                        email_message = email.message_from_bytes(raw_email)
                        decoded_header = decode_header(email_message["Subject"])[0][0]
                        if isinstance(decoded_header, bytes):
                            email_subject = decoded_header.decode()
                        else:
                            email_subject = decoded_header

                        email_date = email_message["Date"]
                        fail_logger.error(f'Failed to decode email with uid {uid}, subject "{email_subject}", and date "{email_date}"')
                        pbar.update(1)
                        continue
            email_message = email.message_from_string(raw_email)

            # get the delivery date from the email header and format it in ISO format
            delivery_date = parsedate_to_datetime(email_message['Date']).date()

            # downloading attachments
            for part in email_message.walk():
                if part.get_content_maintype() == 'multipart':
                    continue
                if part.get('Content-Disposition') is None:
                    continue
                fileName = part.get_filename()

                if bool(fileName):
                    _, extension = os.path.splitext(fileName)
                    if extension.lower() in allowed_extensions:
                        fileName = delivery_date.isoformat() + '_' + fileName
                        dirPath = os.path.join(os.getcwd(), server_name, extension.lstrip('.'))
                        filePath = os.path.join(dirPath, fileName)

                        # Check if there's already a file with a newer or same date
                        if os.path.exists(dirPath):
                            for file in os.listdir(dirPath):
                                existing_file_date = datetime.datetime.strptime(file.split('_')[0], '%Y-%m-%d').date()
                                if existing_file_date >= delivery_date:
                                    continue

                        os.makedirs(os.path.dirname(filePath), exist_ok=True)
                        if not os.path.isfile(filePath):
                            try:
                                with open(filePath, 'wb') as f:
                                    f.write(part.get_payload(decode=True))
                                logger.info(f'Successfully wrote file {filePath}')
                            except Exception as e:
                                logger.error(f'Failed to write file {filePath}: {str(e)}')

            # Update the progress bar
            pbar.update(1)

        # Close the progress bar
        pbar.close()

        print(f"All attachments have been downloaded for the server {server_name}.")

    except Exception as e:
        logger.error(f'An error occurred for the server {server_name}: {str(e)}')

try:
    # Run the download_attachments function concurrently for each server
    with concurrent.futures.ThreadPoolExecutor() as executor:
        futures = {executor.submit(download_attachments, host): host for host in args.host}
        for future in concurrent.futures.as_completed(futures):
            host = futures[future]
            try:
                future.result()
            except Exception as exc:
                logger.error('%r generated an exception: %s' % (host, exc))

except KeyboardInterrupt:
    logger.info("Received keyboard interrupt, shutting down after the current download(s) finish...")
    executor.shutdown(wait=True)
    logger.info("Shutdown complete.")
