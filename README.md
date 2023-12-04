# Attachment Grabber

This script is designed to connect to specified IMAP servers, search for emails, and download attachments from these emails.

## Prerequisites

- Python 3.x installed
- Required Python modules: `imaplib`, `email`, `os`, `logging`, `argparse`, `glob`, `tqdm`, `concurrent.futures`
- Optional for error reporting: `sentry_sdk`

## Configuration

Before running the script, you need to set up the following configurations:

1. **servers.env file**: Create a file named `servers.env` in the same directory as the script with the IMAP server details and credentials in the following format:

    ```
    SERVERNAME_imap_url=imap.example.com
    SERVERNAME_username=user@example.com
    SERVERNAME_password=yourpassword
    SERVERNAME_port=993
    SERVERNAME_use_ssl=True
    ```

   Replace `SERVERNAME` with a unique identifier for each server.

2. **Environment Variables**: If you want to use Sentry for error reporting, you can set the following environment variables:

    - `_LOG_TO_SENTRY`: Set to `true` to enable logging to Sentry.
    - `_SENTRY_DSN`: Your Sentry DSN for logging.
    - `_SENTRY_LOG_LEVEL`: Sentry log level (e.g., `ERROR`, `INFO`).

## Usage

Run the script using the following command:

`python imap_attachment_downloader.py [hostnames...]`


- If no hostnames are provided, the script will attempt to download attachments from all configured servers in the `servers.env` file.
- If hostnames are provided, the script will only download attachments from those specified servers.

## Features

- The script supports downloading attachments with extensions: `.pdf`, `.doc`, `.docx`, `.xls`, `.xlsx`. To change the allowed file types, modify the `allowed_extensions` tuple in the script.
- The script organizes downloaded attachments into subdirectories by server and file extension.
- Logging of activities is done to `imap_download.log`, and failures to decode emails are logged to `decode_failures.log`.
- For Sentry logging, the script requires a Sentry DSN.

## Note

- It is recommended to use a secure and encrypted method to store the server credentials.
- Ensure that the `servers.env` file is protected and not accessible by unauthorized users.

## License

This script is provided "as is", without warranty of any kind, express or implied. Use at your own risk. Blah blah blah.
