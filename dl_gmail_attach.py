from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.http import MediaFileUpload

import hashlib
import os
import base64
import email
import re
import requests
import mimetypes
import datetime
import pdfplumber

SCOPES = [
    'https://www.googleapis.com/auth/gmail.modify',
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
]

SPREADSHEET_NAME = "ClearPayDB"
SHEET_NAME = "Files"
INV_SHEET_NAME = "Invoices"

# Function to authenticate and build Gmail, Sheets, and Drive services
def authenticate_services():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    
    gmail_service = build('gmail', 'v1', credentials=creds)
    sheets_service = build('sheets', 'v4', credentials=creds)
    drive_service = build('drive', 'v3', credentials=creds)
    return gmail_service, sheets_service, drive_service

# Function to find or create the spreadsheet by name
def get_or_create_spreadsheet(sheets_service, drive_service):
    # Search for an existing spreadsheet by name using Drive API
    query = f"name = '{SPREADSHEET_NAME}' and mimeType = 'application/vnd.google-apps.spreadsheet'"
    results = drive_service.files().list(q=query, fields="files(id, name)").execute()
    files = results.get('files', [])

    if files:
        spreadsheet_id = files[0]['id']
        print(f"Spreadsheet '{SPREADSHEET_NAME}' found with ID: {spreadsheet_id}")
        return spreadsheet_id
    else:
        print(f"Spreadsheet '{SPREADSHEET_NAME}' not found. Creating a new one...")
        # Create a new spreadsheet
        spreadsheet = sheets_service.spreadsheets().create(
            body={
                'properties': {'title': SPREADSHEET_NAME},
                'sheets': [{'properties': {'title': SHEET_NAME}}]
            }
        ).execute()
        spreadsheet_id = spreadsheet.get('spreadsheetId')
        print(f"Spreadsheet '{SPREADSHEET_NAME}' created with ID: {spreadsheet_id}")
        return spreadsheet_id

# Function to log downloaded file in the spreadsheet
def log_file(mail_from, subject, filename, file_link, md5, sheets_service, spreadsheet_id):
    sheets_service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id,
        range=f"{SHEET_NAME}!A:C",
        valueInputOption="RAW",
        body={"values": [[mail_from, subject, filename, file_link, md5, datetime.datetime.now().isoformat()]]}
    ).execute()
    print(f"Logged {filename} in the spreadsheet.")
    #-#process_invoice(filename, sheets_service)

def is_file_logged(filename, md5, sheets_service, spreadsheet_id):
    sheet = sheets_service.spreadsheets()
    result = sheet.values().get(spreadsheetId=spreadsheet_id, range=f"Files!C:E").execute()
    values = result.get('values', [])

    for row in values:
        if row[0] == filename and row[2] == md5:
            return True
    return False

# Function to check if a Google Drive folder exists and create it if needed
def get_or_create_drive_folder(drive_service, folder_name, parent_folder_id=None):
    query = f"name = '{folder_name}' and mimeType = 'application/vnd.google-apps.folder'"
    if parent_folder_id:
        query += f" and '{parent_folder_id}' in parents"

    results = drive_service.files().list(q=query, fields="files(id, name)").execute()
    files = results.get('files', [])

    if files:
        folder_id = files[0]['id']
        print(f"Folder '{folder_name}' found with ID: {folder_id}")
        return folder_id
    else:
        print(f"Folder '{folder_name}' not found. Creating...")
        folder_metadata = {
            'name': folder_name,
            'mimeType': 'application/vnd.google-apps.folder'
        }
        if parent_folder_id:
            folder_metadata['parents'] = [parent_folder_id]

        folder = drive_service.files().create(body=folder_metadata, fields='id').execute()
        print(f"Folder '{folder_name}' created with ID: {folder['id']}")
        return folder['id']

# Function to upload a file to a specific Google Drive folder
def upload_file_to_drive(drive_service, folder_id, file_path):
    file_name = os.path.basename(file_path)
    media = MediaFileUpload(file_path, resumable=True)

    file_metadata = {
        'name': file_name,
        'parents': [folder_id]
    }
    uploaded_file = drive_service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id'
    ).execute()

    file_id = uploaded_file['id']

    # Make the file shareable and get the shareable link
    drive_service.permissions().create(
        fileId=file_id,
        body={'type': 'anyone', 'role': 'reader'}
    ).execute()
    file_link = f"https://drive.google.com/file/d/{file_id}/view"

    print(f"Uploaded file '{file_name}' to Google Drive (ID: {file_id})")
    print(f"Shareable Link: {file_link}")
    return file_link

def detect_invoice_type(file_path):
    with pdfplumber.open(file_path) as pdf:
        text = "\n".join(page.extract_text() for page in pdf.pages)

    if "pango" in text and "תינובשח" in text:
        return "Pango"
    return "Unknown"

# Function to process a Pango invoice
def process_pango_invoice(file_path, sheets_service):
    with pdfplumber.open(file_path) as pdf:
        text = "\n".join(page.extract_text() for page in pdf.pages)

    invoice_number = re.search(r"חשבונית מס/קבלה מספר: (\d+)", text).group(1)
    invoice_date = re.search(r"(\d{2}\.\d{2}\.\d{4})", text).group(1)
    customer_name = re.search(r"לכבוד\n(.*)\n", text).group(1)
    customer_vat = re.search(r"(\d{9})", text).group(1)
    total_amount = re.search(r"סה""כ לתשלום: (\d+\.\d{2})", text).group(1)

    # Update Google Sheet
    sheets_service.spreadsheets().values().append(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{INV_SHEET_NAME}!A:E",
        valueInputOption="RAW",
        body={"values": [[invoice_number, invoice_date, customer_name, customer_vat, total_amount]]}
    ).execute()

    print(f"Processed Pango Invoice: {invoice_number}")

# Meta function to process any invoice
def process_invoice(file_path, sheets_service):
    invoice_type = detect_invoice_type(file_path)

    if invoice_type == "Pango":
        process_pango_invoice(file_path, sheets_service)
    else:
        print("Unknown invoice type. Cannot process.")

def download_attachments_or_links(service, sheets_service, drive_service, spreadsheet_id, messages, destination_dir, label_id, folder_id):
    if not os.path.exists(destination_dir):
        os.makedirs(destination_dir)

    for msg in messages:
        msg_id = msg['id']
        message = service.users().messages().get(userId='me', id=msg_id).execute()
        payload = message['payload']
        headers = payload.get('headers', [])
        parts = payload.get('parts', [])
        subject = next((h['value'] for h in headers if h['name'] == 'Subject'), "No Subject")
        mail_from = next((h['value'] for h in headers if h['name'] == 'From'), "No Subject")
        print(f"Processing email: {subject}")

        for part in parts:
            if part['filename']:
                attachment_id = part['body']['attachmentId']
                attachment = service.users().messages().attachments().get(
                    userId='me', messageId=msg_id, id=attachment_id).execute()
                file_data = base64.urlsafe_b64decode(attachment['data'].encode('UTF-8'))
                filename = part['filename']
                md5 = hashlib.md5(file_data).hexdigest()
                if is_file_logged(filename, md5, sheets_service, spreadsheet_id):
                    print(f"Skipping {filename} (already logged)")
                    continue


                # Save file and log it
                file_path = os.path.join(destination_dir, filename)
                with open(file_path, 'wb') as f:
                    f.write(file_data)
                file_link = upload_file_to_drive(drive_service, folder_id, filename)
                log_file(mail_from, subject, filename, file_link, md5,  sheets_service, spreadsheet_id)
                print(f"Downloaded: {filename}")
            if 'body' in part:
                body_data = part['body'].get('data', '')
                if body_data:
                    body_text = base64.urlsafe_b64decode(body_data).decode('utf-8')
                    links = re.findall(r'https?://\S+', body_text)
                    for link in links:
                        link = link.split('>')[0]
                        link = link.split('"')[0]
                        if is_blacklisted(link): 
                            continue
                        print('link', link)
                        try:
                            response = requests.get(link, allow_redirects=True)
                            content_type = response.headers.get('Content-Type', '')
                            if response.status_code == 200 and 'text/html' not in content_type:
                                filename = os.path.join(destination_dir, os.path.basename(response.url))
                                md5 = hashlib.md5(response.content).hexdigest()
                                if is_file_logged(filename, md5, sheets_service, spreadsheet_id):
                                    print(f"Skipping {filename} (already logged)")
                                    continue
                                with open(filename, 'wb') as f:
                                    f.write(response.content)
                                file_link = upload_file_to_drive(drive_service, folder_id, filename)
                                log_file(mail_from, subject, filename, file_link, md5, sheets_service, spreadsheet_id)
                                print(f"Downloaded file from link: {filename}")
                            else:
                                print('skipped download', response.status_code, content_type)
                        except Exception as e:
                            print(f"Failed to download link {link}: {e}")

 

        # Process links in email body
        body = message.get('snippet', '')
        links = re.findall(r'https?://\S+', body)
        for link in links:
            if is_blacklisted(link):
                print(f"Skipped blacklisted link: {link}")
                continue

            try:
                response = requests.get(link, allow_redirects=True)
                if response.status_code == 200 and 'text/html' not in response.headers.get('Content-Type', ''):
                    filename = os.path.basename(response.url)
                    md5 = hashlib.md5(response.content).hexdigest()
                    if is_file_logged(filename, md5, sheets_service, spreadsheet_id):
                        print(f"Skipping {filename} (already logged)")
                        continue

                    file_path = os.path.join(destination_dir, filename)
                    md5 = hashlib.md5(response.content).hexdigest()
                    if is_file_logged(filename, md5, sheets_service, spreadsheet_id):
                        print(f"Skipping {filename} (already logged)")
                        continue
                    with open(filename, 'wb') as f:
                        f.write(response.content)
                    file_link = upload_file_to_drive(drive_service, folder_id, filename)
                    log_file(mail_from, subject, filename, file_link, md5, sheets_service, spreadsheet_id)
                    print(f"Downloaded file from link: {filename}")
            except Exception as e:
                print(f"Failed to download link {link}: {e}")

def is_blacklisted(link):
    BLACKLIST = [
        'mail-point.png', 'mail_footer_no_marketing.png', 'logo.png',
        'meser10', 'Banner.png', 'fonts.googleapis.com',
        'phone-icon-2x.png', 'link-icon-2x.png',
        'GSR-Horizontal-Black_2x.png', 'www.pango.co.il',
        'https://gsr-it.com/'
    ]
    return any(b in link for b in BLACKLIST)

# Main function
def main():
    senders = ["DoNotReply@pango.co.il", "bill@gsr-it.com"]
    unread_only = False
    destination_dir = "./attachments"
    label_name = "AutoDownloaded"

    gmail_service, sheets_service, drive_service = authenticate_services()
    spreadsheet_id = get_or_create_spreadsheet(sheets_service, drive_service)

    # Ensure the label exists
    labels = gmail_service.users().labels().list(userId='me').execute().get('labels', [])
    label_id = next((label['id'] for label in labels if label['name'] == label_name), None)
    if not label_id:
        label = gmail_service.users().labels().create(
            userId='me', body={'name': label_name}).execute()
        label_id = label['id']

    # Get and process emails
    query = f"to:incoming_invoice@gsr-it.com"
    if unread_only:
        query += " is:unread"
    messages = gmail_service.users().messages().list(userId='me', q=query).execute().get('messages', [])

    if messages:
        print("New emails found.")
        folder_id = get_or_create_drive_folder(drive_service, "ClearPay_Invoices")
        download_attachments_or_links(gmail_service, sheets_service, drive_service, spreadsheet_id, messages, destination_dir, label_id, folder_id)
    else:
        print("No new emails found.")

if __name__ == '__main__':
    main()

