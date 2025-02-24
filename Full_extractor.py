import os
import base64
import csv
import json
import logging
import time
from datetime import datetime
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.auth.transport.requests import Request

# Scopes for Gmail API
SCOPES = ['https://www.googleapis.com/auth/gmail.modify']

# Setup logging
logging.basicConfig(
    filename='gmail_bot.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def load_config():
    default_config = {
        'csv_directory': './emails',
        'max_retries': 3,
        'default_action': 'export'
    }
    if os.path.exists('config.json'):
        with open('config.json', 'r') as f:
            return json.load(f)
    with open('config.json', 'w') as f:
        json.dump(default_config, f)
    return default_config

def authenticate_gmail():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds

def get_emails(service, sender_email, start_date=None, end_date=None, max_retries=3):
    query = f'from:{sender_email}'
    if start_date:
        query += f' after:{start_date.strftime("%Y/%m/%d")}'
    if end_date:
        query += f' before:{end_date.strftime("%Y/%m/%d")}'
    
    for attempt in range(max_retries):
        try:
            results = service.users().messages().list(userId='me', q=query).execute()
            return results.get('messages', [])
        except HttpError as error:
            if error.resp.status in [429, 503] and attempt < max_retries - 1:
                logging.warning(f"API error {error.resp.status}, retrying in {2 ** attempt}s")
                time.sleep(2 ** attempt)
                continue
            logging.error(f"Failed to fetch emails: {error}")
            return []
        except Exception as e:
            logging.error(f"Unexpected error fetching emails: {e}")
            return []

def get_message_body(payload):
    body = ''
    try:
        if 'parts' in payload:
            for part in payload['parts']:
                if part['mimeType'] == 'text/plain' and 'data' in part['body']:
                    data = part['body']['data']
                    body += base64.urlsafe_b64decode(data).decode('utf-8', errors='ignore')
                elif 'parts' in part:
                    body += get_message_body(part)
        elif 'data' in payload['body']:
            data = payload['body']['data']
            body = base64.urlsafe_b64decode(data).decode('utf-8', errors='ignore')
    except Exception as e:
        logging.warning(f"Error decoding message body: {e}")
    return body

def get_email_details(service, msg_id):
    try:
        message = service.users().messages().get(userId='me', id=msg_id, format='full').execute()
        headers = message['payload']['headers']
        subject = next((header['value'] for header in headers if header['name'] == 'Subject'), '')
        from_email = next((header['value'] for header in headers if header['name'] == 'From'), '')
        date = next((header['value'] for header in headers if header['name'] == 'Date'), '')
        body = get_message_body(message['payload'])
        
        return {
            'id': msg_id,
            'date': date,
            'from': from_email,
            'subject': subject,
            'body': body
        }
    except Exception as e:
        logging.error(f"Error getting email details for {msg_id}: {e}")
        return None

def export_to_csv(emails, filename):
    os.makedirs(os.path.dirname(filename), exist_ok=True)
    file_exists = os.path.isfile(filename)
    
    with open(filename, 'a' if file_exists else 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['id', 'date', 'from', 'subject', 'body']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        
        if not file_exists:
            writer.writeheader()
        
        total = len(emails)
        for i, email in enumerate(emails, 1):
            if email:  # Skip if email details couldn't be fetched
                writer.writerow(email)
                print(f"Exporting {i}/{total} emails", end='\r')
        print()  # New line after completion

def read_existing_ids(filename):
    if not os.path.isfile(filename):
        return set()
    
    existing_ids = set()
    try:
        with open(filename, 'r', encoding='utf-8') as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                existing_ids.add(row['id'])
    except Exception as e:
        logging.error(f"Error reading existing CSV: {e}")
    return existing_ids

def delete_or_archive_emails(service, messages, action='delete', max_retries=3):
    for msg in messages:
        for attempt in range(max_retries):
            try:
                if action.lower() == 'delete':
                    service.users().messages().trash(userId='me', id=msg['id']).execute()
                elif action.lower() == 'archive':
                    service.users().messages().modify(
                        userId='me',
                        id=msg['id'],
                        body={'removeLabelIds': ['INBOX']}
                    ).execute()
                break
            except HttpError as error:
                if error.resp.status in [429, 503] and attempt < max_retries - 1:
                    time.sleep(2 ** attempt)
                    continue
                logging.error(f"Failed to {action} email {msg['id']}: {error}")
                break
            except Exception as e:
                logging.error(f"Unexpected error in {action} for {msg['id']}: {e}")
                break

def main():
    config = load_config()
    
    # User inputs
    sender_email = input("Enter the sender's email address: ")
    start_date_str = input("Enter start date (YYYY-MM-DD, optional): ")
    end_date_str = input("Enter end date (YYYY-MM-DD, optional): ")
    
    start_date = datetime.strptime(start_date_str, '%Y-%m-%d') if start_date_str else None
    end_date = datetime.strptime(end_date_str, '%Y-%m-%d') if end_date_str else None
    
    print("\nOptions:")
    print("1. Just export emails to CSV")
    print("2. Export and delete emails")
    print("3. Export and archive emails")
    choice = input("Enter your choice (1-3): ")
    
    csv_filename = os.path.join(config['csv_directory'], f"emails_from_{sender_email.split('@')[0]}.csv")
    
    # Authenticate and build service
    logging.info(f"Starting email processing for {sender_email}")
    creds = authenticate_gmail()
    service = build('gmail', 'v1', credentials=creds)

    # Get emails
    messages = get_emails(service, sender_email, start_date, end_date)
    if not messages:
        print(f"No emails found from {sender_email}")
        logging.info(f"No emails found from {sender_email}")
        return

    # Process emails
    existing_ids = read_existing_ids(csv_filename)
    new_emails = []
    
    total_messages = len(messages)
    for i, msg in enumerate(messages, 1):
        if msg['id'] not in existing_ids:
            email_details = get_email_details(service, msg['id'])
            if email_details:
                new_emails.append(email_details)
        print(f"Processing {i}/{total_messages} messages", end='\r')
    print()

    if new_emails:
        export_to_csv(new_emails, csv_filename)
        print(f"Exported {len(new_emails)} new emails to {csv_filename}")
        logging.info(f"Exported {len(new_emails)} new emails to {csv_filename}")
    else:
        print("No new emails found")
        logging.info("No new emails found")

    # Handle delete/archive option
    if choice == '2':
        delete_or_archive_emails(service, messages, 'delete')
        print("Emails deleted")
        logging.info("Emails deleted")
    elif choice == '3':
        delete_or_archive_emails(service, messages, 'archive')
        print("Emails archived")
        logging.info("Emails archived")

if __name__ == '__main__':
    main()