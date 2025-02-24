import os
import base64
import csv
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from datetime import datetime

# If modifying these scopes, delete the file token.json
SCOPES = ['https://www.googleapis.com/auth/gmail.modify']

def authenticate_gmail():
    creds = None
    # The file token.json stores the user's access and refresh tokens
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no valid credentials, let the user log in
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds

def get_emails(service, sender_email):
    results = service.users().messages().list(userId='me', q=f'from:{sender_email}').execute()
    messages = results.get('messages', [])
    return messages

def get_email_details(service, msg_id):
    message = service.users().messages().get(userId='me', id=msg_id, format='full').execute()
    
    headers = message['payload']['headers']
    subject = next((header['value'] for header in headers if header['name'] == 'Subject'), '')
    from_email = next((header['value'] for header in headers if header['name'] == 'From'), '')
    date = next((header['value'] for header in headers if header['name'] == 'Date'), '')

    # Get email body
    if 'parts' in message['payload']:
        parts = message['payload']['parts']
        data = parts[0]['body']['data'] if 'data' in parts[0]['body'] else ''
    else:
        data = message['payload']['body']['data'] if 'data' in message['payload']['body'] else ''
    
    body = base64.urlsafe_b64decode(data).decode('utf-8') if data else ''
    
    return {
        'id': msg_id,
        'date': date,
        'from': from_email,
        'subject': subject,
        'body': body
    }

def export_to_csv(emails, filename):
    # Check if CSV exists to determine if we append or create new
    file_exists = os.path.isfile(filename)
    
    with open(filename, 'a' if file_exists else 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['id', 'date', 'from', 'subject', 'body']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        
        if not file_exists:
            writer.writeheader()
        
        for email in emails:
            writer.writerow(email)

def read_existing_ids(filename):
    if not os.path.isfile(filename):
        return set()
    
    existing_ids = set()
    with open(filename, 'r', encoding='utf-8') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            existing_ids.add(row['id'])
    return existing_ids

def delete_or_archive_emails(service, messages, action='delete'):
    for msg in messages:
        if action.lower() == 'delete':
            service.users().messages().trash(userId='me', id=msg['id']).execute()
        elif action.lower() == 'archive':
            service.users().messages().modify(
                userId='me',
                id=msg['id'],
                body={'removeLabelIds': ['INBOX']}
            ).execute()

def main():
    # Get sender email and options from user
    sender_email = input("Enter the sender's email address: ")
    csv_filename = f"emails_from_{sender_email.split('@')[0]}.csv"
    
    print("\nOptions:")
    print("1. Just export emails to CSV")
    print("2. Export and delete emails")
    print("3. Export and archive emails")
    choice = input("Enter your choice (1-3): ")

    # Authenticate and build service
    creds = authenticate_gmail()
    service = build('gmail', 'v1', credentials=creds)

    # Get all emails from sender
    messages = get_emails(service, sender_email)
    
    if not messages:
        print(f"No emails found from {sender_email}")
        return

    # Get existing email IDs from CSV
    existing_ids = read_existing_ids(csv_filename)
    new_emails = []

    # Process emails
    for msg in messages:
        if msg['id'] not in existing_ids:
            email_details = get_email_details(service, msg['id'])
            new_emails.append(email_details)

    if new_emails:
        export_to_csv(new_emails, csv_filename)
        print(f"Exported {len(new_emails)} new emails to {csv_filename}")
    else:
        print("No new emails found")

    # Handle delete/archive option
    if choice == '2':
        delete_or_archive_emails(service, messages, 'delete')
        print("Emails deleted")
    elif choice == '3':
        delete_or_archive_emails(service, messages, 'archive')
        print("Emails archived")

if __name__ == '__main__':
    main()