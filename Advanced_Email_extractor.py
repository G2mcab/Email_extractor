import os
import base64
import csv
import json
import logging
import time
import tkinter as tk
from datetime import datetime
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.auth.transport.requests import Request
from tkinter import messagebox, ttk
from threading import Thread
from queue import Queue

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
        
        # Plain text body
        body_text = get_message_body(message['payload'])
        
        # HTML body and attachments
        html_body = ''
        attachments = []
        if 'parts' in message['payload']:
            for part in message['payload']['parts']:
                if part['mimeType'] == 'text/html' and 'data' in part['body']:
                    html_body = base64.urlsafe_b64decode(part['body']['data']).decode('utf-8', errors='ignore')
                elif 'filename' in part and part['filename']:
                    attachment = {
                        'filename': part['filename'],
                        'mimeType': part['mimeType'],
                        'data': base64.urlsafe_b64decode(part['body'].get('data', ''))
                    }
                    attachments.append(attachment)
        elif message['payload']['mimeType'] == 'text/html' and 'data' in message['payload']['body']:
            html_body = base64.urlsafe_b64decode(message['payload']['body']['data']).decode('utf-8', errors='ignore')
        
        return {
            'id': msg_id,
            'date': date,
            'from': from_email,
            'subject': subject,
            'body': body_text,
            'html_body': html_body,
            'attachments': attachments
        }
    except Exception as e:
        logging.error(f"Error getting email details for {msg_id}: {e}")
        return None

def export_to_csv(emails, filename, progress_queue, sender_email):
    os.makedirs(os.path.dirname(filename), exist_ok=True)
    file_exists = os.path.isfile(filename)
    
    with open(filename, 'a' if file_exists else 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)
        fieldnames = ['id', 'date', 'from', 'subject', 'body']
        dict_writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        
        if not file_exists:
            writer.writerow([f"Sender: {sender_email}", f"Total Emails: {len(emails)}"])
            dict_writer.writeheader()
        
        total = len(emails)
        for i, email in enumerate(emails, 1):
            if email:
                dict_writer.writerow({k: v for k, v in email.items() if k in fieldnames})
                progress_queue.put(('progress', i / total * 100, f"Exporting {i}/{total} emails"))
        progress_queue.put(('progress', 100, "Export complete"))

def full_extraction(emails, sender_email, folder_path, progress_queue):
    os.makedirs(folder_path, exist_ok=True)
    csv_filename = os.path.join(folder_path, f"emails_from_{sender_email.replace(' ', '_')}.csv")
    
    # CSV Export
    with open(csv_filename, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)
        fieldnames = ['id', 'date', 'from', 'subject', 'body', 'html_body', 'attachments']
        dict_writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writerow([f"Sender: {sender_email}", f"Total Emails: {len(emails)}"])
        dict_writer.writeheader()
        total = len(emails)
        for i, email in enumerate(emails, 1):
            if email:
                dict_writer.writerow({k: v if k != 'attachments' else str([a['filename'] for a in v]) for k, v in email.items()})
                progress_queue.put(('progress', i / total * 50, f"Exporting CSV {i}/{total}"))
    
    # Save Attachments and Prepare HTML
    emails_by_date = {}
    for i, email in enumerate(emails, 1):
        try:
            date_key = datetime.strptime(email['date'], '%a, %d %b %Y %H:%M:%S %z').strftime('%Y-%m-%d')
        except ValueError:
            date_key = email['date'].split(' ')[0]  # Fallback if date parsing fails
        if date_key not in emails_by_date:
            emails_by_date[date_key] = []
        emails_by_date[date_key].append(email)
        
        for att in email['attachments']:
            att_path = os.path.join(folder_path, f"attachment_{email['id']}_{att['filename']}")
            with open(att_path, 'wb') as f:
                f.write(att['data'])
            att['path'] = att_path
        
        progress_queue.put(('progress', 50 + (i / total * 30), f"Processing attachments {i}/{total}"))
    
    # HTML with JS Calendar
    html_content = """
    <!DOCTYPE html>
    <html>
    <head>
        <title>Emails from {}</title>
        <style>
            .email-container {{ border: 1px solid #ccc; margin: 10px; padding: 10px; }}
            .email-list {{ display: none; }}
            .active {{ display: block; }}
        </style>
        <script>
            function showEmails(date) {{
                document.querySelectorAll('.email-list').forEach(el => el.classList.remove('active'));
                document.getElementById('emails-' + date).classList.add('active');
            }}
        </script>
    </head>
    <body>
        <h1>Emails from {}</h1>
        <div id="calendar">
            <h2>Select Date:</h2>
            {}
        </div>
        {}
    </body>
    </html>
    """.format(sender_email, sender_email,
               ''.join(f'<button onclick="showEmails(\'{date}\')">{date}</button>' for date in sorted(emails_by_date.keys())),
               ''.join(f'<div id="emails-{date}" class="email-list">{"".join(f"<div class=\"email-container\"><h3>{e["subject"]}</h3><p>{e["html_body"] or e["body"]}</p>{"".join(f"<p>Attachment: <a href=\"{a["path"]}\" download>{a["filename"]}</a></p>" for a in e["attachments"])}</div>" for e in sorted(emails_by_date[date], key=lambda x: x["date"], reverse=True))}</div>' for date in sorted(emails_by_date.keys())))
    
    with open(os.path.join(folder_path, 'emails.html'), 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    # JSON for ML
    json_data = {
        'sender': sender_email,
        'total_emails': len(emails),
        'emails': [{k: v if k != 'attachments' else [{'filename': a['filename'], 'mimeType': a['mimeType'], 'path': a['path']} for a in v] for k, v in email.items()} for email in emails]
    }
    with open(os.path.join(folder_path, 'emails.json'), 'w', encoding='utf-8') as f:
        json.dump(json_data, f, indent=2)
    
    progress_queue.put(('progress', 100, "Full extraction complete"))

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

def delete_or_archive_emails(service, messages, action='delete', max_retries=3, progress_queue=None):
    total = len(messages)
    for i, msg in enumerate(messages, 1):
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
                if progress_queue:
                    progress_queue.put(('progress', i / total * 100, f"{action.capitalize()}ing {i}/{total} emails"))
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
    if progress_queue:
        progress_queue.put(('progress', 100, f"{action.capitalize()} complete"))

def process_emails_thread(sender_email, start_date, end_date, choice, mode, config, progress_queue):
    folder_path = os.path.join(config['csv_directory'], f"emails_from_{sender_email.replace(' ', '_')}")
    csv_filename = os.path.join(folder_path if mode == 'full' else config['csv_directory'], f"emails_from_{sender_email.replace(' ', '_')}.csv")
    
    progress_queue.put(('status', 0, "Authenticating..."))
    logging.info(f"Starting email processing for {sender_email}")
    creds = authenticate_gmail()
    service = build('gmail', 'v1', credentials=creds)
    
    progress_queue.put(('status', 10, "Fetching emails..."))
    messages = get_emails(service, sender_email, start_date, end_date)
    if not messages:
        progress_queue.put(('complete', 0, f"No emails found from {sender_email}", "info"))
        logging.info(f"No emails found from {sender_email}")
        return
    
    existing_ids = read_existing_ids(csv_filename)
    new_emails = []
    total_messages = len(messages)
    
    for i, msg in enumerate(messages, 1):
        if msg['id'] not in existing_ids:
            email_details = get_email_details(service, msg['id'])
            if email_details:
                new_emails.append(email_details)
        progress_queue.put(('progress', (i / total_messages) * 50 + 10, f"Processing {i}/{total_messages} messages"))
    
    if new_emails:
        if mode == 'simple':
            export_to_csv(new_emails, csv_filename, progress_queue, sender_email)
            progress_queue.put(('complete', 100, f"Exported {len(new_emails)} new emails to {csv_filename}", "success"))
            logging.info(f"Exported {len(new_emails)} new emails to {csv_filename}")
        else:  # full
            full_extraction(new_emails, sender_email, folder_path, progress_queue)
            progress_queue.put(('complete', 100, f"Full extraction completed for {sender_email}", "success"))
            logging.info(f"Full extraction completed for {sender_email}")
    else:
        progress_queue.put(('complete', 100, "No new emails found", "info"))
        logging.info("No new emails found")
    
    if choice == '2':
        delete_or_archive_emails(service, messages, 'delete', progress_queue=progress_queue)
        progress_queue.put(('complete', 100, "Emails deleted", "success"))
        logging.info("Emails deleted")
    elif choice == '3':
        delete_or_archive_emails(service, messages, 'archive', progress_queue=progress_queue)
        progress_queue.put(('complete', 100, "Emails archived", "success"))
        logging.info("Emails archived")

class GmailBotGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Gmail Bot")
        self.config = load_config()
        self.progress_queue = Queue()
        
        self.create_widgets()
        self.check_queue()
        
    def create_widgets(self):
        tk.Label(self.root, text="Sender Email:").grid(row=0, column=0, padx=5, pady=5)
        self.sender_entry = tk.Entry(self.root, width=40)
        self.sender_entry.grid(row=0, column=1, padx=5, pady=5)
        
        tk.Label(self.root, text="Start Date (YYYY-MM-DD):").grid(row=1, column=0, padx=5, pady=5)
        self.start_date_entry = tk.Entry(self.root, width=20)
        self.start_date_entry.grid(row=1, column=1, padx=5, pady=5, sticky='w')
        
        tk.Label(self.root, text="End Date (YYYY-MM-DD):").grid(row=2, column=0, padx=5, pady=5)
        self.end_date_entry = tk.Entry(self.root, width=20)
        self.end_date_entry.grid(row=2, column=1, padx=5, pady=5, sticky='w')
        
        tk.Label(self.root, text="Extraction Mode:").grid(row=3, column=0, padx=5, pady=5)
        self.mode_var = tk.StringVar(value="simple")
        tk.Radiobutton(self.root, text="Simple CSV Extraction", variable=self.mode_var, value="simple").grid(row=3, column=1, sticky='w')
        tk.Radiobutton(self.root, text="Full Extraction (CSV + HTML)", variable=self.mode_var, value="full").grid(row=4, column=1, sticky='w')
        
        tk.Label(self.root, text="Action:").grid(row=5, column=0, padx=5, pady=5)
        self.action_var = tk.StringVar(value="1")
        tk.Radiobutton(self.root, text="Export Only", variable=self.action_var, value="1").grid(row=5, column=1, sticky='w')
        tk.Radiobutton(self.root, text="Export and Delete", variable=self.action_var, value="2").grid(row=6, column=1, sticky='w')
        tk.Radiobutton(self.root, text="Export and Archive", variable=self.action_var, value="3").grid(row=7, column=1, sticky='w')
        
        self.process_button = tk.Button(self.root, text="Process Emails", command=self.start_processing)
        self.process_button.grid(row=8, column=0, columnspan=2, pady=10)
        
        self.progress_bar = ttk.Progressbar(self.root, length=300, mode='determinate')
        self.progress_bar.grid(row=9, column=0, columnspan=2, pady=5)
        
        self.status_label = tk.Label(self.root, text="Ready")
        self.status_label.grid(row=10, column=0, columnspan=2, pady=5)
        
    def update_progress(self, value, message):
        self.progress_bar['value'] = value
        self.status_label.config(text=message)
        self.root.update_idletasks()

    def check_queue(self):
        while not self.progress_queue.empty():
            msg_type, value, message, *args = self.progress_queue.get()
            if msg_type == 'progress':
                self.update_progress(value, message)
            elif msg_type == 'complete':
                self.update_progress(value, "Process complete")
                self.process_button.config(state='normal')
                if args[0] == "success":
                    messagebox.showinfo("Success", message)
                elif args[0] == "info":
                    messagebox.showinfo("Info", message)
            elif msg_type == 'status':
                self.update_progress(value, message)
        self.root.after(100, self.check_queue)

    def start_processing(self):
        sender_email = self.sender_entry.get()
        if not sender_email:
            messagebox.showerror("Error", "Please enter a sender email")
            return
        
        start_date_str = self.start_date_entry.get()
        end_date_str = self.end_date_entry.get()
        try:
            start_date = datetime.strptime(start_date_str, '%Y-%m-%d') if start_date_str else None
            end_date = datetime.strptime(end_date_str, '%Y-%m-%d') if end_date_str else None
        except ValueError:
            messagebox.showerror("Error", "Invalid date format. Use YYYY-MM-DD")
            return
        
        choice = self.action_var.get()
        mode = self.mode_var.get()
        
        self.process_button.config(state='disabled')
        thread = Thread(target=process_emails_thread, args=(sender_email, start_date, end_date, choice, mode, self.config, self.progress_queue))
        thread.start()

def main():
    root = tk.Tk()
    app = GmailBotGUI(root)
    root.mainloop()

if __name__ == '__main__':
    main()