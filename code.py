import os.path
import time
import pandas as pd
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import re

SCOPES = ['https://www.googleapis.com/auth/gmail.compose']

def get_service():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json')
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return build('gmail', 'v1', credentials=creds)

def create_draft(service, subject, html, from_address, to_address):
    message = MIMEMultipart('alternative')
    message['From'], message['To'], message['Subject'] = from_address, to_address, subject
    message.attach(MIMEText(html, 'html'))
    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
    draft = service.users().drafts().create(userId='me', body={'message': {'raw': raw_message}}).execute()
    print(f'Draft Id: {draft["id"]}')


def is_valid_email(email):
    pattern = r"[^@]+@[^@]+\.[^@]+"  
    return re.match(pattern, email) is not None

def send_emails_based_on_csv(service, filename, from_address):
    df = pd.read_csv(filename)
    required_columns = ['Name', 'Email Address', 'Company Name']
    if not all(col in df.columns for col in required_columns):
        print("Missing required columns in the CSV file. Please fix it.")
        return
    for _, row in df.iterrows():
        # Check for invalid/missing data
        if pd.isnull(row['Email Address']):
            print("Invalid: Email is missing")
            continue
        elif not is_valid_email(row['Email Address']):
            print(f"Invalid: '{row['Email Address']} is not a valid email")
            continue
        if pd.isnull(row['Name']) or row['Name'].strip() == '':
            print("Invalid: Name is blank")
            continue
        if pd.isnull(row['Company Name']) or row['Company Name'].strip() == '':
            print("Invalid: Company name is blank")
            continue

        name = row['Name'].split()[0]
        company = row['Company Name']
        to_address = row['Email Address']

        subject = f"Becker's Health IT Conference" # Include your desired subject line
        html = f"""
        <html>
        <body>
            <p>Hey {name},</p> 
            <p>I'm emailing to reach out about our upcoming <a href="https://conferences.beckershospitalreview.com/hitrcm2023">Becker's Health IT + RCM Conference</a> in Chicago. We have 3,000+ attendees planned, most in executive roles in hospitals and health systems. If you're interested, we have many discounted marketing products to help {company} connect with healthcare CEOs, CFOs, CMOs and other potential leads.</p>
            <p>Would you be available for a quick chat sometime in the next week? I would love to hear a little bit more about who you are looking to connect with to see if any of Becker's marketing services would be a good fit.</p>
            <p>Best,<br>Bobby</p>
            <p style="font-size:12px;">
                <span style="color:#00008B; font-size:24px; font-weight:bold;">Your Name</span><br>
                <span style="color:gray;">Your Phone Number</span><br>
                <span style="color:gray;">Your Email</span><br>
                 <span style="color:gray;">Becker's Healthcare | 17 N State St, Ste 1800, Chicago, IL 60654</span><br>
            </p>
        </body>
        </html>
        """"" 
        # IMPORTANT!!! Use CHATGPT to draft your email into HTML. It will make your job 10 times easier. 

        # General Structure:
        # The {name} puts the propsects name in that place for each email
        # The {company} part puts the company name in that place for each email
        # <p> and </p> makes seperate paragraphs
        # <a href="https://conferences.beckershospitalreview.com/hitrcm2023">Becker's Health IT + RCM Conference</a> Hyperlinks that text with that link
        # Last part is where you can put your email signature with a customized font
        


        time.sleep(0.5)
        create_draft(service, subject, html, from_address, to_address)

service = get_service()

# Change to your email and your csv file!
send_emails_based_on_csv(service, 'YOUR_FILE.csv', 'YOUR_EMAIL@beckershealthcare.com')
