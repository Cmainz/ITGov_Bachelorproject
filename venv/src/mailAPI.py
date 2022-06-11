#!/usr/bin/env python3

### Mail Libaries ###

"""Email module that helps mainController.py with client"""

from os import path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from email import encoders
from base64 import urlsafe_b64encode
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase

SCOPES = ['https://mail.google.com/']
with open('credentials\\email.txt') as email:sender_email= email.read()
token_location="credentials\\token.json"
creds_location="credentials\\credentials.json"


def gmail_authenticate():
    """
    function to look for and create a token from gmail credentials
    """
    creds = None
    if path.exists(token_location):
        creds = Credentials.from_authorized_user_file(token_location, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(creds_location, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_location, "w") as token:
            token.write(creds.to_json())
    return build('gmail', 'v1', credentials=creds)

service = gmail_authenticate()

def add_attachment(message, filename):
    """
       function to make an excel file into an attachment
    """
    with open(filename, "rb") as file_path:
        attached_file = MIMEBase('application', 'vnd.ms-excel')
        attached_file.set_payload(file_path.read())
        file_path.close()
        encoders.encode_base64(attached_file)
        filename = path.basename(filename)
        attached_file.add_header('Content-Disposition', 'attachment', filename=filename)
        message.attach(attached_file)

def build_message(destination, obj, body, attachments):
    """
           function to make the header,subject and body of the email
    """
    message = MIMEMultipart()
    message['to'] = destination
    message['from'] = sender_email
    message['subject'] = obj
    message.attach(MIMEText(body))
    for filename in attachments:
        add_attachment(message, filename)
    return {'raw': urlsafe_b64encode(message.as_bytes()).decode()}

def send_message(service, destination, obj, body, attachments):
    """
           function to send an  email
    """
    return service.users().messages().send(
      userId="me",
      body=build_message(destination, obj, body, attachments)
    ).execute()


