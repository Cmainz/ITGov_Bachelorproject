#!/usr/bin/env python3

### Mail Libaries ###

"""System Module for finding creds"""
from os import path
"""google Library to build gmail credentials"""
from googleapiclient.discovery import build
"""google Library to create OAuth authorisation flow"""
from google_auth_oauthlib.flow import InstalledAppFlow
"""google Library to utilize the credentials in the project"""
from google.oauth2.credentials import Credentials
"""google Library to make the requests to google for authorisation """
from google.auth.transport.requests import Request
"""email Library to encode attachments  """
from email import encoders
"""base64 Library to encode messages to b64 """
from base64 import urlsafe_b64encode
"""email Library to replicate mime objects in python"""
from email.mime.text import MIMEText
"""email Library to make a mime class"""
from email.mime.multipart import MIMEMultipart
"""email Library to make a mime class for the attachment"""
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


