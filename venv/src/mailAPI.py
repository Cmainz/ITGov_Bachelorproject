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
from sharedScripts import err,event,creating_logfile

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


#### Without attachment ###

def build_message_without_attachment(destination, obj, body):
    """
           function to make the header,subject and body of the email
    """
    message = MIMEMultipart()
    message['to'] = destination
    message['from'] = sender_email
    message['subject'] = obj
    message.attach(MIMEText(body))
    return {'raw': urlsafe_b64encode(message.as_bytes()).decode()}

def send_without_attach(service, destination, obj, body):
    """
           function to send an  email
    """
    return service.users().messages().send(
      userId="me",
      body=build_message_without_attachment(destination, obj, body)
    ).execute()

def send_mail(list_item,script_name):
  """ sends an email without attachment to new controller"""
  for item in list_item:
    ctrl = item[0]
    date = str(item[1])
    contact = item[2]
    email = item[3]
    send_without_attach(service, "chr.maints@gmail.com", f"A new control has been created in your name: {ctrl}",
    f"You have  the responsible of the following control {ctrl} which is due {date}\n\n"
    f"10 days before the due date you will recieve the control sheet\n\n"
    f"Have a Good Day" )
    log_file=f"Information about control {ctrl} was sent to{contact} with the email {email}"
    creating_logfile(event, log_file, script_name)
