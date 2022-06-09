#!/usr/bin/env python3
""" This function is to find and download the correct attachments"""
from re import sub as substractor
from re import compile as compiler
from mailAPI import service
from base64 import urlsafe_b64decode
from os import path,getcwd
from json import load

senders_dict = {}
title_List = []
email_sender = []
downloadable_Msg = []



def previous_control():
  """ Reads a file created by Controller.py"""
  with open("missingControls.json") as json_file:
    sent_emails = load(json_file)
    for item in sent_emails:
      title = str(item[0]) + " " + item[1] + " " + str(item[2])
      title_List.append(title)
      sender = item[4]
      email_sender.append(sender)

  return title_List


def finding_msg_id(list_item):
  """ with a list of previous controls, finds the correct emailsender and attachment"""
  results = service.users().messages().list(userId='me', labelIds=['INBOX']).execute()
  messages = results.get('messages', [])

  pattern = compiler(r'Fwd: |FWD: |re: |Re: ')

  if not messages:
    print("No messages found.")
  else:
    for message in messages:
      msg = service.users().messages().get(userId='me', id=message['id']).execute()
      title = msg['payload']['headers'][21]['value']
      reciever = msg['payload']['headers'][6]['value'][1:-1]
      mod_title = substractor(pattern, "", title)
      if mod_title in list_item and reciever in email_sender:
        attached_file=msg['payload']['parts'][1]['filename']
        attached_name=mod_title + ".xlsx"
        if attached_file ==attached_name:
          msg_id = (msg["id"])
          downloadable_Msg.append(msg_id)

  return downloadable_Msg


def downloadable_attachment(emails):
  """ Downloads the attachment if they rememeber to attach the file"""
  filename = ""
  for item in emails:
    message = service.users().messages().get(userId='me', id=item).execute()
    try:
      att_id = message['payload']['parts'][1]['body']['attachmentId']
      att = service.users().messages().attachments().get(userId='me', messageId=message['id'], id=att_id).execute()
      data = att['data']
      file_data = urlsafe_b64decode(data.encode('UTF-8'))
      filename = message['payload']['parts'][1]['filename']
      dl_path = path.join(getcwd() + '\\' + 'Downloaded controls' + '\\' + filename)
      print("download")
      with open(dl_path, 'wb') as f:
        f.write(file_data)
        f.close()
      service.users().messages().delete(userId='me', id=item).execute()

    except KeyError:
      service.users().messages().delete(userId='me', id=item).execute()
      return "No attachments in Control. Control will be deleted"

  return filename



##############LOGIC###########

finding_msg_id(previous_control())
downloadable_attachment(downloadable_Msg)