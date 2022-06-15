#!/usr/bin/env python3
""" This function is to find and download the correct attachments"""
from re import sub as substractor
from re import compile as compiler
from openpyxl import load_workbook
from mailAPI import service
from base64 import urlsafe_b64decode
from os import path,getcwd
from json import load
from sharedScripts import event,err,creating_logfile
from datetime import date,datetime

script_name="downloaderAttach"

senders_dict = {}
title_List = []
email_sender = []
downloadable_Msg = []
emails_downloaded=[]
controller_contact=[]

today = date.today()
date_of_today = int(today.strftime("%d%m%Y"))

report_file_location="Reports\\Reports.xlsx"
report_file=load_workbook(filename=report_file_location)
report_sheet = report_file["Delayed"]
max_reports_controls=len(report_sheet['A'])

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
    controller_email=message['payload']['headers'][18]['value']
    try:

      att_id = message['payload']['parts'][1]['body']['attachmentId']
      att = service.users().messages().attachments().get(userId='me', messageId=message['id'], id=att_id).execute()
      data = att['data']
      file_data = urlsafe_b64decode(data.encode('UTF-8'))
      filename = message['payload']['parts'][1]['filename']
      dl_path = path.join(getcwd() + '\\' + 'Downloaded controls' + '\\' + filename)
      log_file=f"{filename} was downloaded from {controller_email}"
      creating_logfile(event,log_file,script_name)
      emails_downloaded.append(filename)
      controller_contact.append(controller_email)
      with open(dl_path, 'wb') as f:
        f.write(file_data)
        f.close()
        #service.users().messages().delete(userId='me', id=item).execute()

    except KeyError:
      log_file = f"The controller {controller_email} forgot to add attachment"
      creating_logfile(event, log_file, script_name)
      #service.users().messages().delete(userId='me', id=item).execute()
      return "No attachments in Control. Control will be deleted"

  return emails_downloaded

def reporting_func():
  """ Make a reports of all downloads"""
  reporting_length=0
  for file, emails in zip(emails_downloaded, controller_contact):
    control = file.split(".")[0]
    print(control)
    control_name = " ".join(control.split(" ")[1:-1])
    control_date = control.split(" ")[-1]
    due_date = datetime.strptime(control_date, "%Y-%m-%d")
    due_date_int = int(due_date.strftime("%d%m%Y"))
    is_due = due_date_int - date_of_today
    controller = " ".join(emails.split(" ")[0:2])
    input_coord = str(max_reports_controls + reporting_length + 1)
    new_coord_a = "A" + input_coord
    new_coord_b = "B" + input_coord
    new_coord_c = "C" + input_coord
    new_coord_d = "D" + input_coord
    new_coord_e = "E" + input_coord
    report_sheet[new_coord_a] = int(input_coord) - 1
    report_sheet[new_coord_b] = control_name
    report_sheet[new_coord_c] = control_date
    report_sheet[new_coord_e] = controller
    if is_due >= 0:
      delay = "No"
    else:
      delay = "Yes"
    report_sheet[new_coord_d] = delay
    report_file.save(report_file_location)
    reporting_length=reporting_length+1

##############LOGIC###########

## Logging ##
log_info = ("Script has been initiated")
creating_logfile(event, log_info, script_name)
## End Log ##

finding_msg_id(previous_control())
downloadable_attachment(downloadable_Msg)
zippet=zip(emails_downloaded, controller_contact)
reporting_func()

## Logging ##
log_file=f"{len(downloadable_Msg)} email(s) were found and {len(emails_downloaded)} attachment(s) were downloaded"
creating_logfile(event,log_file,script_name)
## End Log ##

