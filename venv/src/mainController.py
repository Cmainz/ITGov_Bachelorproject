#!/usr/bin/env python3
"""
This file checks if reminders should be sent or not.

"""
from openpyxl import load_workbook
from os import getcwd, walk
from shutil import copyfile
from datetime import date
from pytz import timezone
from mailAPI import send_message, service
from json import dump
from sharedScripts import event,err,creating_logfile

### Variables ###
script_name="mainController"

sheet = load_workbook(filename="mainControllerDoc\\Kontroller.xlsx")

wsCtrl = sheet["Controls"]

wsControllers = sheet["Controllers"]


today = date.today()
date_of_today = int(today.strftime("%d%m%Y"))

conInfo = {}
filesToSend = []
mailingList = []

maxControlRow = len(wsCtrl['A'])
maxContactsRow = len(wsControllers['A'])


### Classes ###


class Ctrls:
  """ Creates order of the individual Controls"""
  
  ctrls_list=[]
  
  def __init__(self, number, control, due, verification,responsible):
    self.number = number
    self.control = control
    self.due = due
    self.verification = verification
    self.responsible = responsible
    
    
    self.ctrls_list.append(self)
    
  

### Functions ###
def contact_info_func():
  """ Creates a list of available controllers and their emails"""
  for value, val in wsControllers.iter_rows(
    min_row=2,
    max_row=maxContactsRow,
    min_col=1,
    max_col=2,
    values_only=True):
    if value is None:
      continue
    else:
      conInfo[value] = val
  return conInfo

def create_ctrls():
  """ Function to find all the controls that are undone"""
  for value in wsCtrl.iter_rows(min_row=2,
                            max_row=maxControlRow,
                            min_col=1,
                            max_col=5,
                            values_only=True):
   
    if value[3] == "X":
      continue
      
    else:
      
      Ctrls(value[0],
            value[1],
            value[2],
            value[3],
            value[4]
            )

def class_maker(list_item):
  """ ensures that all controls goes into check for due func"""
  global contactInfo
  for item in list_item:
    check_for_due(item.number,
                  item.control,
                  item.due,
                  item.verification,
                  item.responsible)

  return mailingList
      
def check_for_due(value0, value1, value2, value3, value4, today_date=date_of_today) -> str:
  """ Finds the controls actions are needed"""
  global notes

  try:
    type(int(value0)) == int
    log_info = f"{value0} {value1} is due {value2}. "
    due_date=value2.date()
    due_dateint = int(due_date.strftime("%d%m%Y"))
    is_due=due_dateint - today_date
    if is_due == 0:
      notes = "Send The email!"
      send_email = True
      logs=log_info+notes
      creating_logfile(event, logs, script_name)

    elif is_due == 10000000: # send a reminder if 10 days is left
      notes = "Send a reminder! He got 10 days left"
      send_email = True
      logs = log_info + notes
      creating_logfile(event, logs, script_name)

    elif is_due == 5000000: # send a reminder if 5 days is left
      notes = "Send a reminder!"
      send_email = True
      logs = log_info + notes
      creating_logfile(event, logs, script_name)

    elif is_due == -1000000: # send a reminder if delayed 1 day
      notes = "You are late! Please finish this control before end of date"
      send_email = True
      logs = log_info + notes
      creating_logfile(event, logs, script_name)

    elif is_due == -2000000: # send a reminder if delayed 2 days
      notes = "You are 2 days late! Please finish your control or contact your nearest supervisor before end of date"
      send_email = True
      logs = log_info + notes
      creating_logfile(event, logs, script_name)

    elif is_due == -3000000: # send an email to security responsible
      notes = "this control has not been finished in time or has been incorrectly made."
      send_email = True
      logs = log_info + notes
      creating_logfile(event, logs, script_name)

    else:
      send_email=False
      notes ="Nothing will be done"
      logs = log_info + notes
      creating_logfile(event, logs, script_name)

    if value4 not in conInfo or conInfo[value4] == None:
      print(f"Update needed for {value4}")
      return "Missing Contact Information"
    elif send_email == True:
      mailingList.append([value0, value1, due_date, value3, conInfo[value4],notes])
    else:
      print(notes)
          
  except ValueError:
    error_text=f"\"{value0}\" is not an index number. Control " \
           f"\"{value1}\" will not be correctly analysed \nPlease check your " \
             "Excel Sheet"
    creating_logfile(err, error_text, script_name)
    return error_text
  
  return notes


def make_control_doc():
  """ Creates an email template from the controls"""
  for item in mailingList:
    control = item[1] + ".xlsx"
    control_title = str(item[0]) + " " + item[1] + " " + str(item[2]) + ".xlsx"
    
    for roots, dirs, files in walk("."):
      for file in files:
        if control in file:
          original = getcwd() + "\Templates" + "\\" + file
          target = getcwd() + "\Temps" + "\\" + control_title
          copyfile(original, target)
          filesToSend.append(target)
  return filesToSend

def sending_email():
  """" Sends the email """
  for item, file in zip(mailingList, filesToSend):
    title=str(item[0])+" "+item[1]+" "+str(item[2])
    send_message(service, "chr.maints@gmail.com", title,item[5], [file])
    log_info="An email with title "+title+" was sent"
    creating_logfile(event, log_info, script_name)


### Logic ###
if __name__ == "__main__":

  ## Logging ##
  log_info=("Script has been initiated")
  creating_logfile(event,log_info,script_name)
  ## End Log ##

  contact_info_func()
  create_ctrls()
  class_maker(Ctrls.ctrls_list)
  make_control_doc()
  sending_email()

  ## Logging ##
  log_info=f"{len(mailingList)} Control(s) was sent out of {maxControlRow-1}"
  creating_logfile(event, log_info, script_name)
  ## End Log ##

  with open("missingControls.json", "w") as out_file:
    dump(mailingList, out_file,indent=6,default=str)
sheet.close()

