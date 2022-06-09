#!/usr/bin/env python3
"""
This file checks if reminders should be sent or not.

"""
from openpyxl import load_workbook
from os import getcwd, walk
from shutil import copyfile
from datetime import date
from mailAPI import send_message, service
from json import dump


### Variables ###

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

    due_date=value2.date()
    due_dateint = int(due_date.strftime("%d%m%Y"))

    if due_dateint - today_date == 0:
      notes = "Send The email!"
      send_email = True

    elif due_dateint - today_date == 10000000: # send a reminder if 10 days is left
      notes = "Send a reminder! He got 10 days left"
      send_email = True

    elif due_dateint - today_date == 5000000: # send a reminder if 5 days is left
      notes = "Send a reminder!"
      send_email = True

    elif due_dateint - today_date == -1000000: # send a reminder if delayed 1 day
      notes = "You are late! Please finish this control before end of date"
      send_email = True

    elif due_dateint - today_date == -2000000: # send a reminder if delayed 2 days
      notes = "You are late!"
      send_email = True

    elif due_dateint - today_date == -3000000: # send an email to security responsible
      notes = "this control has not been finished in time or has been incorrectly made."
      send_email = True

    else:
      send_email=False
      notes ="Nothing will be done"
      
    if value4 not in conInfo or conInfo[value4] == None:
      print(f"Update needed for {value4}")
      return "Missing Contact Information"
    elif send_email == True:
      mailingList.append([value0, value1, due_date, value3, conInfo[value4],notes])
    else:
      print(notes)
          
  except ValueError:
    return f"\"{value0}\" is not an index number. Control " \
           f"\"{value1}\" will not be correctly analysed \nPlease check your " \
             "Excel Sheet"
  
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
    send_message(service, "chr.maints@gmail.com", str(item[0])+" "+item[1]+" "+str(item[2]), item[5], [file])
    print(item,"was sent")


### Logic ###
if __name__ == "__main__":
  contact_info_func()
  create_ctrls()
  class_maker(Ctrls.ctrls_list)
  make_control_doc()
  sending_email()
  with open("missingControls.json", "w") as out_file:
    dump(mailingList, out_file,indent=6,default=str)
sheet.close()

