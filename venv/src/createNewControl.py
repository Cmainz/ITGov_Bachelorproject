#!/usr/bin/env python3
"""
This file creates a new control from mainControls.xlsx file.
For the control to function correctly a Template also needs to be written for each new control
Please review "Verify Screening processes.xlsx" for hints of creating a new template

"""

from openpyxl import load_workbook
from openpyxl.styles import NamedStyle
from datetime import date
from sharedScripts import input_to_excel,event,err,creating_logfile,check_contacts_emails,new_contacts_update
from mailAPI import service,send_mail

script_name="createNewControl"

sheet = load_workbook(filename="mainControllerDoc\\mainControls.xlsx")
ws_ctrl = sheet.active
max_control_row = len(ws_ctrl['A'])
all_main_ctrls=set()


prod_controller_file="mainControllerDoc\\Kontroller.xlsx"
production_sheet = load_workbook(filename=prod_controller_file)
ws_prod_ctrl = production_sheet["Controls"]
max_prod_ctrl_row=len(ws_prod_ctrl['A'])

wsControllers = production_sheet["Controllers"]
maxContactsRow = len(wsControllers['A'])

all_prod_ctrls=set()


ctrl_dict={}
list_for_excel=[]
contacts_dict={}
mail_service=[]


def set_ctrl(worksheet, max_row, final_set):
  """ Creates a set from an excel sheet"""

  for pages in worksheet:
    for row in pages.iter_rows(min_row=2,
                               max_row=max_row,
                               min_col=2,
                               max_col=5,
                               values_only=True):

      if row[0]== None:
          continue
      else:
        final_set.add(row[0])
        ctrl_dict[row[0]]=(row[1],row[2],row[3])
  return final_set



def check_for_match(a_ctrls, b_ctrls):
  """ Function that verifies that controls are in production and if not creates one"""

  for controls in a_ctrls:
    date=ctrl_dict[controls][0]
    responsible=ctrl_dict[controls][1]
    contact=ctrl_dict[controls][2]
    if controls in b_ctrls:
      continue
    else:
      list_for_excel.append((controls, date, responsible))
      log_info = f"\"{controls}\" has been inserted with date {date} and responsible {responsible}"
      creating_logfile(event,log_info,script_name)

      mail_service.append((controls,date,responsible,contact))
      if contact == None or responsible ==None :
        continue
      else:
        contacts_dict[responsible]=contact



## Logging ##
log_info = ("Script has been initiated")
creating_logfile(event, log_info, script_name)
## End Log ##


set_ctrl(sheet, max_control_row, all_main_ctrls)
set_ctrl(production_sheet, max_prod_ctrl_row, all_prod_ctrls)
check_for_match(all_main_ctrls, all_prod_ctrls)

input_to_excel(list_for_excel)

new_contacts_update(check_contacts_emails(contacts_dict),script_name)


send_mail(mail_service,script_name)
sheet.close()

## Logging ##
log_info=f"{len(list_for_excel)} new control was created"
creating_logfile(event,log_info,script_name)
## End Log ##