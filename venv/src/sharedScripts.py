#!/usr/bin/env python3
from openpyxl import load_workbook
from datetime import date, datetime
from pytz import timezone

event="Logs\\Event.LOG"
err="Logs\\ERROR.LOG"
europe=timezone('Europe/Paris')
time_format=('%d/%b/%Y:%H:%M:%S %z')

contacts_dict_missing={}
prod_controller_file = "mainControllerDoc\\Kontroller.xlsx"



def input_to_excel(chosen_ctrls):
  """ Inserts the missing ctrls into production """
  production_sheet = load_workbook(filename=prod_controller_file)
  ws_prod_ctrl = production_sheet['Controls']
  max_prod_ctrl_row = len(ws_prod_ctrl['A'])


  for i in range(len(chosen_ctrls)):
    input_coord = str(max_prod_ctrl_row + i+1)
    new_name = chosen_ctrls[i - 1][0]
    new_ctrl_date = chosen_ctrls[i - 1][1]
    new_responsible = chosen_ctrls[i - 1][2]
    new_coord_a = "A" + input_coord
    new_coord_b = "B" + input_coord
    new_coord_c = "C" + input_coord
    new_coord_e = "E" + input_coord
    ws_prod_ctrl[new_coord_a] = int(input_coord) - 1
    ws_prod_ctrl[new_coord_b] = new_name
    ws_prod_ctrl[new_coord_c] = new_ctrl_date
    ws_prod_ctrl[new_coord_e] = new_responsible
    coord_with_date = ws_prod_ctrl.cell(int(input_coord), 3)
    coord_with_date.number_format = 'DD-MM-YYYY'

    production_sheet.save(prod_controller_file)

def creating_logfile(file,log_info,script_name):
  time_log= datetime.now().astimezone(europe).strftime(time_format)
  if file==err:
    log_info = ("ERROR {} -- [{}] \"{}\"".format(script_name, time_log[:-9], log_info))
  else:
    log_info = ("{} -- [{}] \"{}\"".format(script_name, time_log[:-9], log_info))
  with open(file, "a") as f:
    f.write("".join(log_info))
    f.write("\n")


def check_contacts_emails(dictionary_item):
  """ The function verifies if contact info is allready in  Kontroller document """
  contacts_dict_missing = dictionary_item.copy()
  production_sheet = load_workbook(filename=prod_controller_file)
  wsControllers = production_sheet["Controllers"]
  maxContactsRow = len(wsControllers['A'])
  for contacts in dictionary_item:
    for rows in wsControllers.iter_rows(
      min_row=2,
      max_row=maxContactsRow,
      min_col=1,
      max_col=2,
      values_only=True):
      for cell in rows:
        if cell in contacts:
          contacts_dict_missing.pop(cell)
  return contacts_dict_missing


def new_contacts_update(dictionary_item,script_name):
  """ Inserts the new controller in Kontroller document"""
  production_sheet = load_workbook(filename=prod_controller_file)
  wsControllers = production_sheet["Controllers"]
  maxContactsRow = len(wsControllers['A'])
  length = 0
  for contacts, email in dictionary_item.items():
    length = length + 1

    input_coord = str(maxContactsRow + length)
    new_coord_a = "A" + input_coord
    new_coord_b = "B" + input_coord
    wsControllers[new_coord_a] = contacts
    wsControllers[new_coord_b] = email
    production_sheet.save(prod_controller_file)
    log_info=f"A new contact {contacts} was inserted with the email {email}"
    creating_logfile(event, log_info, script_name)
