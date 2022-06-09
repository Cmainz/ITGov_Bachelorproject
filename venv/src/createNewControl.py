#!/usr/bin/env python3
"""
This file creates a new control from mainControls.xlsx file.
For the control to function correctly a Template also needs to be written for each new control
Please review "Verify Screening processes.xlsx" for hints of creating a new template

"""

from openpyxl import load_workbook
from openpyxl.styles import NamedStyle
from datetime import date
from sharedScripts import input_to_excel

sheet = load_workbook(filename="mainControllerDoc\\mainControls.xlsx")
ws_ctrl = sheet.active
max_control_row = len(ws_ctrl['A'])
all_main_ctrls=set()

prod_controller_file="mainControllerDoc\\Kontroller.xlsx"
production_sheet = load_workbook(filename=prod_controller_file)
ws_prod_ctrl = production_sheet.active
max_prod_ctrl_row=len(ws_prod_ctrl['A'])
all_prod_ctrls=set()

ctrlDict={}
list_for_excel=[]

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
        ctrlDict[row[0]]=(row[1],row[2])
  return final_set



def check_for_match(a_ctrls, b_ctrls):
  """ Function that verifies that controls are in production and if not creates one"""
  for controls in a_ctrls:
    if controls in b_ctrls:
      continue
    else:
      print("\"" + controls + "\" is missing ")
      print(ctrlDict[controls])
      list_for_excel.append((controls, ctrlDict[controls][0], ctrlDict[controls][1]))



set_ctrl(production_sheet, max_prod_ctrl_row, all_prod_ctrls)
set_ctrl(sheet, max_control_row, all_main_ctrls)

check_for_match(all_main_ctrls, all_prod_ctrls)
input_to_excel(list_for_excel)
sheet.close()