#!/usr/bin/env python3
from openpyxl import load_workbook
from os import getcwd, walk
from shutil import move

from datetime import date
from openpyxl.styles import NamedStyle
from sharedScripts import input_to_excel

#downloaded sheets
dl_path = getcwd() + '\\' + 'Downloaded controls'



# Main excel with controls
prod_controller_file= "mainControllerDoc\\Kontroller.xlsx"
production_sheet = load_workbook(filename=prod_controller_file)
ws_prod_ctrl = production_sheet.active
max_prod_ctrl_row = len(ws_prod_ctrl['A'])


found_files=[]
dl_controls = []
validatedControls = []
input_list_excel=[]

def date_to_excel(day, month, year):
  """ Takes a date and make it readable for excel"""
  offset = 693594
  current = date(year, month, day)
  n = current.toordinal()
  return (n - offset)

def finding_files():
  """ Finds the controls"""
  for roots, dirs, files in walk(dl_path):
    for file in files:
      found_files.append(file)

def check_completion(list_item):
  """ Checks if controls have been filled"""
  for file in list_item:
    sheet = load_workbook(dl_path + '\\' + file)
    ws = sheet.active
    max_control_row = len(ws['B'])
    for value in ws.iter_rows(
      min_row=1,
      max_row=max_control_row,
      min_col=2,
      max_col=10,
      values_only=True):
      if value[0] != None:
        dl_controls.append(value[3])
  if (validating_control(dl_controls) == 100.0):
    if isinstance(value[6], date) and value[7]!= None:
      validatedControls.append([file,value[6],value[7],value[8]])
    else:
      print("control Went bad")
      print("The Date value \"",value[6],  "\" or responsible value \"",value[7],"\" is not correct")


    sheet.close()

def validating_control(list_item):
  """ Validates if controls have been filled correctly, helping func for check_completion"""
  count = 0
  try:
    for item in list_item:
      if item.lower() == "yes":
        count += 1
    percentage = count / len(list_item) * 100
    if (percentage == 100.0):
      print("Control Done")
      print(percentage)
      return percentage
    else:
      print("Control Failed")
      return percentage
  except (ZeroDivisionError, AttributeError):
    print("list is empty. Controller forgot to finish his Control!")
    return 0


def update_controls(valid_list):
  """Finds the correct control to validate"""
  empty_string = ""

  for item in valid_list:
    validated = item[0].rsplit('.', 1)[0]
    new_ctrl_date=str(item[1]).strip()[:-9]
    new_date=int(new_ctrl_date.split("-")[2])
    new_month = int(new_ctrl_date.split("-")[1])
    new_year = int(new_ctrl_date.split("-")[0])
    new_ctrl_date=date_to_excel(new_date, new_month, new_year)


    new_responsible=item[2]
    for rows in ws_prod_ctrl.iter_rows(min_row=0,
                                       max_row=max_prod_ctrl_row,
                                       min_col=1,
                                       max_col=5, ):
      count = 0
      for cell in rows:
        if cell.value == None:
          count += 1

          if (empty_string.strip()[:-9] == validated):
            coord_of_interest = str(rows[3]).split('.')[1][:-1]
            ws_prod_ctrl[coord_of_interest] = 'X'
            production_sheet.save(prod_controller_file)
            ctrl_name=empty_string.strip()[:-20][2:]
            input_list_excel.append((ctrl_name,new_ctrl_date,new_responsible))
            move("Downloaded controls\\" + item[0], "Evidence\\"+ctrl_name + "\\" + item[0])

            empty_string = ""
          else:
            empty_string = ""

        else:
          if count == 4:
            empty_string = ""
          else:
            empty_string += str(cell.value) + " "
            count += 1

  input_to_excel(input_list_excel)



finding_files()
check_completion(found_files)
update_controls(validatedControls)