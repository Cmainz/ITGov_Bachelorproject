#!/usr/bin/env python3
from openpyxl import load_workbook
from datetime import date

def input_to_excel(chosen_ctrls):
  """ Inserts the missing ctrls into production """

  prod_controller_file = "mainControllerDoc\\Kontroller.xlsx"
  production_sheet = load_workbook(filename=prod_controller_file)
  ws_prod_ctrl = production_sheet.active
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

