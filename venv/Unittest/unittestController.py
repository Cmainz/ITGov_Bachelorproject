#!/usr/bin/env python3
import unittest
import os
os.chdir(".."+"\\src")
from mainController import check_for_due,contact_info_func,Ctrls,mailingList
from datetime import datetime,date

test_date=19042021
test_kontroller= "Jens Hansen"
contact_email= 'jensh6247@gmail.com'

class CheckForDueUnittest(unittest.TestCase):
  def setUp(self) -> None:
    contact_info_func()

  def testForValueError(self) -> None:
    message=check_for_due("testIndex", "controlTest", 2, 3, 4)
    expected= "\"testIndex\" is not an index number. Control \"controlTest\" will not be correctly analysed \nPlease check your Excel Sheet"
    self.assertEqual(expected,message,"Error")

  def testForZeroDay(self):
    datetimestr=datetime.strptime("19.04.2021", "%d.%m.%Y")
    message = check_for_due(1, "testControlZero", datetimestr, "", test_kontroller, today_date=test_date)
    expected = "Send The email!"
    self.assertEqual(expected, message)

  def testForFiveDays(self):
    datetimestr = datetime.strptime("24.04.2021", "%d.%m.%Y")
    message = check_for_due(2, "testControlFive", datetimestr, "", test_kontroller, today_date=test_date)
    expected = "Send a reminder!"
    self.assertEqual(expected, message)

  def testForTenDays(self):
    datetimestr = datetime.strptime("29.04.2021", "%d.%m.%Y")
    message =check_for_due(3, "testControlTen", datetimestr, "", test_kontroller, today_date=test_date)
    expected= "Send a reminder! He got 10 days left"
    self.assertEqual(expected,message)

  def testForPlentyDays(self):
    datetimestr = datetime.strptime("19.07.2021", "%d.%m.%Y")
    message = check_for_due(4, "testControlPlenty", datetimestr, "", test_kontroller, today_date=test_date)
    expected = "Nothing will be done"
    self.assertEqual(expected, message)

  def testForNegativeOneDay(self):
    datetimestr = datetime.strptime("18.04.2021", "%d.%m.%Y")
    message = check_for_due(5, "testControlLateOne", datetimestr, "", test_kontroller, today_date=test_date)
    expected = "You are late! Please finish this control before end of date"
    self.assertEqual(expected, message)

  def testForNegativeTwoDays(self):
    datetimestr = datetime.strptime("17.04.2021", "%d.%m.%Y")
    message = check_for_due(6, "testControlLateTwo", datetimestr, "", test_kontroller, today_date=test_date)
    expected = "You are late!"

    self.assertEqual(expected, message)

  def testForFailedControl(self):
    datetimestr = datetime.strptime("16.04.2021", "%d.%m.%Y")
    message = check_for_due(7, "testControlFailed", datetimestr, "", test_kontroller, today_date=test_date)
    expected= 'this control has not been finished in time or has been incorrectly made.'

    self.assertEqual(expected, message)

  def testForMissingResponsibility(self):
    contact_info_func()
    datetimestr = datetime.strptime("29.04.2021", "%d.%m.%Y")
    message =check_for_due(1, 2, datetimestr, "", "Jens Tester", today_date=test_date)
    expected= "Missing Contact Information"
    self.assertEqual(expected,message)


class ClassUnittest(unittest.TestCase):
  def setUp(self):
    datetimestr = datetime.strptime("29.04.2021", "%d.%m.%Y")
    Ctrls.ctrls_list.clear()
    Ctrls(
      "1",
      'Verify Screening processes',
      datetimestr,
      "X",
      test_kontroller)
    Ctrls(
      "2",
      'Verify terms and conditions',
      datetimestr,
      "",
      test_kontroller)
    Ctrls(
      "3",
      'Verify Screening processes',
      datetimestr,
      "",
      test_kontroller)

  def test_class_maker(self):
    message=mailingList
    expected=[[7, 'testControlFailed', datetime.date(datetime.strptime("16.04.2021", "%d.%m.%Y")), '', contact_email, 'this control has not been finished in time or has been incorrectly made.'],
              [2, 'testControlFive', datetime.date(datetime.strptime("24.04.2021", "%d.%m.%Y")), '', contact_email, 'Send a reminder!'],
              [5, 'testControlLateOne', datetime.date(datetime.strptime("18.04.2021", "%d.%m.%Y")), '', contact_email, 'You are late! Please finish this control before end of date'],
              [6, 'testControlLateTwo', datetime.date(datetime.strptime("17.04.2021", "%d.%m.%Y")), '', contact_email, 'You are late!'],
              [3, 'testControlTen', datetime.date(datetime.strptime("29.04.2021", "%d.%m.%Y")), '', contact_email, 'Send a reminder! He got 10 days left'],
              [1, 'testControlZero', datetime.date(datetime.strptime("19.04.2021", "%d.%m.%Y")), '', contact_email, 'Send The email!']]

    self.assertEqual(expected, message)

  def test_ctrls_class(self):
    message=""
    for item in Ctrls.ctrls_list:
      message+=str(item.number)+" "+item.control+" "+str(item.due)+ " "+item.verification+ " "+ item.responsible +"\n"

    expected="1 Verify Screening processes 2021-04-29 00:00:00 X Jens Hansen"+"\n"+\
             "2 Verify terms and conditions 2021-04-29 00:00:00  Jens Hansen"+"\n"+\
             "3 Verify Screening processes 2021-04-29 00:00:00  Jens Hansen"+"\n"
    self.assertEqual(expected, message)

if __name__ == '__main__':
  unittest.main()
