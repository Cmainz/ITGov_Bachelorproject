# ITGov_Bachelorproject

## Purpose
To ensure accessability to easy automation to all who seeks to automate Compliance controls. Focus
has been on ISO 27001/02 and without SQL and Apache server but can be adapted.


## Changelog


### Version 1.3 - Notification

#### New Features
Notification
It had been pointed out that participants only knew they were owner of a control 10 days before delivery.
As this drew some concerns, controllers will now be informed of their ownership. It is still up to the
controller to insert the appointment in their calendar.


### Version 1.2 - Reporting

#### New Features
##### Reports
After each control a script will document the status of the control. Each control will first be graded as
either delayed or not, followed by a check if the control then failed aswell. The purpose of this, is to
emphasize on the importance on finishing controls in a timely manner. In regard to this an Excel has
been created to show statistics of the runs.


### Version 1.1 - Logging
- A crash appeared as openpyxl could not parse ".md" files. ValidatingControls will now only look at Excel
files when making validation
- A bug appeared when product had to validate more than two products where one of them had status of
"Failed". ValidatingControls loop is now iterating properly.

#### New Features
##### Logging
The product will log most functionalities and save them as either event- or errorlog. The implementation
has introduced the library Pytz for Timezone Management to the product. Changes to the overall risk has
not been evident.


### Version 1.0
The product has been released and is able to perform automated controls
