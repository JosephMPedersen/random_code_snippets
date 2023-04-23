# -*- coding: utf-8 -*-
"""
Created on Sat Apr 15 20:42:32 2023

@author: joseph.m.pedersen

A script for adding an appointment to my Outlook calendar

Stack Overflow posts I looked up for this:

https://stackoverflow.com/questions/69279817/create-outlook-appointment-in-subfolder-subcalendar-with-python

https://stackoverflow.com/questions/57285899/is-there-a-way-to-get-id-of-particular-email-id

Outlook item types:
https://learn.microsoft.com/en-us/office/vba/api/outlook.olitemtype
"""

import win32com.client as win32

# Open Outlook
outlook = win32.Dispatch('outlook.application')

# Create an appointment item
appt = outlook.CreateItem(1) # OlItemType=1 is an olAppointmentItem
appt.Start = '2023-04-16 16:16'
appt.Subject = 'My New Appt'
appt.Duration = 16

# Save it
appt.Save()

# Print the EntryID created for this appt
print(appt.EntryID)
