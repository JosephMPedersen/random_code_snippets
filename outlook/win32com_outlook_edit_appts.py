# -*- coding: utf-8 -*-
"""
Created on Sat Apr 15 20:42:32 2023

@author: joseph.m.pedersen

A script for openning an Outlook appointment by EntryID, and changing some of
its properties.

Stack Overflow posts I looked up for this:

https://stackoverflow.com/questions/57285899/is-there-a-way-to-get-id-of-particular-email-id

"""

import win32com.client as win32

# The EntryID of the appointment to change
this_appt = ''

# Open Outlook
outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Get that appointment
appt = outlook.GetItemFromID(this_appt)

# Change some of its attributes
appt.Subject = 'Updated Entry'
appt.Duration = 35

# Save the changes
appt.Save()
