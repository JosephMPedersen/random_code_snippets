# -*- coding: utf-8 -*-
"""
Created on Sat Apr 15 20:42:32 2023

@author: joseph.m.pedersen

A script for writing all Outlook calendar appointmens to file, with EntryID, 
Subject, Start, and End.

Stack Overflow posts I looked up for this:

https://stackoverflow.com/questions/69279817/create-outlook-appointment-in-subfolder-subcalendar-with-python

Outlook OlDefaultFolders enumeration
https://learn.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders
"""

import win32com.client as win32

# Open Outlook
outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

# folder 9 is the calendar
calendar = outlook.GetDefaultFolder(9) 

# The appointments
appointments = calendar.Items

# Write their information to file
with open("outlook_appts.txt", "w+") as the_file:
    for appt in appointments:
        print(f"{appt.EntryID}\n\t{appt.Subject}: {appt.Start} to {appt.End}",
              file=the_file)
