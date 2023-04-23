# -*- coding: utf-8 -*-
"""
Created on Sat Apr 15 20:42:32 2023

@author: joseph.m.pedersen

A script for sending Outlook emails via Python. (basic template)

Stack Overflow posts I looked up for this:

https://stackoverflow.com/questions/6332577/send-outlook-email-via-python

https://stackoverflow.com/questions/36400683/adding-attachment-to-email-through-outlook-python

Outlook item types:
https://learn.microsoft.com/en-us/office/vba/api/outlook.olitemtype
"""

import win32com.client as win32

outlook = win32.Dispatch('outlook.application')

# create new mail message
mail = outlook.CreateItem(0) # OlItemType 0 is an email
mail.To = 'put_your_email_address_here@gmail.com'
mail.Subject = 'Message sent by Python'

# Create either `mail.HTMLBody` or `mail.Body` (not both)
#mail.HTMLBody = """<h2>HTML Message body</h2>"""
mail.Body = """Dear Me,

I am using Python to send you an email.  This is just for practice.

Best regards,
You
"""

# To attach file(s) to the email (optional):
# list the paths to attachments (use empty list for no attachments)
attachments = [r'C:\Users\YourFirst.YourLast\Desktop'
               r'\YourFolder\file1.txt',
               r'C:\Users\YourFirst.YourLast\Desktop'
               r'\YourFolder\file2.txt',]

# add each attachment
for file_path in attachments:
    mail.Attachments.Add(file_path)

# send the message
mail.Send()
