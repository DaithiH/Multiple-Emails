# -*- coding: utf-8 -*-
"""
Created on Mon Jan 31 15:59:28 2022

@author: daveh
"""

## PROGRAM TO SEND AN EMAIL TO MULTIPLE RECIPIENTS ##

import pandas as pd
import smtplib

# CREDENTIALS #
my_name = "Michael D Mouse"
my_email = "mickeyDmouse@gmail.com"
my_password = "PasswordPassword123"

# CREATE A SERVER OBJECT WITH SECURE SSL/TLS CONNECTION
# CONNECT TO GMAIL THROUGH PORT NUMBER 465 #
# NEED TO CHANGE THIS FOR OTHER SERVICES #

try:
    server = smtplib.SMTP_SSL("smtp.gmail.com", 465)
    server.ehlo()
    server.login(my_email, my_password)
except:
    print('Something went wrong!')
# READ THE SPREADSHEET FILE #
email_list = pd.read_excel('EmailTest.xlsx')

# ASSIGN ALL THE NAMES, ADDRESSES, AND MESSAGES #
all_names = email_list['Name']
all_emails = email_list['Email']
#link = email_list['Link']
link = 'https://www.irishtimes.com/'
subject = "This is a test"


# LOOP THROUGH AND GET ALL EMAIL DETAILS
for idx in range(len(all_names)):
    name = all_names[idx]
    email = all_emails[idx]
    subject = subject
    #link = link[idx]

## CREATE THE MAIL TO SEND ##
full_mail = ("From: {0} <{1}>\n"
             "To: {2} <{3}>\n"
             "Subject: {4}\n\n"
             """Hi {2}, this is {0}.
I am writing this as a test to see if it works.
Did the link come through?
{5}

Thank you.
{0}"""
             .format(my_name, my_email, name, email, subject, link)

             ) ## CLOSING THE full_mail OBJECT

try:
    server.sendmail(my_email, [email], full_mail)
    #print("Email to {} sent successfully.\n" .format(email))
    print(f"Email to {email} sent successfully.\n")
except Exception as ex:
    #print("Email to {} not sent: (because {} \n" .format(email, str(ex)))
    print(f"Email to {email} not sent because: {ex} \n")

## YOU MUST CLOSE THE SERVER!!! ##

server.close()
