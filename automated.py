# @author -  Patrick Sacchet
# @version 1.0 - 5/7/19
# This program will take contacts fed to it via text file, create a SMTP server and send all contacts an email containing their
#    training schedule

import smtplib
import getpass
import xlrd
import datetime
import sys

from string import Template
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

MY_ADDRESS = None
PASSWORD = None
# For each team member of mine, each has their own unique excel file mapped to their name
# Switching back and forth between work computers, this is for the Ubuntu subsystem for Windows (so it can be run from the bash shell)
# Most recent changes will make it so CRON can run it from the directory on my Raspberry Pi
TEAM_MEMBERS = {
"Sacchet":"~/Projects/STEPHENSON - SUMMER TRAINING .xlsx",
"Stephenson":"~/Projects/STEPHENSON - SUMMER TRAINING .xlsx",
"Kruegler":"~/Projects/KRUEGLER - SUMMER TRAINING .xlsx",
"Conjelko":"~/Projects/CONJELKO - SUMMER TRAINING .xlsx",
"Roberts":"~/Projects/ROBERTS - SUMMER TRAINING .xlsx",
"Garrison":"~/Projects/GARRISON - SUMMER TRAINING .xlsx",
"Rangwala":"~/Projects/RANGWALA - SUMMER TRAINING .xlsx",
"Mahmoud":"~/Projects/MAHMOUD - SUMMER TRAINING .xlsx",
"Ulrich":"~/Projects/ULRICH - SUMMER TRAINING .xlsx",
"Eskridge":"~/Projects/ESKRIDGE - SUMMER TRAINING .xlsx"}

# We will start our training on May 15th, I want this program to automatically choose which week to send dependent on how far along in the training we are
START_DATE = 15
TRAINING_WEEK = None
date_time = datetime.datetime.now()
cur_day = date_time.day
cur_month = date_time.month

# Attemping to figure out which week we should be using in the excel sheet based on which day it is today
# PLease note: while this is not the most 'efficent' solution to this problem, the reason for it primarily was the formatting in the excel sheets, since we're sending out an entire weeks worth of training
#   I wanted to make sure all days of the week would be included in our check of the current date
# First checking for the weeks in May
if(cur_month == 5):
    if(cur_day >= 13 and cur_day < 20):
        TRAINING_WEEK = 1
    elif(cur_day >= 20 and cur_day < 27):
        TRAINING_WEEK = 2
    else:
        TRAINING_WEEK = 3
# Then the weeks in June
elif(cur_month == 6):
    if(cur_day >= 3 and cur_day < 10):
        TRAINING_WEEK = 4
    elif(cur_day >= 10 and cur_day < 17):
        TRAINING_WEEK = 5
    elif(cur_day >= 17 and cur_day < 24):
        TRAINING_WEEK = 6
    else:
        TRAINING_WEEK = 7
# Next up is July
elif(cur_month == 7):
    if(cur_day >= 1 and cur_day < 8):
        TRAINING_WEEK = 8
    elif(cur_day >= 8 and cur_day < 15):
        TRAINING_WEEK = 9
    elif(cur_day >= 15 and cur_day < 22):
        TRAINING_WEEK = 10
    elif(cur_day >= 22 and cur_day < 29):
        TRAINING_WEEK = 11
    else:
        TRAINING_WEEK = 12
# Finally we have August
else:
    if(cur_day >= 5 and cur_day < 12):
        TRAINING_WEEK = 13
    elif(cur_day >= 12 and cur_day < 19):
        TRAINING_WEEK = 14
    else:
        TRAINING_WEEK = 15

# This function will take in a text file containing all team member names and their respective email addresses
# @param - filename - Name of the file containing team member names and email addresses
# @return - The lists of the team member names and email addresses
def get_contacts(filename):
    names = []
    emails = []
    # Open the file and grab each person's name and email
    with open(filename, mode='r', encoding='utf-8') as contact_file:
        for contact in contact_file:
            names.append(contact.split()[0])
            emails.append(contact.split()[1])
    return names, emails

# Function will read contents of the file and create a Template object out of it
# @param - filename - Name of file with the template for our message
# @return - Template object
def read_template(filename):
    with open(filename, 'r', encoding='utf-8') as template_file:
        template_file_content = template_file.read()
    return Template(template_file_content)

# Function will take in the week number to grab the correct row containing that week's training data
# @param - week_num - The number corresponding to that week number in the training, will be used to grab the corresponding row
# @param - train_file - The location to file particular to that specific runner
# @return - week_mileage - The miles to be run that week
def get_run_info(week_num, train_file):
    # Grab the location of the file and create a workbook
    file_loc = train_file
    wb = xlrd.open_workbook(file_loc)
    # Creating sheet and setting start point at the beginning
    sheet = wb.sheet_by_index(0)
    sheet.cell_value(0, 0)
    # Grab the row number that is passed to the function
    week_mileage = sheet.row_values(week_num)
    return week_mileage

def main():
    # Grab the template, names and emails
    names, emails = get_contacts('teammembers.txt')
    message_template = read_template('messagetemplate.txt')
    # Starting stmp server for Outlook and getting login info from user
    s = smtplib.SMTP(host='smtp-mail.outlook.com', port=587)
    s.starttls()
    # We will pass our email and password as command line arguments to CRON on our Raspberry Pi
    MY_ADDRESS = sys.argv[1]
    PASSWORD = sys.argv[2]
    s.login(MY_ADDRESS, PASSWORD)
    # For each person, put their name in the template and fill out the appropiate fields, then send the message
    for name, email in zip(names, emails):
        # Get that specific runner's info
        # Search dictionary for the current runner to get their file location
        print(name)
        train_file_loc = TEAM_MEMBERS.get(name)
        week_mileage = get_run_info(TRAINING_WEEK, train_file_loc)
        msg = MIMEMultipart()
        # Grab the mileage for each day of the week, along with the week timeframe
        message = message_template.substitute(NAME=name.title(), WEEK=str(week_mileage[0]), MONDAY=str(week_mileage[1]), TUESDAY=str(week_mileage[2]), WEDNESDAY=str(week_mileage[3]), THURSDAY=str(week_mileage[4]),
        FRIDAY=str(week_mileage[5]), SATURDAY=str(week_mileage[6]), SUNDAY=str(week_mileage[7]))
        msg['From']=MY_ADDRESS
        msg['To']=email
        msg['Subject']="Army Ten Miler Training"
        msg.attach(MIMEText(message, 'plain'))
        s.send_message(msg)
        del msg
    s.quit()

if __name__ == "__main__":
    main()
