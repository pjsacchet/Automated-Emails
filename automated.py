# @author -  Patrick Sacchet
# @version 1.0 - 5/7/19
# This program will take contacts fed to it via text file, create a SMTP server and send all contacts an email containing their training schedule

import smtplib

from string import Template
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

MY_ADDRESS = None
PASSWORD = None

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

def main():
    # Grab the template, names and emails
    names, emails = get_contacts('teammembers.txt')
    message_template = read_template('messagetemplate.txt')
    # Starting stmp server for Outlook and getting login info from user
    s = smtplib.SMTP(host='smtp-mail.outlook.com', port=587)
    s.starttls()
    MY_ADDRESS = input("Email Address: ")
    PASSWORD = input("Password: ")
    s.login(MY_ADDRESS, PASSWORD)
    # For each person, put their name in the template and fill out the appropiate fields, then send the message
    for name, email in zip(names, emails):
        msg = MIMEMultipart()
        message = message_template.substitute(NAME=name.title())
        msg['From']=MY_ADDRESS
        msg['To']=email
        msg['Subject']="Army Ten Miler Training"
        msg.attach(MIMEText(message, 'plain'))
        s.send_message(msg)
        del msg
    s.quit()

if __name__ == "__main__":
    main()
