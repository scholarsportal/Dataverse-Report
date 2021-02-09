# https://www.freecodecamp.org/news/send-emails-using-code-4fcea9df63f/
import os.path
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from string import Template
from metrics.config import config


def get_contacts(filename):
    names = []
    emails = []
    with open(filename, mode='r', encoding='utf-8') as contacts_file:
        for a_contact in contacts_file:
            names.append(a_contact.split()[0])
            emails.append(a_contact.split()[1])
    return names, emails


def read_template(filename):
    with open(filename, 'r', encoding='utf-8') as template_file:
        template_file_content = template_file.read()
    return Template(template_file_content)


def send_email():
    email_config = config.read('email')
    server = smtplib.SMTP(email_config.get('host'), email_config.get('port'))
    names, emails = get_contacts(os.path.join('metrics/email', 'emailcontacts.txt'))
    message_template = read_template(os.path.join('metrics/email', 'emailtemplate.txt'))
    for name, to_address in zip(names, emails):
        msg = MIMEMultipart()  # create a message
        message = message_template.substitute(PERSON_NAME=name.title())
        msg['To'] = to_address
        msg['From'] = email_config.get('from_address')
        msg['Subject'] = email_config.get('subject')
        msg.attach(MIMEText(message, 'plain'))
        server.sendmail(email_config.get('from_address'), to_address, msg.as_string())
        del msg
    server.quit()


# send_email(os.path.join('Dataverse Usage Report1.xlsx'))
