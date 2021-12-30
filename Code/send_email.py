#!/usr/bin/env python3

"""Emails a custom message to a list of actors/actresses in an Excel spreadsheet."""

__author__ = 'Luke Swaby (lds20@ic.ac.uk)'
__version__ = '0.0.1'

## Imports ##
import re
import os
import sys

import markdown
import pandas as pd
import argparse
from pwinput import pwinput
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email import encoders
from email.mime.base import MIMEBase
import tkinter as tk
from tkinter import filedialog


# TODO: add email attachments signature:
#  https://realpython.com/python-send-email/#sending-fancy-emails
#  https://stackoverflow.com/questions/10496902/pgp-signing-multipart-e-mails-with-python
# Add footer to email - would you like to add signature? Y/N  https://stackoverflow.com/questions/60316249/how-to-include-inline-images-in-e-mail-signature-when-sent-out-with-python
# parse boolean of whether to attach file(s)
# if yes then askopenfilenames() to select all - need list (if none then empty list)
# in main(): for file in files: then all code


# show example template email before sending

## Variables ##
providers = {"gmail": "smtp.gmail.com",
             "hotmail": "smtp.live.com",
             "ionos": "smtp.ionos.de",
             "icloud": "smtp.mail.me.com"}

## Functions ##
def yes_no(prompt):
    x = input(prompt).lower()

    while x not in ['y', 'n']:
        x = input("Please enter 'y' (yes) or 'n' (no): ").lower()

    return True if x == 'y' else False

def parse_args():
    """
    Parses arguments from the command line
    """
    parser = argparse.ArgumentParser(description="Script for sending a template email to all (or a subset of) actors "
                                                 "in an Excel spreadsheet (output by spotlight_scrape.py).")

    parser.add_argument('-p', dest='provider', default="ionos", choices=list(providers.keys()),
                        help=f"Email service provider. Choose one of: {', '.join(providers.keys())}.")
    parser.add_argument('--all', dest='all', action='store_true',
                        help="Include this flag if you simply want to contact everybody listed in the spreadsheet. "
                             "This essentially overrides the function of the 'CONTACT?' field. If this flag is "
                             "omitted, then only the agents with any contents in this field will be contacted.")

    args = parser.parse_args()

    # Parse paths to input files
    print('\nPLEASE FILL THE FOLLOWING:\n')

    excel_prompt = 'Press ENTER to find and select Excel spreadsheet: '
    input(excel_prompt)
    root = tk.Tk()  # Initialise dialog box
    root.withdraw()
    data = filedialog.askopenfilename()
    print(' '*len(excel_prompt) + '\033[A' + os.path.basename(data))

    text_prompt = 'Press ENTER to find and select email template .txt file: '
    input(text_prompt)
    text = filedialog.askopenfilename()
    print(' '*len(text_prompt) + '\033[A' + os.path.basename(text))

    # Document add
    docs_to_add = []
    doc_add = yes_no("Do you wish to add any documents to this email? ('y'/'n'): ")
    if doc_add:
        doc = filedialog.askopenfilename()  # fetch doc
        docs_to_add.append(doc)

        doc_add = yes_no("Do you wish to add another document? ('y'/'n'): ")
        while doc_add:
            doc = filedialog.askopenfilename()  # fetch doc
            docs_to_add.append(doc)
            doc_add = yes_no("Do you wish to add another document? ('y'/'n'): ")

    root.destroy()  # remove root window

    # Parse email info
    subject = input("Email Subject: ")
    usn = input(f"Sender {args.provider.title()} Email Address: ")
    pwd = pwinput("Account Password: ")

    sign = yes_no("Would you like this email to contain your signature? If yes, then this script will assume the "
                  "relevant HTML is saved in the current directory as 'signature.txt' ('y'/'n'): ")

    # check if user wants to preview message before sending
    preview = yes_no("Would you like to preview the email before sending? ('y'/'n'): ")

    return args.provider, data, usn, pwd, subject, text, docs_to_add, sign, args.all, preview

def convert_to_html(text, sign=False):
    """

    """
    out = text.replace('\n', '<br>')  # convert linebreaks
    out = markdown.markdown(out)  # bold and italics

    # Convert colours: [green]{...} -> <span style="color: green">...</span>
    for lefttag in re.findall(r'\[\w+\]\{', out):
        col = re.search(r'\[(\w+)\]', lefttag).group(1)
        lefttag_html = f'<span style="color: {col}">'
        out = out.replace(lefttag, lefttag_html)

    out = out.replace('}', '</span>')

    # Add signature
    if sign:
        with open('signature.txt') as sig:
            signature = sig.read()
        out += ('<br><br>' + signature)

    return out

def convert_to_plain(text):
    """

    """
    out = text.replace('*', '')

    # Convert colours: [green]{...} -> ...
    for lefttag in re.findall(r'\[\w+\]\{', out):
        out = out.replace(lefttag, '')

    out = out.replace('}', '')

    return out

def attach_documents(msg, doc_list):
    """

    """
    for doc in doc_list:

        # Open file in binary mode
        with open(doc, "rb") as attachment:
            # Add file as application/octet-stream
            # Email client can usually download this automatically as attachment
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())

        # Encode file in ASCII characters to send by email
        encoders.encode_base64(part)

        # Add header as key/value pair to attachment part
        part.add_header(
            "Content-Disposition",
            f"attachment; filename= {os.path.basename(doc)}",
        )

        # Add attachment to message and convert message to string
        msg.attach(part)

    return msg

def send_mail(session, msg_template, subject, from_address, to_address, firstname, surname, docs_to_add=False, sign=False):
    """
    Send mail
    """
    # Insert name into email message
    content = msg_template.replace('$1', firstname).replace('$2', surname)
    msg = MIMEMultipart("alternative")
    msg['Subject'] = subject
    msg['From'] = from_address  # the sender's email address
    msg['To'] = to_address  # the recipient's email address

    # Add body to email
    part1 = MIMEText(convert_to_plain(content), "plain")
    part2 = MIMEText(convert_to_html(content, sign), "html")

    # Add HTML/plain-text parts to MIMEMultipart message
    # The email client will try to render the last part first
    msg.attach(part1)
    msg.attach(part2)

    # Attach documents
    if docs_to_add:
        msg = attach_documents(msg, docs_to_add)

    session.sendmail(from_address, to_address, msg.as_string())  # send email

    return

def main(provider, data, from_address, password, subject, text, docs_to_add, sign, all, preview):
    """
    Function that sedns email to a load of addresses, replacing '-' with their names
    """
    # Read data
    df = pd.read_excel(data, keep_default_na=False)
    if not all:
        df = df.loc[df['CONTACT?'].astype(bool)]  # subset only those you wish to contact

    # Open text file and extract email template
    with open(text) as email:
        msg_template = email.read()

    # Login to email account
    session = smtplib.SMTP(providers[provider], 587)
    session.login(from_address, password)

    print('\nLOGIN SUCCESSFUL')

    # Check user is ok with email format
    if preview:
        test_name1 = 'Firstname'
        test_name2 = 'Surname'
        send_mail(session, msg_template, subject, from_address, from_address, test_name1, test_name2, docs_to_add)
        email_ok_prompt = f"A formatted email has been sent to {from_address} for you to inspect. " \
                          f"Are you happy to proceed with contacting agencies? ('y'/'n'): "
        email_ok = yes_no(email_ok_prompt)

        if email_ok:
            pass
        else:
            sys.exit('Program terminated')

    for _, row in df.iterrows():

        # Extract and format name
        names = row.NAME.split()
        firstname = names[0]
        surname = ' '.join(names[1:])
        to_address = row.EMAIL

        print(f'Mailing {to_address} regarding {firstname} {surname}...')
        send_mail(session, msg_template, subject, from_address, to_address, firstname, surname, docs_to_add, sign)
        """
        # Insert name into email message
        content = msg_template.replace('$1', firstname).replace('$2', surname)
        #msg = MIMEText(content)
        msg = MIMEMultipart("alternative")
        msg['Subject'] = subject
        msg['From'] = from_address   # the sender's email address
        msg['To'] = to_address  # the recipient's email address

        # Add body to email
        part1 = MIMEText(convert_to_plain(content), "plain")
        part2 = MIMEText(convert_to_html(content), "html")

        # Add HTML/plain-text parts to MIMEMultipart message
        # The email client will try to render the last part first
        msg.attach(part1)
        msg.attach(part2)

        # Attach documents
        if docs_to_add:
            msg = attach_documents(msg, docs_to_add)

        session.sendmail(from_address, to_address, msg.as_string())  # send email
        """
    session.quit()
    print('\nDone!')

    return

if __name__ == '__main__':
    main(*parse_args())