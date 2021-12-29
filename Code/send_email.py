#!/usr/bin/env python3

"""Emails a custom message to a list of actors/actresses in an Excel spreadsheet."""

__author__ = 'Luke Swaby (lds20@ic.ac.uk)'
__version__ = '0.0.1'

## Imports ##
import re
import markdown
import pandas as pd
import argparse
from pwinput import pwinput
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import tkinter as tk
from tkinter import filedialog


# TODO: add email attachments signature:
#  https://realpython.com/python-send-email/#sending-fancy-emails
# bold = bold{...}
# italics = italics{...}
#
# show example template email before sending

## Variables ##
providers = {"gmail": "smtp.gmail.com",
             "hotmail": "smtp.live.com",
             "ionos": "smtp.ionos.de",
             "icloud": "smtp.mail.me.com"}

## Functions ##
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
    parser.set_defaults(open=False)

    args = parser.parse_args()

    # Parse paths to input files
    print('\nPLEASE FILL THE FOLLOWING:\n')

    excel_prompt = 'Press ENTER to find and select Excel spreadsheet: '
    input(excel_prompt)
    root = tk.Tk()  # Initialise dialog box
    root.withdraw()
    data = filedialog.askopenfilename()
    print(' '*len(excel_prompt) + '\033[A' + data)

    text_prompt = 'Press ENTER to find and select email template .txt file: '
    input(text_prompt)
    text = filedialog.askopenfilename()
    print(' '*len(text_prompt) + '\033[A' + text)

    root.destroy()  # remove root window

    # Parse email info
    subject = input("Email Subject: ")
    usn = input(f"Sender {args.provider.title()} Email Address: ")
    pwd = pwinput("Account Password: ")

    return args.provider, data, usn, pwd, subject, text, args.all

def convert_to_html(text):
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

def main(provider, data, from_address, password, subject, text, all):
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

    print('\nLogin Successful. Sending mail now...'.upper())

    for _, row in df.iterrows():

        # Extract and format name
        names = row.NAME.split()
        firstname = names[0]
        surname = ' '.join(names[1:])
        to_address = row.EMAIL

        print(f'Mailing {to_address} regarding {firstname} {surname}...')

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

        session.sendmail(from_address, to_address, msg.as_string())  # send email

    session.quit()
    print('\nDone!')

    return

if __name__ == '__main__':
    main(*parse_args())