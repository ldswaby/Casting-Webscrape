#!/usr/bin/env python3

"""Sends out template email to actors/actresses in spreadsheet."""

__author__ = 'Luke Swaby (lds20@ic.ac.uk)'
__version__ = '0.0.1'

## Imports ##
import pandas as pd
import argparse
from pwinput import pwinput
import smtplib
from email.mime.text import MIMEText

## Functions ##

def parse_args():
    """
    Parses arguments from the command line
    """
    parser = argparse.ArgumentParser(description="Script for sending a template email to all (or a subset of) actors "
                                                 "in an Excel spreadsheet (output by spotlight_scrape.py).")

    parser.add_argument('-t', default="../Data/email.txt", dest='text',
                        help='The path to the .txt file containing the email template. Note that this script assumes '
                             'any underscores in the text are placeholders for the actor/actresses name (found in the '
                             'first field of the spreadsheet), and will accordingly replace them.')
    parser.add_argument('-d', default="../Data/", dest='data',
                        help="Path to data excel spreadsheet. At the very least, this should contain the fields: "
                             "'NAME' (actor name), 'EMAIL' (email address), and 'CONTACT?' (should this person be "
                             "emailed? - note that only people for whom this field is completely blank will be "
                             "ignored; anybody with any characters in this field will be contacted).")
    parser.add_argument('--all', dest='all', action='store_true',
                        help="Include this flag if you simply want to contact everybody listed in the spreadsheet. "
                             "This essentially overrides the function of the 'CONTACT?' field. If this flag is "
                             "omitted, then only the agents with any contents in this field will be contacted.")
    parser.set_defaults(open=False)

    args = parser.parse_args()

    subject = input("Email Subject: ")
    usn = input("Sender Email Address: ")
    pwd = pwinput("Account Password: ")

    return args.data, usn, pwd, subject, args.text, args.all

def main(data, from_address, password, subject, text, all):
    """
    Function that sedns email to a load of addresses, replacing '-' with their names
    """
    # Format inputs
    if not data.startswith('../Data/'):
        data = '../Data/' + data
    if not data.endswith('.xlsx'):
        data += '.xlsx'

    df = pd.read_excel(data, keep_default_na=False)
    if not all:
        df = df.loc[df['CONTACT?'].astype(bool)]  # subset only those you wish to contact

    # Open text file and extract email template
    with open(text) as email:
        msg_template = email.read()

    # Login to email account
    session = smtplib.SMTP("smtp.ionos.de", 587)
    session.login(from_address, password)

    for _, row in df.iterrows():

        name = row.NAME.split()[0].capitalize()
        to_address = row.EMAIL
        print(f'Mailing {to_address} regarding {name}...')

        # Insert name into email message
        content = msg_template.replace('_', name)
        msg = MIMEText(content)
        msg['Subject'] = subject
        msg['From'] = from_address   # the sender's email address
        msg['To'] = to_address  # the recipient's email address

        session.sendmail(from_address, to_address, msg.as_string())  # send email

    session.quit()

    return

if __name__ == '__main__':
    main(*parse_args())
