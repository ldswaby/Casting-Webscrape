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
import tkinter as tk
from tkinter import filedialog
from email.message import EmailMessage


# casting@treepetts.co.uk
# noxdih-zonfav-Tafdu7


# TODO:
#  1. Why isn't email rendering on Mac Mail app? Test across platforms (compare to manual send)
#     https://stackoverflow.com/questions/30945195/trying-to-send-alternative-with-mime-but-it-also-shows-up-in-capable-mail-clie
#     https://stackoverflow.com/questions/3902455/mail-multipart-alternative-vs-multipart-mixed
#     https://www.google.com/search?q=emails+not+rendering+in+apple+mail+sent+by+python&rlz=1C5CHFA_enGB781GB783&sxsrf=AOaemvKAEqGJo-Vz7FVLQzzzEHHoHJPACg%3A1640949879384&ei=d-jOYfK2Ft6ChbIP0pKrqAQ&ved=0ahUKEwiyzp3V9o31AhVeQUEAHVLJCkUQ4dUDCA8&uact=5&oq=emails+not+rendering+in+apple+mail+sent+by+python&gs_lcp=Cgdnd3Mtd2l6EAM6BwgAEEcQsAM6BAgjECc6BAgAEEM6CgguEMcBENEDEEM6BQgAEJECOg0ILhCxAxDHARDRAxBDOgcIABCxAxBDOggIABCxAxCRAjoKCAAQsQMQgwEQQzoHCCMQ6gIQJzoRCC4QgAQQsQMQgwEQxwEQ0QM6CwgAEIAEELEDEIMBOg4ILhCABBCxAxDHARDRAzoLCC4QgAQQsQMQgwE6CAgAEIAEELEDOgUIABCABDoKCAAQgAQQhwIQFDoGCAAQFhAeOggIIRAWEB0QHjoFCCEQoAE6BwghEAoQoAE6BAghEBVKBAhBGABKBAhGGABQ8AVYn2pgjm5oCHACeACAAYABiAGbI5IBBDQ2LjiYAQCgAQGwAQrIAQjAAQE&sclient=gws-wiz
#     https://stackoverflow.com/questions/55036268/sending-email-in-python-mimemultipart
#  2. Group by email address so you don't send multiple emails to the same person (Then for each group, load all
#     suggested names; e.g. 'Thanks for suggesting Luke Cage, Aidan Swaby, and Penelope Smith...'
#  3. Change all He/She or Her/Him to They and Them - ask Mum about this
#



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
    pwd = pwinput(f"{args.provider.title()} Password: ")

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
    #out = text.replace('*', '')

    left_col_tags = re.findall(r'\[\w+\]\{', text)
    rm = ['*', '}', '#', *left_col_tags]

    for x in rm:
        text = text.replace(x, '')

    return text

def attach_documents(msg, doc_list):
    """

    """
    for doc in doc_list:
        # Open file in binary mode
        with open(doc, "rb") as attachment:
            # Add file as application
            content = attachment.read()
            msg.add_attachment(content, maintype='application',
                               subtype=os.path.splitext(doc)[1],
                               filename=os.path.basename(doc))

    return msg

def send_mail(session: smtplib.SMTP, msg_template: str, subject: str,
              from_address: str, to_address: str, names: str,
              docs_to_add: list = None, sign: bool = False):
    """
    Send mail
    """
    # Insert name into email message
    content = msg_template.replace('$N', names)

    msg = EmailMessage()
    msg['Subject'] = subject
    msg['To'] = to_address
    msg['From'] = from_address

    msg.set_content(convert_to_plain(content))
    msg.add_alternative(convert_to_html(content, sign), subtype='html')

    if docs_to_add:
        msg = attach_documents(msg, docs_to_add)

    session.send_message(msg)

    return

def main(provider: str, data: str, from_address: str, password: str,
         subject: str, text: str, docs_to_add: list,
         sign: bool, all: bool, preview: bool):
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

    print('\nLOGIN SUCCESSFUL\n')

    # Check user is ok with email format
    if preview:
        send_mail(session, msg_template, subject, from_address, from_address, '$NAMES', docs_to_add)
        email_ok_prompt = f"A formatted email has been sent to {from_address} for you to inspect. " \
                          f"Are you happy to proceed with contacting agencies? ('y'/'n'): "
        email_ok = yes_no(email_ok_prompt)

        if email_ok:
            print('\nMAILING AGENCIES...')
        else:
            sys.exit('\nPROGRAM TERMINATED\n')

    # Mail agencies by group
    for to_address, group in df.groupby('EMAIL'):

        names = list(group.NAME)

        if len(names) == 1:
            n_string = names[0]
        elif len(names) == 2:
            n_string = ' and '.join(names)
        else:
            n_string = ', '.join(names[:-1]) + f", and {names[-1]}"

        print(f'Mailing {to_address} regarding {n_string}...')

        send_mail(session, msg_template, subject, from_address, to_address, n_string, docs_to_add, sign)

    session.quit()
    print('\nDone!')

    return

if __name__ == '__main__':
    main(*parse_args())