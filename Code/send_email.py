#!/usr/bin/env python3

"""Emails a custom message to a list of actors/actresses in an Excel spreadsheet."""

__author__ = 'Luke Swaby (lds20@ic.ac.uk)'
__version__ = '0.0.1'

## Imports ##
import os
import sys
import pandas as pd
import argparse
import tkinter as tk
from tkinter import filedialog
import core  # import custom module
from core import CustomizedSMPTSession
import smtplib

# TODO:
#  1. Why isn't email rendering on Mac Mail app? Test across platforms (compare to manual send)
#     https://stackoverflow.com/questions/30945195/trying-to-send-alternative-with-mime-but-it-also-shows-up-in-capable-mail-clie
#     https://stackoverflow.com/questions/3902455/mail-multipart-alternative-vs-multipart-mixed
#     https://www.google.com/search?q=emails+not+rendering+in+apple+mail+sent+by+python&rlz=1C5CHFA_enGB781GB783&sxsrf=AOaemvKAEqGJo-Vz7FVLQzzzEHHoHJPACg%3A1640949879384&ei=d-jOYfK2Ft6ChbIP0pKrqAQ&ved=0ahUKEwiyzp3V9o31AhVeQUEAHVLJCkUQ4dUDCA8&uact=5&oq=emails+not+rendering+in+apple+mail+sent+by+python&gs_lcp=Cgdnd3Mtd2l6EAM6BwgAEEcQsAM6BAgjECc6BAgAEEM6CgguEMcBENEDEEM6BQgAEJECOg0ILhCxAxDHARDRAxBDOgcIABCxAxBDOggIABCxAxCRAjoKCAAQsQMQgwEQQzoHCCMQ6gIQJzoRCC4QgAQQsQMQgwEQxwEQ0QM6CwgAEIAEELEDEIMBOg4ILhCABBCxAxDHARDRAzoLCC4QgAQQsQMQgwE6CAgAEIAEELEDOgUIABCABDoKCAAQgAQQhwIQFDoGCAAQFhAeOggIIRAWEB0QHjoFCCEQoAE6BwghEAoQoAE6BAghEBVKBAhBGABKBAhGGABQ8AVYn2pgjm5oCHACeACAAYABiAGbI5IBBDQ2LjiYAQCgAQGwAQrIAQjAAQE&sclient=gws-wiz
#     https://stackoverflow.com/questions/55036268/sending-email-in-python-mimemultipart
#  2. Run script, inputting and confirming incorrect password. See what error it throws. Then delete password ("ionos", "casting...")
#  3. Select HTML signature once then store in keychain - give option to change also (while loop with 3 options? change/continue...)
#

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
    parser.add_argument('--ghost', dest='ghost', action='store_true',
                        help="If flag included will not log email")

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
    doc_add = core.yes_no("Do you wish to add any documents to this email? ('y'/'n'): ")
    if doc_add:
        doc = filedialog.askopenfilename()  # fetch doc
        docs_to_add.append(doc)

        doc_add = core.yes_no("Do you wish to add another document? ('y'/'n'): ")
        while doc_add:
            doc = filedialog.askopenfilename()  # fetch doc
            docs_to_add.append(doc)
            doc_add = core.yes_no("Do you wish to add another document? ('y'/'n'): ")

    root.destroy()  # remove root window

    sign = core.yes_no("Would you like this email to contain your signature? If yes, then this script will assume the "
                  "relevant HTML is saved in the current directory as 'signature.txt' ('y'/'n'): ")

    # check if user wants to preview message before sending
    preview = core.yes_no("Would you like to preview the email before sending? ('y'/'n'): ")

    # Parse email login info
    subject = input("Email Subject: ")
    usn = input(f"Sender {args.provider.title()} Email Address: ")
    pwd = core.fetch_password(args.provider, usn)  # Obtain keyring from keychain. Set it if absent

    return args.provider, data, usn, pwd, subject, text, docs_to_add, sign, args.all, preview, args.ghost


def create_name_string(names: list) -> str:
    """Function to take list of name strings and joins them together into a single grammatically correct string.
    """
    assert names, "Input names list has length = 0"  # debugging

    if len(names) == 1:
        n_string = names[0]
    elif len(names) == 2:
        n_string = ' and '.join(names)
    else:
        n_string = ', '.join(names[:-1]) + f", and {names[-1]}"

    return n_string


def main(provider: str, data: str, from_address: str, password: str,
         subject: str, text: str, docs_to_add: list,
         sign: bool, all: bool = False, preview: bool = True, ghost: bool = False):
    """
    Function that sends email to a load of addresses, replacing '-' with their names
    """
    # Read data
    df = pd.read_excel(data, keep_default_na=False)
    if not all:
        df = df.loc[df['CONTACT?'].astype(bool)]  # subset only those you wish to contact

    # Open text file and extract email template
    with open(text) as email:
        msg_template = email.read()

    # Login to email account
    session = CustomizedSMPTSession(providers[provider], 587)
    from_address, password = session.repeat_attempt_login(provider, from_address, password)

    # Check user is ok with email format
    # TODO: catch smtplib.SMTPSenderRefused thrown by send_email
    while preview:

        try:
            session.send_email(msg_template, subject, from_address, from_address, '$NAMES', docs_to_add, sign,
                               ghost=True)
        except smtplib.SMTPSenderRefused:
            # If session times out then re-create it
            session = CustomizedSMPTSession(providers[provider], 587)
            session.login(from_address, password)
            session.send_email(msg_template, subject, from_address, from_address, '$NAMES', docs_to_add, sign,
                               ghost=True)

        email_ok_prompt = f">> A formatted email has been sent to {from_address} for you to inspect. " \
                          f"Are you happy to forward this to agencies? ('y'/'n'): "
        email_ok = core.yes_no(email_ok_prompt)

        if email_ok:
            break
        else:
            input(f">> Please edit template at path '{text}'. Hit ENTER to re-preview when you have saved new contents.")
            with open(text) as email:
                msg_template = email.read()

    # Mail agencies by group
    print('\nMAILING AGENCIES...')
    for to_address, group in df.groupby('EMAIL'):

        names = list(group.NAME)

        # Create names string
        if len(names) == 0:
            print(f"WARNING: No names provided for {to_address}. Skipping...")
            continue  # move on to next address/group

        n_string = create_name_string(names)

        print(f'Mailing {to_address} regarding {n_string}...')
        session.send_email(msg_template, subject, from_address, to_address, n_string, docs_to_add, sign, ghost)

    session.quit()
    print('\nDone!')

    return

if __name__ == '__main__':
    main(*parse_args())