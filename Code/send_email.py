#!/usr/bin/env python3

"""Emails a custom message to a list of actors/actresses in an Excel spreadsheet."""

__author__ = 'Luke Swaby (lds20@ic.ac.uk)'
__version__ = '0.0.1'

## Imports ##
import os
import pandas as pd
import argparse
import tkinter as tk
from tkinter import filedialog
import core  # import custom module
from core import CustomizedSMPTSession
import warnings

# TODO:
#  1. Test across platforms (compare to manual send)
#  2. Run script, inputting and confirming incorrect password. See what error it throws. Then delete password ("ionos", "casting...")
#  3. Select HTML signature once then store in keychain - give option to change also (while loop with 3 options? change/continue...)
#  4. Make a change_keychain_password script

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
        print(f"Document '{os.path.basename(doc)}' added.")

        doc_add = core.yes_no("Do you wish to add another document? ('y'/'n'): ")
        while doc_add:
            doc = filedialog.askopenfilename()  # fetch doc
            docs_to_add.append(doc)
            print(f"Document '{os.path.basename(doc)}' added.")
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
         subject: str, text_path: str, docs_to_add: list,
         sign: bool, all: bool = False, preview: bool = True, ghost: bool = False):
    """
    Function that sends email to a load of addresses, replacing '-' with their names
    """
    # Read data
    df = pd.read_excel(data, keep_default_na=False)

    if not all:
        if any(df['CONTACT?']):
            df = df.loc[df['CONTACT?'].astype(bool)]  # subset only those you wish to contact
        else:
            # if no --all flag and nothing in CONTACT?, raise exception
            raise Exception("ERROR: Nothing found in the 'CONTACT?' column and --all flag not used. "
                            "Please use one or the other.")

    # Login to email account
    host = providers[provider]
    session = CustomizedSMPTSession(host, 587)
    from_address, password = session.repeat_attempt_login(provider, from_address, password, return_creds=True)

    # Check user is ok with email format
    if preview:
        session.preview_email(text_path, from_address, password, subject, docs_to_add, sign)

    # Mail agencies by group
    print('\nMAILING AGENCIES...')

    with open(text_path) as email:
        msg_template = email.read()  # Open text file and extract email template

    for to_address, group in df.groupby('EMAIL'):

        names = list(group.NAME)

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
    warnings.simplefilter("ignore")
    main(*parse_args())