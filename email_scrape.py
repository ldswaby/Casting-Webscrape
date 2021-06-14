#!/usr/bin/env python3

"""Scraping emails for names, numbers and emails etc."""

__author__ = 'Luke Swaby (lds20@ic.ac.uk)'
__version__ = '0.0.1'

## Imports ##
import imaplib
from imaplib import IMAP4_SSL
import re, getpass, sys, subprocess, argparse
import email
import email.header
import datetime
import pandas as pd

emailadd = "lukeswabypetts@gmail.com"
mailbox = 'INBOX'

## Functions ##

def process_mailbox(M, subject=None):
    """
    Do something with emails messages in the folder.
    For the sake of this example, print some headers.
    """

    #TODO:
    # 1. make this only return correct emails
    # 2. make this scan for names, numbers, and emails
    # 3. Return in dataframe

    if subject:
        rv, data = M.uid('search', None, f'(HEADER Subject "{subject}")')
    else:
        rv, data = M.uid('search', None, 'ALL')

    #rv, data = M.search(None, "ALL")
    if rv != 'OK':
        return "No messages found!"

    diclist = []

    for num in data[0].split():
        #rv, data = M.fetch(num, '(RFC822)')
        rv, data = M.uid('fetch', num, '(RFC822)')
        # TODO: WHY IS THERE NO DATA?

        if rv != 'OK':
            print("ERROR getting message", num)
            #return

        msg = email.message_from_string(data[0][1].decode()) # .decode() added
        decode = email.header.decode_header(msg['Subject'])[0]
        subject = str(decode[0])
        print('Message %s: %s' % (num.decode(), subject))  # added .decode()
        print('Raw Date:', msg['Date'])

        # Now convert to local date-time
        date_tuple = email.utils.parsedate_tz(msg['Date'])
        if date_tuple:
            local_date = datetime.datetime.fromtimestamp(
                email.utils.mktime_tz(date_tuple))
            print("Local Date:", local_date.strftime("%a, %d %b %Y %H:%M:%S"))

        ## see https://stackoverflow.com/questions/48985722/finding-links-in-an-emails-body-with-python
        for part in msg.walk():
            # each part is a either non-multipart, or another multipart message
            # that contains further parts... Message is organized like a tree
            if part.get_content_type() == 'text/plain':
                plain_text = part.get_payload()
                #print('*'*50)
                #print('MESSAGE:')
                #print('*'*50)
                #print(plain_text)

                #string = 'asdghmjsa  NAME: Luke Swaby-Petts  kjdkjsdsa' \
                #         'n PHONE: 07966283252 sjhdjhasKs asjhgdbkal' \
                #         'NAME: Tiana Milanovich  PHONE: 07595 946 214'

                #string = 'dkjagdkasjdgsahfjdjghadhkgDALJDKGJF\r\n\r\nMDAUDJAGHFDFH\r\n\r\nNAME: Tiana Milanovich\r\nPHONE: 07985 768 657\r\n\r\ntn  NAME: Luke Swabo\nPHONE:07595946214 xx\r\n\r\ntdot\r\n'
                match = r"NAME:?\s*([A-Za-z ,.'-]+)\s+PHONE:?\s*([\d ]+)"
                regex = re.compile(match)

                #re.findall(r"NAME:?\s*([A-Za-z ,.'-]+)\s+PHONE:?\s*([\d ]+)", string)

                # TODO: Why doesn't this^ match \r\n\r\n? (\r\n)

                for match in regex.finditer(plain_text):
                    diccy = {}

                    name = match.group(1).strip()
                    no = match.group(2).replace(' ', '')
                    print(f"Name: {name}\nNumber: {no}")

                    diccy['Name'] = name
                    diccy['Phone Number'] = no

                    diclist.append(diccy)

        print('-'*100)

    deets = pd.DataFrame(diclist)

    return deets

def main(email, subject, output):
    # create connection
    #M = IMAP4_SSL('imap.gmail.com')

    with IMAP4_SSL('imap.gmail.com') as M:

        # login
        try:
            rv, data = M.login(email, getpass.getpass())
        except imaplib.IMAP4.error:
            sys.exit("LOGIN FAILED!!!")

        print(rv, data)

        #irrelevant??
        rv, mailboxes = M.list()
        if rv == 'OK':
            print("Mailboxes:")
            print(mailboxes)

        rv, data = M.select(mailbox, readonly=True)  # readonly = True added
        if rv == 'OK':
            print("Processing mailbox...\n")
            deets = process_mailbox(M, subject)
            deets.to_csv(output, index=False)

            # Open results
            p = subprocess.Popen(['open', 'TREEPETTS.csv'], stdout=subprocess.PIPE,
                                 stderr=subprocess.PIPE)
            stdout, stderr = p.communicate()

            # Print outputs and check for errors
            print(stdout.decode())

            if stderr.decode():
                print(f"{stderr.decode()}\nRUN FAILED")
            else:
                print("RUN SUCCESS")


        else:
            print("ERROR: Unable to open mailbox ", rv)

        M.close()
        M.logout()

    return 0

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Scrapes emails for casting personal info")
    parser.add_argument('-e', dest='email', required=True, help='Email address')
    parser.add_argument('-s', dest='subject', required=True, help='Email subject')
    parser.add_argument('-o', dest='output', required=True, help='Csv output path.')
    args = parser.parse_args()

    out = main(args.email, args.subject, args.output)
    sys.exit(out)