#!/usr/bin/env python3

"""Emails a custom message to a list of actors/actresses in an Excel spreadsheet."""

__author__ = 'Luke Swaby (lds20@ic.ac.uk)'
__version__ = '0.0.1'

## Imports ##
import re
import os
import markdown
import keyring
from pwinput import pwinput
import smtplib
from email.message import EmailMessage
import logging
from datetime import datetime

## Variables ##
providers = {"gmail": "smtp.gmail.com",
             "hotmail": "smtp.live.com",
             "ionos": "smtp.ionos.de",
             "icloud": "smtp.mail.me.com"}

## Functions ##
def yes_no(prompt):
    """
    Function that aks user a yes/no question, returning a boolean value (Yes = True, No = False)
    """
    x = input(prompt).lower()

    while x not in ['y', 'n']:
        x = input("Please enter 'y' (yes) or 'n' (no): ").lower()

    return True if x == 'y' else False


def fetch_password(service, usn):
    """Fetches password from keychain, asks for user input if nothing is there.
    """
    pw = keyring.get_password(service, usn)

    if pw:
        print("Password fetched from keychain.")
    else:
        pw = pwinput(f"{service.title()} Password: ")

    return pw


def password_to_keychain(service, un, pw):
    """Checks if login details are already saved in keychain. If so, does nothing. If not, asks user if they would
    like to save them.
    """

    if keyring.get_password(service, un) == pw:
        pass  # if pw already in keychain, do nothing
    else:
        set_pw = yes_no("Would you like to save these login details to keychain? ('y'/'n'): ")
        if set_pw:
            keyring.set_password(service, un, pw)  # replace login details in keychain
            print("Saved to keychain.")
        else:
            pass  # do nothing

    return


class EmailText():
    """Class that enables conversion of customised input text to either HTML or plain text.
    """

    def __init__(self, text):
        self.text = text

    def convert_to_html(self, sign=False):
        """Converts customised markdown text to html, including an HTML signature if desired.
        """
        out = markdown.markdown(self.text)  # bold and italics

        # TODO: enable coloured titles
        # # H1 -> <h1>H1</h1>
        # [blue]{# H1} -> <h1 style="color:blue;">H1</h1>

        # Convert colours: [green]{...} -> <span style="color: green">...</span>
        for lefttag in re.findall(r'\[\w+\]\{', out):
            col = re.search(r'\[(\w+)\]', lefttag).group(1)  # extract colour string
            lefttag_html = f'<span style="color: {col}">'
            out = out.replace(lefttag, lefttag_html)

        out = out.replace('}', '</span>')

        # Add signature
        if sign:
            with open('signature.txt') as sig:
                signature = sig.read()
            out += ('<br><br>' + signature)

        return out

    def convert_to_plain(self):
        """Converts customised markdown text to plain text.
        """
        left_col_tags = re.findall(r'\[\w+\]\{', self.text)
        rm = ['*', '}', '#', *left_col_tags]

        text = self.text
        for x in rm:
            text = text.replace(x, '')

        return text


class CustomEmailMessage(EmailMessage):
    """Modified EmailMessage() that allows attachments of multiple docs specified in a list of path strings
    """

    def add_attachments(self, doc_list: list):
        """Function to add multiple attachments to an email.
        """
        for doc in doc_list:
            # Open file in binary mode
            with open(doc, "rb") as attachment:
                # Add file as application
                content = attachment.read()
                self.add_attachment(content, maintype='application',
                                    subtype=os.path.splitext(doc)[1],
                                    filename=os.path.basename(doc))

        return


class CustomizedSMPTSession(smtplib.SMTP):
    """Customized SMPT session that includes a function to repeatedly try logging in, giving the user the option for the
    correct password to be saved to the keychain, and a function to send a signed Multipart email with multiple
    attachments.
    """

    # TODO:
    #  1. somehow add an extra attribute, so you can put providers dict in here as CustomizedSMPTSession.providers
    #  2. Create a repeat attempt login function for spotlight scrape

    def repeat_attempt_login(self, service: str, un: str, pw: str, return_creds: bool = False):
        """Function that tries to log in with provided details, re-trying with new ones over and over until logged in
        successfully, saving the new (correct) password to the keychain, and returning correct credentials.
        """
        logged_in = False

        while not logged_in:
            try:
                # Attempt to log in
                self.login(un, pw)
                logged_in = True
            except smtplib.SMTPAuthenticationError:
                # If error, then re-enter details until successful
                print("LOGIN FAILED. Please re-enter details: ")
                un = input(f"Username: ")
                pw = pwinput(f"Password: ")

        print('\nLOGIN SUCCESSFUL')

        password_to_keychain(service, un, pw)

        return un, pw if return_creds else None

    def reconnect(self, from_address, password):
        """Reconnects to server
        """
        host = self._host
        self.connect(host, 587)
        self.ehlo(host)
        self.repeat_attempt_login(host, from_address, password)

        return

    def send_email(self, msg_template: str, subject: str,
                  from_address: str, to_address: str, names: str,
                  docs_to_add: list = None, sign: bool = False, ghost: bool = False):
        """Send email (without logging)
        """
        content = msg_template.replace('$N', names)  # Insert name into email message
        content = EmailText(content)

        msg = CustomEmailMessage()
        msg['Subject'] = subject
        msg['From'] = from_address
        msg['To'] = to_address

        msg.set_content(content.convert_to_plain())
        msg.add_alternative(content.convert_to_html(sign), subtype='html')

        if docs_to_add:
            msg.add_attachments(docs_to_add)  # attach documents if any

        self.send_message(msg)

        # Logging
        if not ghost:
            logging.basicConfig(filename="../email.log", level=logging.INFO)
            now = datetime.now()
            dt_string = now.strftime("%d/%m/%Y %H:%M")  # fetch date and time
            logging.info(f"{dt_string}: Email sent from {from_address} to {to_address} (Subject: '{subject}'.")

        return

    def preview_email(self, template_path: str, from_address: str, password: str, subject: str,
                      docs_to_add: list = None, sign: bool = False):
        """Sends email to self, allowing corrections until user satisfied
        """
        # Check user is ok with email format

        with open(template_path) as email:
            msg_template = email.read()  # Read in template

        while True:
            # Handle any timeout errors
            try:
                self.send_email(msg_template, subject, from_address, from_address, '$NAMES', docs_to_add, sign,
                                ghost=True)
            except smtplib.SMTPSenderRefused:
                # If session times out then re-create it
                self.reconnect(from_address, password)
                self.send_email(msg_template, subject, from_address, from_address, '$NAMES', docs_to_add, sign,
                                ghost=True)

            email_ok_prompt = f">> A formatted email has been sent to {from_address} for you to inspect. " \
                              f"Are you happy to forward this to agencies? ('y'/'n'): "
            email_ok = yes_no(email_ok_prompt)

            if email_ok:
                return
            else:
                input(f">> Please edit template at path '{template_path}'. "
                      f"Hit ENTER to re-preview when you have saved new contents.")
                with open(template_path) as email:
                    msg_template = email.read()


