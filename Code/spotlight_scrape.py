#!/usr/bin/env python3

"""Scraping Starnow webpage for names, nationalities etc."""

__author__ = 'Luke Swaby (lds20@ic.ac.uk)'
__version__ = '0.0.1'

## Imports ##
import re
import pandas as pd
import subprocess
import argparse
import sys
import time
from pwinput import pwinput
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.chrome.service import Service

## Functions ##

def parse_args():
    """
    Parses arguments from the command line.
    """
    parser = argparse.ArgumentParser(description="Script for scraping Spotlight webpage for actor/actress details "
                                                 "(names, agents, contact numbers, and email addresses) and loading "
                                                 "them into an Excel spreadsheet.")
    parser.add_argument('-uf', default="Username",
                        dest='usn_field',
                        help='The HTML field for username login. This is unlikely to change any time soon, so leaving '
                             'the default is fine.')
    parser.add_argument('-pf', default="Password",
                        dest='pwd_field',
                        help='The HTML field for password login. This is unlikely to change any time soon, so leaving '
                             'the default is fine.')
    parser.add_argument('--open', dest='open', action='store_true', help='Include this flag to open the file on completion.')
    parser.set_defaults(open=False)

    args = parser.parse_args()

    usn = input("Spotlight Username: ")
    pwd = pwinput("Spotlight Password: ")
    webpage = input("Spotlight URL: ")
    outfile = input("Desired file name for output spreadsheet: ")

    return webpage, args.usn_field, usn, args.pwd_field, pwd, outfile, args.open

def main(webpage, usn_field, usn, pwd_field, pwd, outfile, open=False):
    """
    Function that scrapes starnow site for names etc
    """
    # Format inputs
    if not outfile.endswith('.xlsx'):
        outfile += '.xlsx'
    if not outfile.startswith('../Data/'):
        outfile = '../Data/' + outfile

    s = Service('./chromedriver')
    driver = webdriver.Chrome(service=s)

    driver.get(webpage)  # load webpage
    driver.implicitly_wait(3)

    # Login details
    username = driver.find_element(By.ID, usn_field)
    password = driver.find_element(By.ID, pwd_field)
    username.send_keys(usn)
    password.send_keys(pwd)

    driver.find_element(By.ID, "sign-in-button").click()  # Sign in

    delay = 10  # seconds
    try:
        myElem = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.ID, "Radio-Signal")))
        time.sleep(5)
    except TimeoutException:
        driver.close()
        sys.exit("Login Failed. Please check username and password and/or internet connection.")

    txt = driver.page_source

    # RETRIEVE HTML FROM ALL PAGES
    while driver.find_elements(By.CLASS_NAME, "c-pagination-control__arrow-icon.icon-chevronright"):

        driver.find_element(By.CLASS_NAME, "c-pagination-control__arrow-icon.icon-chevronright").click()

        try:
            # wait to load
            myElem = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.ID, "Radio-Signal")))
            time.sleep(3)
        except TimeoutException:
            print("Loading next page took too much time. Page skipped.")

        # append source to txt
        txt += driver.page_source

    driver.close()  # close driver

    pattern = r'alt="(.+?)"[\S\n\t\v ]+?"c-agency__card-agency-name">(.+?)<' \
              r'[\S\n\t\v ]+?tel\:\/\/(.+?)"[\S\n\t\v ]+?"mailto:(.+?)"'

    regex = re.compile(pattern)

    rows = [(match.group(1).strip(), match.group(2).strip(),
             match.group(3).strip(), match.group(4).strip())
            for match in regex.finditer(txt)]

    print(f'Saving dataframe to {outfile}...')

    # Building and saving dataframe
    out = pd.DataFrame(rows, columns=['NAME', 'AGENT', 'CONTACT NUMBER', 'EMAIL'])

    out['CONTACT NUMBER'] = out['CONTACT NUMBER'].astype("str")
    out['CONTACT NUMBER'] = out['CONTACT NUMBER'].str.replace(r'\D+', '', regex=True)  # keep numericals only
    out['CONTACT?'] = None

    out.to_excel(outfile, index=None, header=True)

    print('Done!')

    if open:
        print('Opening...')
        subprocess.call(['open', outfile])  # Mac

    return

if __name__ == '__main__':
    main(*parse_args())