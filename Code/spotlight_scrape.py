#!/usr/bin/env python3

"""Script for scraping Spotlight shortlist for actor/actress details (names, agents, contact numbers, and email
addresses) and loading them into an Excel spreadsheet."""

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
import tkinter as tk
from tkinter import filedialog
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.chrome.service import Service
from core import yes_no, fetch_password

## Functions ##

def parse_args():
    """
    Parses arguments from the command line.
    """
    parser = argparse.ArgumentParser(description="Script for scraping Spotlight shortlist for actor/actress details "
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

    args = parser.parse_args()

    # More args
    print('\nPLEASE FILL THE FOLLOWING:\n')
    usn = input("Spotlight Username: ")
    pwd = fetch_password("Spotlight", usn)  # Obtain keyring from keychain. Set it if absent
    webpage = input("Spotlight shortlist URL: ")

    outdir_prompt = 'Hit ENTER to select desired output folder: '
    input(outdir_prompt)
    root = tk.Tk()  # Initialise dialog box
    root.withdraw()
    outdir = filedialog.askdirectory()  # fetch directory path
    print(' ' * len(outdir_prompt) + '\033[A' + outdir)  # print at end of previous line
    root.destroy()  # delete dialog window

    if not outdir.endswith('/'):
        outdir += '/'

    outfile = input("Desired output file name: ")
    outpath = outdir + outfile

    openfile = yes_no("Would you like to open the file on completion? ('y'/'n'): ")

    return webpage, args.usn_field, usn, args.pwd_field, pwd, outpath, openfile

def main(webpage, usn_field, usn, pwd_field, pwd, outfile, open=False):
    """
    Function that scrapes starnow site for names etc
    """
    # Format inputs
    if not outfile.endswith('.xlsx'):
        outfile += '.xlsx'
    #if not outfile.startswith('../Data/'):
    #    outfile = '../Data/' + outfile

    s = Service('./chromedriver')
    driver = webdriver.Chrome(service=s)

    print('\nLoading webpage...')
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

    print('Scraping data...')
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

    out = pd.DataFrame(rows, columns=['NAME', 'AGENT', 'CONTACT NUMBER', 'EMAIL'])  # Building dataframe

    # Format cols
    out['CONTACT NUMBER'] = out['CONTACT NUMBER'].astype("str")
    out['CONTACT NUMBER'] = out['CONTACT NUMBER'].str.replace(r'\D+', '', regex=True)  # keep numericals only
    out['NAME'] = out['NAME'].str.title()
    out['CONTACT?'] = None

    print(f'Saving data to {outfile}...')
    out.to_excel(outfile, index=None, header=True)

    print('Done!')

    if open:
        print('Opening...')
        subprocess.call(['open', outfile])  # Mac

    return

if __name__ == '__main__':
    main(*parse_args())