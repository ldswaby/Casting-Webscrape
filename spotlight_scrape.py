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
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException

## Functions ##
def main(webpage="https://www.spotlight.com/shortlists/2737534",
         usn_field="Username",
         usn='TreePetts',
         pwd_field="Password",
         pwd='kitchen150%',
         outfile='Spotlight_spreadsheet.xlsx'):
    """
    Function that scrapes starnow site for names etc
    """

    driver = webdriver.Chrome(executable_path='./chromedriver')

    driver.get(webpage)

    # if already logged in, hash these lines out
    username = driver.find_element_by_id(usn_field)
    password = driver.find_element_by_id(pwd_field)
    username.send_keys(usn)
    password.send_keys(pwd)
    driver.find_element_by_id("sign-in-button").click()
    #############################################

    delay = 10  # seconds
    try:
        myElem = WebDriverWait(driver, delay).until(
            EC.presence_of_element_located((By.ID, "Radio-Signal")))
        time.sleep(5)
        print("Page is ready!")
    except TimeoutException:
        print("Loading took too much time!")

    txt = driver.page_source

    driver.close()

    pattern = r'alt="(.+?)"[\S\n\t\v ]+?"c-agency__card-agency-name">(.+?)<' \
              r'[\S\n\t\v ]+?tel\:\/\/(.+?)"[\S\n\t\v ]+?"mailto:(.+?)"'

    regex = re.compile(pattern)

    rows = [(match.group(1).strip(), match.group(2).strip(),
             match.group(3).strip(), match.group(4).strip())
            for match in regex.finditer(txt)]

    print('Saving dataframe...')

    # Building and saving dataframe
    out = pd.DataFrame(rows, columns=['NAME', 'AGENT', 'CONTACT NUMBER', 'EMAIL'])
    out['CONTACT NUMBER'] = out['CONTACT NUMBER'].astype("str")
    out.to_excel(outfile, index=None, header=True)

    print('Done!')

    return 0

if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description="Script for scraping Starnow webpage for names, "
                    "nationalities, and contact numbers.")

    parser.add_argument('-w', default="https://www.spotlight.com/shortlists/2723718",
                        dest='webpage',
                        help='The URL of the webpage you wish to scrape')
    parser.add_argument('-uf', default="Username",
                        dest='usn_field',
                        help='The HTML field for username login.')
    parser.add_argument('-un', required=True, type=str, dest='usn',
                        help='Username')
    parser.add_argument('-pf', default="Password",
                        dest='pwd_field',
                        help='The HTML field for password login.')
    parser.add_argument('-pw', required=True, type=str, dest='pwd',
                        help='Password')
    parser.add_argument('-o', default='Spotlight_spreadsheet.xlsx',
                        dest='outfile', help='Out file path')
    parser.add_argument('-op', default=True, type=bool,
                        dest='open', help='Open the file on completion?')

    args = parser.parse_args()

    # Ru program
    status = main(webpage=args.webpage,
                  usn_field=args.usn_field,
                  usn=args.usn,
                  pwd_field=args.pwd_field,
                  pwd=args.pwd,
                  outfile=args.outfile)

    if args.open:
        subprocess.call(['open', args.outfile])  # Mac

    sys.exit(status)