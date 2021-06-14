#!/usr/bin/env python3

"""Scraping Starnow webpage for names, nationalities etc."""

__author__ = 'Luke Swaby (lds20@ic.ac.uk)'
__version__ = '0.0.1'

## Imports ##
import mechanize
import http.cookiejar
import re
import pandas as pd
import argparse
import sys
import subprocess

webpage="https://www.starnow.co.uk/casting/1123833/applicants"
usn_field="ctl00$cphMain$signinForm$email"
usn='casting@treepetts.co.uk'
pwd_field="ctl00$cphMain$signinForm$password"
pwd='kitchen69'
outfile='GOAL NATIONAL TEAM FOOTBALL FANS.xlsx'

## Functions ##
def main(webpage="https://www.starnow.co.uk/casting/1123833/applicants",
         pages=4,
         usn_field="ctl00$cphMain$signinForm$email",
         usn='casting@treepetts.co.uk',
         pwd_field="ctl00$cphMain$signinForm$password",
         pwd='kitchen69',
         outfile='GOAL NATIONAL TEAM FOOTBALL FANS.xlsx'):
    """
    Function that scrapes starnow site for names etc
    """

    rows = []
    for x in range(1, (pages+1)):
        print(f'Scraping p{x}...')

        cj = http.cookiejar.CookieJar()
        br = mechanize.Browser()
        br.set_handle_robots(False)
        br.set_cookiejar(cj)

        br.open(f"{webpage}?p={x}")

        br.select_form(nr=0)

        # Set username and password.
        # Note that the key names of this dict will differ depending on the
        # site. to find them, go to the login page and right click the fields
        # that need flling, and see corresponding the value of 'name' in
        # the html.
        br.form[usn_field] = usn
        br.form[pwd_field] = pwd

        # Login
        br.submit()

        # Grab all text from page
        txt = str(br.response().read())

        # Specify regex pattern
        pattern = r"\"What Country Do You Represent\?\",\"answer\"\:\"(.*?)\"" \
                  r".*?\"fullName\":\"(.*?)\".*?\"phoneNumber\"\:\"(.*?)\""
        regex = re.compile(pattern)

        # Compile into list
        deets = [(match.group(2).strip(), match.group(1),
                  match.group(3).strip()) for match in regex.finditer(txt)]

        # Merge with list of rows
        rows += deets


    print('Saving dataframe...')

    # Building and saving dataframe
    out = pd.DataFrame(rows, columns=['NAME', 'NATIONAL TEAM', 'CONTACT NUMBER'])
    out['CONTACT NUMBER'] = out['CONTACT NUMBER'].astype("str")
    out['EMAIL'] = 'STARNOW'
    out['Asked to self tape?'] = out['Received self tape?'] = out['AGE'] = None
    out = out[['NAME', 'EMAIL', 'CONTACT NUMBER', 'AGE', 'NATIONAL TEAM', 'Asked to self tape?', 'Received self tape?']]
    out.to_excel(outfile, index=None, header=True)

    # Opening file
    subprocess.call(['open', outfile])  # Mac

    print('Done!')

    return 0

if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description="Script for scraping Starnow webpage for names, "
                    "nationalities, and contact numbers.")

    parser.add_argument('-w', default="https://www.starnow.co.uk/casting/1123833/applicants",
                        dest='webpage',
                        help='The URL of the webpage you wish to scrape')
    parser.add_argument('-p', type=int, required=True, dest='pages',
                        help='The number of pages you wish to scrape')
    parser.add_argument('-uf', default="ctl00$cphMain$signinForm$email",
                        dest='usn_field',
                        help='The HTML field for username login.')
    parser.add_argument('-un', default='casting@treepetts.co.uk', dest='usn',
                        help='Username')
    parser.add_argument('-pf', default="ctl00$cphMain$signinForm$password",
                        dest='pwd_field',
                        help='The HTML field for password login.')
    parser.add_argument('-pw', default='kitchen69', dest='pwd',
                        help='Password')
    parser.add_argument('-o', default='GOAL NATIONAL TEAM FOOTBALL FANS.xlsx',
                        dest='outfile', help='Out file path')

    args = parser.parse_args()

    # Ru program
    status = main(webpage=args.webpage,
                  pages=args.pages,
                  usn_field=args.usn_field,
                  usn=args.usn,
                  pwd_field=args.pwd_field,
                  pwd=args.pwd,
                  outfile=args.outfile)

    sys.exit(status)