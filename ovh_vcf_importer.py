#!/usr/bin/python
# -*- coding: utf-*-
import argparse
import configparser
import os
import re
import sys

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.ie.webdriver import WebDriver
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions

# Configuration keys
URL_CONFIGURATION_PROPERTY: str = 'url'
EMAIL_CONFIGURATION_PROPERTY: str = 'mail'
PASSWORD_CONFIGURATION_PROPERTY: str = 'pass'
VCF_FILE_CONFIGURATION_PROPERTY: str = 'vcf_file'
# Labels keys
LOGOUT_LABEL: str  = 'logout'
NEW_LABEL: str = 'new'
FIRSTNAME_LABEL: str = 'firstname'
LASTNAME_LABEL: str = 'lastname'
NAME_LABEL: str = 'name'
EMAIL_ADDRESS_LABEL: str = 'email_address'
PHONE_NUMBER_LABEL: str = 'phone_number'
WORK_LABEL: str = 'work'
CANCEL_LABEL: str = 'cancel'
IGNORE_LABEL: str = 'ignore'
SAVE_LABEL: str = 'save'

# Global variables
account_configuration_dict: dict = {}
labels_dict: dict = {}
dry_run: bool = False

class Contact(object):
    def __init__(self):
        # noinspection PyTypeChecker
        self.name: str = None
        # noinspection PyTypeChecker
        self.full_name: str = None
        # noinspection PyTypeChecker
        self.email: str = None
        # noinspection PyTypeChecker
        self.nickname: str = None
        # noinspection PyTypeChecker
        self.birthday: str = None
        # noinspection PyTypeChecker
        self.anniversary: str = None
        # noinspection PyTypeChecker
        self.gender: str = None
        # noinspection PyTypeChecker
        self.phone_number: str = None

    def is_valid(self) -> bool:
        return ((self.name is not None and len(self.name) > 0) or (self.full_name is not None and len(self.full_name) > 0)) and (self.email is not None and len(self.email) > 0)

    def __str__(self):
        return f'Name: {self.name}, Full Name: {self.full_name}, Email: {self.email}, Phone: {self.phone_number}, Nickname: {self.nickname}, Birthday: {self.birthday}, Anniversary: {self.anniversary}, Gender: {self.gender}'


def add_contact(driver: WebDriver, contact: Contact) -> None:
    global labels_dict
    """ Add contact using Selenium automation """
    print("Processing contact: " + str(contact))
    elt = find_element(driver, By.XPATH, "//button[@title='%s']/span[2]" % labels_dict[NEW_LABEL])
    elt.click()
    popup_elt = find_element(driver, By.CLASS_NAME, 'popupPanel')
    # noinspection PyTypeChecker
    firstname: str = None
    # noinspection PyTypeChecker
    lastname: str = None
    if contact.full_name is not None and len(contact.full_name) > 0:
        if ' ' in contact.full_name:
            firstname = contact.full_name.split(' ')[0]
            lastname = contact.full_name.split(' ')[1]
        else:
            lastname = contact.full_name
    if contact.name is not None and len(contact.name) > 0:
        lastname = contact.name
    # firstname
    if firstname is not None and len(firstname) > 0:
        firstname = firstname.lower().capitalize()
        elt = find_element_in_parent(driver, popup_elt, By.XPATH, "//input[@aria-label='%s']" % labels_dict[FIRSTNAME_LABEL])
        elt.click()
        print("Using firstname: " + firstname)
        elt.send_keys(firstname)
    # lastname
    if lastname is not None and len(lastname) > 0:
        elt = find_element_in_parent(driver, popup_elt, By.XPATH, "//input[@role='textbox' and @aria-label='%s']" % labels_dict[LASTNAME_LABEL])
        elt.click()
        print("Using lastname: " + lastname)
        elt.send_keys(lastname)
    # email
    if contact.email is not None and len(contact.email) > 0:
        elt = find_element_in_parent(driver, popup_elt, By.XPATH, "//input[@role='textbox' and @aria-label='%s']" % labels_dict[EMAIL_ADDRESS_LABEL])
        elt.click()
        print("Using email: " + contact.email)
        elt.send_keys(contact.email)
    # phone number
    if contact.phone_number is not None and len(contact.phone_number) > 0:
        parent_elt = find_element_in_parent(driver, popup_elt, By.XPATH, "//span[text()='%s']" % labels_dict[PHONE_NUMBER_LABEL])
        elt = find_element_in_parent(driver, parent_elt, By.XPATH, "//input[@role='textbox' and @aria-label='%s']" % labels_dict[WORK_LABEL])
        elt.click()
        print("Using phone number: " + contact.phone_number)
        elt.send_keys(firstname)
    if dry_run:
        elt = find_element(driver, By.XPATH, "(//span[text()='%s'])[2]" % labels_dict[CANCEL_LABEL])
        elt.click()
        elt = find_element(driver, By.XPATH, "//span[text()='%s']" % labels_dict[IGNORE_LABEL])
        elt.click()
    else:
        elt = find_element(driver, By.XPATH, "//span[text()='%s']" % labels_dict[SAVE_LABEL])
        elt.click()

def find_element(driver: WebDriver, by: By, value: str) -> WebElement:
    print("Searching element by: %s, value: %s" % (by, value))
    return WebDriverWait(driver, 10).until(
        expected_conditions.visibility_of_element_located((by.__str__(), value))
    )


# noinspection PyUnusedLocal
def find_element_in_parent(driver: WebDriver, parent: WebElement, by: By, value: str) -> WebElement:
    print("Searching element by: %s, value: %s" % (by, value))
    return parent.find_element(by, value)


def get_value(line: str) -> str:
    result: str = line.split(':')[1]
    if ';' in result:
        result = result.split(';')[0]
    return  result.strip()


def load_labels(driver: WebDriver) -> None:
    global labels_dict
    language: str = 'fr-FR'
    elt = find_element(driver, By.TAG_NAME, 'html')
    attr = elt.get_attribute('lang')
    if attr:
        language = attr
    configuration_parser = configparser.RawConfigParser()
    configuration_parser.read('labels.conf')
    print('Loading labels for language: %s' % language)
    labels_dict.update(dict(configuration_parser.items(language)))


def logout(driver: WebDriver) -> None:
    global labels_dict
    # noinspection BroadException
    try:
        top_menu_elt = find_element(driver, By.CSS_SELECTOR, '#O365_TopMenu > div > div > div:nth-child(1) > div:nth-child(13) > button')
        top_menu_elt.click()
        elt = find_element(driver, By.XPATH, "//span[text()='%s']" % labels_dict[LOGOUT_LABEL])
        elt.click()
    except Exception as e:
        print(f"Error while logging out {e=}, {type(e)=}")


def main():
    global dry_run
    parser = argparse.ArgumentParser(prog='ovh_vcf_importer', description='OVH Emails Pro VCF import')
    parser.add_argument('-d', action='store_true', help="Dry run flag")
    parser.add_argument('--conf', required=True, help='The configuration file')
    parser.add_argument('--account', required=False, default="account", help='The account section to use (specified in the configuration file, default is \'account\')')
    args = parser.parse_args()
    dry_run = args.d
    if args.conf:
        configuration_parser = configparser.RawConfigParser()
        configuration_parser.read(args.conf)
        configuration_dict = dict(configuration_parser.items('global'))
        if not URL_CONFIGURATION_PROPERTY in configuration_dict:
            print('URL not specified in configuration (missing property: ' + URL_CONFIGURATION_PROPERTY + ')')
            sys.exit(1)
        url: str = configuration_dict[URL_CONFIGURATION_PROPERTY]
        account_section = 'account'
        if args.account:
            account_section = args.account
        account_configuration_dict.update(dict(configuration_parser.items(account_section)))
        if not EMAIL_CONFIGURATION_PROPERTY in account_configuration_dict:
            print('No mail account specified in configuration (missing property: ' + EMAIL_CONFIGURATION_PROPERTY + ')')
            sys.exit(1)
        account: str = account_configuration_dict[EMAIL_CONFIGURATION_PROPERTY]
        if not PASSWORD_CONFIGURATION_PROPERTY in account_configuration_dict:
            print('No password specified in configuration for the mail account: ' + account + ' (missing property: ' + PASSWORD_CONFIGURATION_PROPERTY + ')')
            sys.exit(1)
        password: str = account_configuration_dict[PASSWORD_CONFIGURATION_PROPERTY]
        if not VCF_FILE_CONFIGURATION_PROPERTY in account_configuration_dict:
            print('VCard file not specified in configuration (missing property: ' + VCF_FILE_CONFIGURATION_PROPERTY + ')')
            sys.exit(1)
        vcf_file: str = account_configuration_dict[VCF_FILE_CONFIGURATION_PROPERTY]
        if not os.path.exists(vcf_file):
            print('VCard file specified in configuration does not exists: ' + vcf_file)
            sys.exit(1)
        count: int = 0
        driver: WebDriver = webdriver.Chrome()
        try:
            print('Opening webmail: ' + url)
            driver.get(url)
            driver.maximize_window()
            driver.set_page_load_timeout(15)
            driver.set_script_timeout(15)
            driver.implicitly_wait(15)
            print('Authenticating on webmail using: ' + account)
            elt = find_element(driver, By.ID, 'username')
            elt.send_keys(account)
            elt = find_element(driver, By.ID, 'password')
            elt.send_keys(password)
            elt.send_keys(Keys.RETURN)
            load_labels(driver)
            elt = find_element(driver, By.ID, 'O365_MainLink_NavMenu')
            elt.click()
            elt = find_element(driver, By.ID, 'O365_AppTile_ShellPeople')
            elt.click()
            print('Processing VCard file: ' + vcf_file)
            with open(vcf_file, 'r', encoding='UTF-8') as file:
                # noinspection PyTypeChecker
                contact: Contact = None
                while line := file.readline():
                    if line.startswith('BEGIN:VCARD'):
                        contact = Contact()
                    elif contact is not None:
                        if line.startswith('END:VCARD'):
                            if contact.is_valid():
                                add_contact(driver, contact)
                                count += 1
                            else:
                                print("Contact is not valid: " + str(contact))
                        elif line.startswith('N:'):
                            contact.name = get_value(line)
                        elif line.startswith('FN:'):
                            contact.full_name = get_value(line)
                        elif re.match("^EMAIL[^:]*:", line):
                            contact.email = get_value(line)
        except Exception as e:
            print(f"Error during processing {e=}, {type(e)=}")
        finally:
            logout(driver)
            driver.quit()
            print('Completed, %s contacts processed.' % str(count))
    else:
        print('No configuration specified.')
        sys.exit(1)


if __name__ == '__main__':
    main()