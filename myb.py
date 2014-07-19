import argparse
import json
import os
import shutil
import smtplib
import string
import sys
import datetime
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
import time

LOGIN_URL = 'https://online.colesmastercard.com.au/access/login'
JSON_FILE = 'transactions.json'

assert os.path.exists('config.py'), 'Does config.py exist? Possibly use config_example.py as template for new config'
from config import config

previous_history = {}


class TooManyTransactionsError(Exception):
    pass


class Transaction(object):
    identity = string.maketrans('', '')

    @classmethod
    def collect_pending_transactions(cls, driver):

        fields = {'block_xpath': '//div[@name="pendingTransactionsTag"]',
                  'date': 'Pending_transactionDate',
                  'name': 'Pending_cardName',
                  'description': 'Pending_transactionDescription',
                  'amount': 'Pending_transactionAmount'}

        return cls.collect_transactions(driver, fields)

    @classmethod
    def collect_processed_transactions(cls, driver, last_transaction=None):

        max_transactions = 50

        fields = {'block_xpath': '//div[@name="transactionHistoryPage"]',
                  'date': 'Transaction_TransactionDate',
                  'name': 'Transaction_CardName',
                  'description': 'Transaction_TransactionDescription',
                  'amount': 'Transaction_Amount'}

        return cls.collect_transactions(driver, fields, last_transaction, max_transactions=max_transactions)

    @classmethod
    def collect_transactions(cls, driver, fields, last_transaction=None, max_transactions=50):

        transactions = []

        while True:
            block_element = driver.find_element_by_xpath(fields['block_xpath'])
            transactiontable_elements = block_element.find_elements_by_name('DataContainer')

            for transactiontable_element in transactiontable_elements:

                transaction = Transaction({'date': transactiontable_element.find_element_by_name(fields['date']).text,
                                           'name': transactiontable_element.find_element_by_name(fields['name']).text,
                                           'description': transactiontable_element.find_element_by_name(
                                               fields['description']).text,
                                           'amount': transactiontable_element.find_element_by_name(
                                               fields['amount']).text})

                transactions.append(transaction)

                if last_transaction is not None and last_transaction == transaction:
                    return transactions

                if len(transactions) >= max_transactions:
                    sys.stderr.write('*** WARNING ***: Max Transactions (%s) retrieved.\n' % max_transactions)
                    return transactions

            try:
                nextbutton_element = block_element.find_element_by_name('nextButton')
                nextbutton_element.click()

            except NoSuchElementException:
                break

        return transactions

    @classmethod
    def currency_to_float(cls, value):
        return float(str(value).translate(cls.identity, '$,'))

    def __init__(self, data=None):
        self.ignore_transaction = False
        self._dict = data if data is not None else {}

    @property
    def raw_data(self):
        return self._dict

    @property
    def date(self):
        return self._dict.get('date', '<unknown date>')

    @property
    def name(self):
        return self._dict.get('name', '<unknown name>')

    @property
    def description(self):
        return self._dict.get('description', '<unknown description>')

    @property
    def amount(self):
        return self.currency_to_float(self._dict.get('amount', '0.00'))

    def __repr__(self):
        return '(%s, %s, %s, %s)' % (self.date, self.name, self.description, currency(self.amount))

    def __hash__(self):
        return (self.date, self.name, self.description, self.amount).__hash__()

    def __eq__(self, other):
        assert type(other) is Transaction

        return self.__hash__() == other.__hash__()

    def __ne__(self, other):
        return not self.__eq__(other)


def currency(value):
    if value < 0:
        neg = True
    else:
        neg = False

    result = '%s$%.2f' % ('-' if neg is True else '', abs(value))

    return result


def login(driver, username, password):
    driver.get(LOGIN_URL)

    username_element = driver.find_element_by_id('AccessToken_Username')
    username_element.send_keys(username)

    password_element = driver.find_element_by_id('AccessToken_Password')
    password_element.send_keys(password)

    password_element.submit()

    WebDriverWait(driver, 60).until(EC.text_to_be_present_in_element((By.NAME, 'weclomeMessage'), 'Welcome back'))


def logout(driver):
    # Logout
    logout_element = driver.find_element_by_id('logout')
    logoutlink_element = logout_element.find_element_by_link_text('Log Out')
    logoutlink_element.click()

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, 'cardsonline-access-logout')))


def process_pending(driver):

    existing_pending = []
    for transaction_dict in previous_history.get('pending', []):
        existing_pending.append(Transaction(transaction_dict))

    current_transactions = Transaction.collect_pending_transactions(driver)

    new_transactions = []
    for current_transaction in current_transactions:
        if current_transaction not in existing_pending:  # A new transaction has been added
            new_transactions.append(current_transaction)

    for existing_pending_transaction in existing_pending:
        if existing_pending_transaction not in current_transactions:  # A pending transaction has been processed
            existing_pending_transaction.raw_data['description'] += ' - PROCESSED'
            existing_pending_transaction.ignore_transaction = True
            new_transactions.append(existing_pending_transaction)

    previous_history['pending'] = [x.raw_data for x in current_transactions]

    return new_transactions


def process_transactions(driver):

    existing_processed = []
    for transaction_dict in previous_history.get('processed', ''):
        existing_processed.append(Transaction(transaction_dict))

    last_transaction = existing_processed[0] if len(existing_processed) > 0 else None

    current_transactions = Transaction.collect_processed_transactions(driver, last_transaction)

    new_transactions = []
    for current_transaction in current_transactions:
        if current_transaction not in existing_processed:
            new_transactions.append(current_transaction)

    previous_history['processed'] = [x.raw_data for x in current_transactions]

    return new_transactions


def notify_user(args, availablecredit, creditlimit, currentbalance, new_pending_transactions, new_transactions):
    msg = ['Subject: Credit Card usage report\n',
           'Email report for Credit Card usage.',
           'Current balance: %s' % currency(currentbalance),
           'Available credit: %s' % currency(availablecredit),
           '']

    if len(new_pending_transactions) > 0:
        msg.append('Pending Transactions:')

        for transaction in new_pending_transactions:
            msg.append('%s\t%s\t\t%s\t\t"%s"' % (transaction.date, currency(transaction.amount), transaction.name,
                                                 transaction.description))

        transactions_total = sum([x.amount for x in new_pending_transactions if x.ignore_transaction is False])

        msg.extend(['',
                    'Pending transactions total: %s' % currency(transactions_total)])

        overlimit = -1 * (creditlimit - (currentbalance - transactions_total))
        if overlimit > 0:
            msg.append('Over limit by: %s' % currency(overlimit))

        msg.extend(['', ''])

    if len(new_transactions) > 0:
        msg.append('New Transactions:')

        for transaction in new_transactions:
            msg.append('%s\t%s\t\t%s\t\t"%s"' % (transaction.date, currency(transaction.amount), transaction.name,
                                                 transaction.description))

        transactions_total = sum([x.amount for x in new_transactions])

        msg.extend(['',
                    'New transactions total: %s' % currency(transactions_total)])

        msg.extend(['', ''])

    server = smtplib.SMTP('%s' % config['server'])

    if config['secure'] is True:
        server.starttls()

    server.login(args.email_username, args.email_password)
    server.sendmail(config['fromaddr'], config['toaddrs'], '\n'.join(msg))
    server.quit()


def main(argv):
    parser = argparse.ArgumentParser(description='Automatically retrieve account details from Coles Mastercard')
    parser.add_argument('-u', '--username', required=True, help='Username of account to retrieve')
    parser.add_argument('-p', '--password', required=True, help='Password of account')
    parser.add_argument('-eu', '--email-username', required=True, help='Username of email account to send alerts to')
    parser.add_argument('-ep', '--email-password', required=True, help='Password of email user')

    args = parser.parse_args(argv)

    username = args.username
    password = args.password

    sleep_time = 15

    while True:

        print('Processing started: %s' % datetime.datetime.now())
        driver = webdriver.Chrome()

        try:

            if os.path.exists(JSON_FILE):
                if os.path.exists('%s.bak' % JSON_FILE):
                    os.unlink('%s.bak' % JSON_FILE)

                shutil.copy2(JSON_FILE, '%s.bak' % JSON_FILE)

                with open(JSON_FILE, 'rb') as inf:
                    previous_history.update(json.load(inf))

            login(driver, username, password)

            # Current Balance
            previous_currentbalance = previous_history.get('current_balance', 0.00)
            currentbalance_element = driver.find_element_by_name('AccountSummary_CurrentBalanceAmount')
            currentbalance = Transaction.currency_to_float(currentbalance_element.text)

            #Credit Limit
            creditlimit_element = driver.find_element_by_name('AccountSummary_CreditLimitAmount')
            creditlimit = Transaction.currency_to_float(creditlimit_element.text)

            # Available Credit
            previous_availablecredit = previous_history.get('available_credit', creditlimit)
            availablecredit_element = driver.find_element_by_name('AccountSummary_AvailableCreditAmount')
            availablecredit = Transaction.currency_to_float(availablecredit_element.text)

            #All transactions since last transaction
            # My account
            myaccount_element = driver.find_element_by_link_text('My Account')
            myaccount_element.click()

            # Transactions
            transactions_element = driver.find_element_by_link_text('Transactions')
            transactions_element.click()

            new_pending_transactions = process_pending(driver)
            new_transactions = process_transactions(driver)

            balance_diff = int(currentbalance * 100) - int(previous_currentbalance * 100)
            credit_diff = int(availablecredit * 100) - int(previous_availablecredit * 100)

            if balance_diff != 0 or credit_diff != 0 or len(new_pending_transactions) > 0 or len(new_transactions) > 0:
                sleep_time = 15
                print('Current Balance: %s' % currentbalance)
                print('Available Credit: %s' % availablecredit)
                print('Credit Limit: %s' % creditlimit)

                previous_history['available_credit'] = availablecredit
                previous_history['current_balance'] = currentbalance

                try:
                    notify_user(args, availablecredit, creditlimit, currentbalance, new_pending_transactions,
                                new_transactions)
                    print('Report emailed to %s' % config['toaddrs'])

                except Exception as e:
                    sys.stderr.write('Failed to notiy user of account activity: %s (%s)' % (type(e), str(e)))
                    sys.stderr.flush()
                    print('')

                with open(JSON_FILE, 'wb') as outf:
                    json.dump(previous_history, outf, sort_keys=True, indent=4, separators=(',', ': '))

            else:
                print('No account activity - Nothing to report - will try again in %s mins' % sleep_time)

            logout(driver)

        except Exception as e:
            if os.path.exists('%s.bak' % JSON_FILE):
                if os.path.exists(JSON_FILE):
                    os.unlink(JSON_FILE)

                shutil.copy2('%s.bak' % JSON_FILE, JSON_FILE)
                os.unlink('%s.bak' % JSON_FILE)

            sys.stderr.write('Failed process new account activity activity: %s (%s)\n' % (type(e), str(e)))
            sys.stderr.flush()
            print('')

        finally:
            driver.quit()
            print('Processing completed: %s' % datetime.datetime.now())
            time.sleep(sleep_time * 60)  # Sleep for minutes before trying again

            if sleep_time < 60:
                sleep_time += 5


if __name__ == '__main__':
    main(sys.argv[1:])

