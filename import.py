from credentials import *
#from mint.tags import *
#from mint.utils import *
from mint.api import Mint
import pprint
import xlrd
import sys, traceback
from time import sleep
from datetime import datetime
import time
from selenium import webdriver
from selenium.webdriver.common import action_chains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException, ElementNotVisibleException


class MintAdder(object):

    def __init__(self, file, default_worksheet, default_category, date_col, description_col, debit_col, credit_col, header_row, token, default_merchant):
        self.driver = webdriver.Chrome()
        self.file=file
        self.default_worksheet = default_worksheet
        self.default_category=default_category
        self.date_col=date_col
        self.description_col=description_col
        self.debit_col=debit_col
        self.credit_col=credit_col
        self.header_row=header_row
        self.token=token
        self.default_merchant=default_merchant
        self.import_list=[]
        self.header_vals=None
        self.dict_vals=None

    def load_data(self):
        try:
            workbook = xlrd.open_workbook(self.file)
            sh = workbook.sheet_by_name(self.default_worksheet)
            #get number of rows
            num_rows = (sh.nrows - 1) - self.header_row
            current_row = -1 + self.header_row
            #print dict(worksheet.row_values(rownum) for rownum in range(worksheet.nrows))
            #setup the map with customer specific values
            header_vals=sh.row_values(self.header_row)
            dict_vals={'date':{'header_index':self.header_vals.index(self.date_col),'vals':[]}, 'description':{'header_index':self.header_vals.index(self.description_col), 'vals':[]},
                             'debit':{'header_index':self.header_vals.index(self.debit_col), 'vals':[]} , 'credit':{'header_index':self.header_vals.index(self.credit_col), 'vals':[]} }

            #initialize the dict on the first row after the header
            current_row+=1
            #populate the dict
            while current_row < num_rows:
                current_row+=1
                row = sh.row_values(current_row)
                dict_vals['date']['vals'].append(xlrd.xldate_as_tuple(row[dict_vals['date']['header_index']], workbook.datemode))
                dict_vals['description']['vals'].append(row[dict_vals['description']['header_index']])
                dict_vals['debit']['vals'].append(row[dict_vals['debit']['header_index']])
                dict_vals['credit']['vals'].append(row[dict_vals['credit']['header_index']])

            self.import_list = import_dict(dict_vals, self.token, self.default_category, self.default_merchant)

        except Exception as e:
            sys.exit('Error loading file and data: %s' % str(e))


    def import_dict(self):

        try:
            import_list=[]
            for index,v in enumerate(self.dict_vals['date']['vals']):
                #build the dictionary for posting to mint
                #get just a single value
                if self.dict_vals['debit']['vals'][index]:
                    txnamount=self.dict_vals['debit']['vals'][index]
                    mtIsExpense=True
                else:
                    txnamount=self.dict_vals['credit']['vals'][index]
                    mtIsExpense=False

                dict_to_import=build_post_dict(self.dict_vals['description']['vals'][index],
                                           self.default_category,
                                           self.default_merchant,
                                           datetime(v[0],v[1], v[2]).strftime('%m/%d/%Y'),
                                           txnamount,
                                           mtIsExpense,
                                           self.token
                                           )
                import_list.append(dict_to_import)

            return import_list
        except Exception as e:
            sys.exit('Error loading values to import into dict %s' % str(e))


    def mint_login(self):
        self.driver.get("https://wwws.mint.com/login.event?task=L&messageId=1&country=US&nextPage=overview.event")
        self.driver.set_window_size(1400,800)
        self.self.driver.find_element_by_id("form-login-username").send_keys(MINT_USERNAME)
        self.self.driver.find_element_by_id("form-login-password").send_keys(MINT_PASSWORD)
        self.driver.find_element_by_id('submit').click()
        sleep(5)
        self.driver.find_element_by_link_text('Transactions').click()
        sleep(10)



    def selenium_transactions(import_list):

        try:
            print 'Starting Transaction import....\n'
            for c,i in enumerate(import_list):
                pass

        except Exception as e:
            print 'Error handling the specific add: %s' % str(e)
            print 'Errored out on this line: %s\n' % i
            print 'Traceback:\n'
            traceback.print_exc()





def main(file, default_worksheet, default_category, date_col, description_col, debit_col, credit_col, header_row, token, default_merchant):
    if not file[-4:] == '.xls':
        sys.exit('You must use an XLS file')
    
    try:
        workbook = xlrd.open_workbook(file)
        sh = workbook.sheet_by_name(default_worksheet)
        #get number of rows
        num_rows = (sh.nrows - 1) - header_row
        current_row = -1 + header_row
        #print dict(worksheet.row_values(rownum) for rownum in range(worksheet.nrows))
        #setup the map with customer specific values
        header_vals=sh.row_values(header_row)
        dict_vals={'date':{'header_index':header_vals.index(date_col),'vals':[]}, 'description':{'header_index':header_vals.index(description_col), 'vals':[]},
                         'debit':{'header_index':header_vals.index(debit_col), 'vals':[]} , 'credit':{'header_index':header_vals.index(credit_col), 'vals':[]} }
        
        #initialize the dict on the first row after the header
        current_row+=1
        #populate the dict
        while current_row < num_rows:
            current_row+=1
            row = sh.row_values(current_row)
            dict_vals['date']['vals'].append(xlrd.xldate_as_tuple(row[dict_vals['date']['header_index']], workbook.datemode))
            dict_vals['description']['vals'].append(row[dict_vals['description']['header_index']])
            dict_vals['debit']['vals'].append(row[dict_vals['debit']['header_index']])
            dict_vals['credit']['vals'].append(row[dict_vals['credit']['header_index']])
        
        
        try:
            #import_dict(dict_vals, token, default_category, default_merchant)
            import_list = import_dict(dict_vals, token, default_category, default_merchant)

            selenium_login(driver)
            selenium_transactions(import_list, driver)
        except Exception as e:
            print 'Error running the import module: %s' % str(e)
            
    except Exception as e:
        print 'There was an error opening your file: %s' % str(e)


def import_specific_transaction():
    print 'Adding %s expense=%s on %s\n\t for %s\n' % (str(i['amount']), str(i['mtIsExpense']), str(i['date']), str(i['note']))
    driver.find_element_by_id('controls-add').click()
    sleep(2)
    #reset it to expense
    driver.find_element_by_id('txnEdit-mt-expense').send_keys(Keys.SPACE)
    #disable the auto subtract from ATM cash
    driver.find_element_by_id('txnEdit-mt-cash-split').click()
    driver.find_element_by_id('txnEdit-date-input').click()
    driver.find_element_by_id('txnEdit-date-input').clear()
    driver.find_element_by_id('txnEdit-date-input').send_keys(i['date']+Keys.TAB)

    ######################Miscellaneous code for complex keystrokes#####################
    #driver.find_element_by_id('txnEdit-date-input').send_keys(i['date'])
    #driver.key_down(Keys.COMMAND).send_keys('c').key_up(Keys.COMMAND).perform()
    #a=action_chains.ActionChains(driver)
    #a.key_down(Keys.COMMAND).send_keys('a').key_up(Keys.COMMAND).perform()
    #driver.find_element_by_id('txnEdit-date-input').
    #driver.find_element_by_id('txnEdit-date-input').send_keys(1)
    ####################################################################################

    driver.find_element_by_id('txnEdit-merchant_input').click()
    driver.find_element_by_id('txnEdit-merchant_input').clear()
    driver.find_element_by_id('txnEdit-merchant_input').send_keys(i['merchant'])
    driver.find_element_by_id('txnEdit-category_input').clear()
    driver.find_element_by_id('txnEdit-category_input').send_keys(i['category'])

    #Before we set the amount, we need to set the type of expense
    if not i['mtIsExpense']:
        #this is a credit, click the credit button
        driver.find_element_by_id('txnEdit-mt-income').send_keys(Keys.SPACE)

    driver.find_element_by_id('txnEdit-amount_input').clear()
    driver.find_element_by_id('txnEdit-amount_input').send_keys(str(i['amount']))
    driver.find_element_by_id('txnEdit-note').click()
    driver.find_element_by_id('txnEdit-note').send_keys('Mint Importer:\n\n%s' % (str(i['note'])))
    driver.find_element_by_id('txnEdit-submit').click()
    print 'Transaction %s of %s successfully added\n' % (str(c+1), str(len(import_list)))
    print '----------------------\n\n'
    sleep(3)





def build_post_dict(description, default_category, default_merchant, txndate, txnamount, mtIsExpense, token):
    post_dict={
        'cashTxnType':'on',
        'mtCheckNo':'',
        'tag461974':'0',
        'tag461975':'0',
        'tag461976':'0',
        'task':'txnadd',
        'txnId':':0',
        'mtType':'cash',
        'mtAccount':'4928795',
        'note':description,
        'isInvestment':'false',
        'catId':'20',
        'category':default_category,
        'merchant':default_merchant,
        'date':txndate, #06/16/2014
        'amount': txnamount, #0.99
        'mtIsExpense':mtIsExpense,
        'mtCashSplitPref':'2',
        'token':token
        }
    return post_dict


if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("file", help='filepath of xls file', nargs='?')
    parser.add_argument("--default_worksheet", help='default name of worksheet in excel', nargs='?',
                        default='Sheet1')
    parser.add_argument("--default_category", help='default category to import transactions into Mint', nargs='?',
                        default='Uncategorized')
    parser.add_argument("--date_col", help='default name of the column that contains the Date of transaction, e.g. Date', nargs='?',
                        default='Date')
    parser.add_argument("--description_col", help='default name of the column that contains the Description of transaction, e.g. Description', nargs='?',
                        default='Transaction Details')
    parser.add_argument("--debit_col", help='default name of the column that contains debit transactions, e.g. Debit', nargs='?',
                        default='Debit')
    parser.add_argument("--credit_col", help='default name of the column that contains crecit transactions, e.g. Credit', nargs='?',
                        default='Credit')
    parser.add_argument("--header_row", help='row that the header starts on, e.g. 0', nargs='?',
                        default=0)
    parser.add_argument("--token", help='if having problems manually input your token you get from inspecting the REST calls', nargs='?',
                        default=None)
    parser.add_argument("--default_merchant", help='this is the value that the merchant will be set to', nargs='?',
                        default='default merchant')
    
    
    args = parser.parse_args()
    main(args.file, args.default_worksheet, args.default_category, args.date_col, args.description_col, args.debit_col, args.credit_col, args.header_row, args.token, args.default_merchant)


