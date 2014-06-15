from credentials import *
#from mint.tags import *
#from mint.utils import *
from mint.api import Mint
import pprint
import xlrd
import sys
from time import sleep
from datetime import datetime
import time

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
            import_dict(dict_vals, token, default_category, default_merchant)
        except Exception as e:
            print 'Error running the import module: %s' % str(e)
            
    except Exception as e:
        print 'There was an error opening your file: %s' % str(e)

def import_dict(dict_vals, token, default_category, default_merchant):
    mint=Mint()
    if not token:
        mint=login(username=MINT_USERNAME, password=MINT_PASSWORD)
    else:
        mint.logged_in=True
        mint.token=token
    
    #now iterate through each object
    import_list=[]
    for index,v in enumerate(dict_vals['date']['vals']):
        #build the dictionary for posting to mint
        #get just a single value
        if dict_vals['debit']['vals'][index]:
            txnamount=dict_vals['debit']['vals'][index]
        else:
            txnamount=-dict_vals['credit']['vals'][index]
            
        dict_to_import=build_post_dict(dict_vals['description']['vals'][index],
                                   default_category,
                                   default_merchant,
                                   datetime(v[0],v[1], v[2]).strftime('%m/%d/%Y'),
                                   txnamount,
                                   token
                                   )
        import_list.append(dict_to_import)
    
    for i in import_list:
        response=mint.import_transaction(i)    
        sleep(1)



def build_post_dict(description, default_category, default_merchant, txndate, txnamount, token):
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
        'mtIsExpense':'true',
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


