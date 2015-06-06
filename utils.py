
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


