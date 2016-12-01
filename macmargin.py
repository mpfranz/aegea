"""
Calculates MAC margin requirements for a portfolio.
    * http://pandas.pydata.org/pandas-docs/stable/install.html
"""

import pandas as pd
import numpy as np
import scipy.io as sio
import os.path
import optionpricing_tte

def calc_pricerisk():
    return None
    
def calc_volrisk():
    return None
    
def calc_volliquidityrisk():
    return None
    
def calc_volliquidityshift():
    return None
    
def calc_volliquiditycharge():
    return None
    
def calc_volshock():
    return None
    
def get_position(fileName):
    """Loads a .mat file into a numpy nd-array. 
       File name must include full path."""
    data = sio.loadmat(fileName)
    return data['pos'], data['comboList']
    
def test():
    fileName = '/home/ubuntu/workspace/Position_Total1.mat'
    if not os.path.isfile(fileName):
        print 'File not found.'
    else:
        portfolio, comboList = get_position(fileName)
        for position in range(portfolio.shape[0]):
            # forward
            # tte
            oType    = portfolio[position, 1]
            strike   = portfolio[position, 2]
            quantity = portfolio[position, 5]
            
            # print oType, strike, quantity

    print 'Done.'
    
## Main
test()