"""
Calculates MAC margin requirements for a portfolio.
"""

import sys
import os.path
import numpy as np
import optionpricing_tte as bs
import time

def calc_pricerisk(portfolio, timestamp):
    """Accept a portfolio. Calculate MAC price risk.
    Return an array: rows = positions, cols = price shocks."""
    priceShocks = [-0.20, -0.175, -0.15, -0.125, -0.10, -0.08,
                   -0.06,  -0.04, -0.03,  -0.02, -0.01,     0, 
                    0.01,   0.02,  0.03,   0.04,  0.06, 
                    0.08,   0.10, 0.125,   0.15, 0.175,  0.20]
                    
    nPos = portfolio.shape[0]
    nPriceShock = len(priceShocks)
    priceRisk = np.zeros((nPos, nPriceShock))
    
    for row in range(0,nPos):
        oType = portfolio[row, 2]
        
        for col in range(0, len(priceShocks)):
            priceShock = priceShocks[col]
            
            if oType == 3 or oType == 4:
                    # Stock and futures
                    mult = portfolio[row, 16]
                    quantity = portfolio[row, 3]
                    curPrice = portfolio[row, 5]
                    newPrice = curPrice * (1 + priceShock)
                    pnl = mult * quantity * (newPrice - curPrice)
                    priceRisk[row, col] = pnl
                    
            elif oType == 1 or oType == 2:
                    # Options
                    mult = portfolio[row, 16]
                    quantity = portfolio[row, 3]
                    curForward = portfolio[row, 15]
                    newForward = curForward * (1 + priceShock)

                    strike = portfolio[row, 1]
                    iv = portfolio[row, 4]
                    tte = portfolio[row, 14]
                    newPrice = bs.calc_price(newForward, strike, iv, tte, oType)
                    curPrice = portfolio[row, 5]
                    pnl = mult * quantity * (newPrice - curPrice)
                    priceRisk[row, col] = pnl
                    
            else:
                print 'Error - Position type (%r) not recognized.' % oType
                return None
                
    totalPriceRisk = priceRisk.sum(axis = 0)
    worstIdx = np.argmin(totalPriceRisk)
    print 'The worst price shock bucket is %r.' % priceShocks[worstIdx]
    
    outFileName = 'priceRisk' + timestamp + '.csv' 
    np.savetxt(outFileName, priceRisk, delimiter = ',')
    return priceRisk  
                
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
    
def getPortfolio(fileName):
    """CSV to numpy array containing a portfolio."""
    if not os.path.isfile(fileName):
        raise Exception('File does not exist.')
    else:
        portfolio = np.genfromtxt(fileName, delimiter = ',')
    
    return portfolio
    
def timing(f):
    def wrap(*args):
        time1 = time.time()
        ret = f(*args)
        time2 = time.time()
        print 'The %s function took %0.3f ms.' % (f.func_name, (time2-time1)*1000.0)
    return wrap

@timing
def calc_margin(inFileName):
    portfolio = getPortfolio(inFileName)
    timestamp = inFileName[-18:-4]
    
    priceRisk = calc_pricerisk(portfolio, timestamp)


## Main
script, fileName = sys.argv
print 'Calculating margin...'
calc_margin(fileName)
print 'Done.'