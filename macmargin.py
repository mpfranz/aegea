"""
Calculates MAC margin requirements for a portfolio.
"""


from __future__ import division
import sys
import os.path
import numpy as np
import optionpricing_dte as bs
import time
from math import sqrt


def calc_pricerisk(portfolio, timestamp):
    """Accept a portfolio. Calculate MAC price risk.
    Return an array: rows = positions, cols = price shocks.
    Save array as a .csv"""
    priceShocks, nShocks = __define__priceshocks()
                    
    nPos = portfolio.shape[0]
    priceRisk = np.zeros((nPos, nShocks))
    
    for row in range(0, nPos):
        pos = portfolio[:][row]
        
        for col in range(0, nShocks):
            shock = priceShocks[col]
            
            if pos['CP'] == 3 or pos['CP'] == 4:
                priceRisk[row, col] = __stock__pricerisk(pos, shock)
            elif pos['CP'] == 1 or pos['CP'] == 2:
                priceRisk[row, col] = __option__pricerisk(pos, shock)
            else:
                raise Exception('Position type not recognized.')
                
    totalPriceRisk = priceRisk.sum(axis = 0)
    worstIdx = np.argmin(totalPriceRisk)
    worstBucket = priceShocks[worstIdx]
    portfolioPriceRisk = np.min(totalPriceRisk)
    
    # print 'The worst price shock bucket is %r.' % worstBucket
    
    np.savetxt('priceRisk' + timestamp + '.csv' , priceRisk, delimiter = ',')
    return portfolioPriceRisk, worstBucket
      
                
def __stock__pricerisk(pos, shock):
    newPrice = pos['Theo_Price'] * (1 + shock)
    pnl      = (pos['Multiplier'] * pos['Quantity']
               * (newPrice - pos['Theo_Price']))
    return pnl


def __option__pricerisk(pos, shock):
    newForward = pos['Theo_Price'] * (1 + shock)
    newPrice   = bs.calc_price(newForward, pos['Strike'], 
                 pos['Implied_Vol'], pos['DTE'], pos['CP'])
    pnl        = (pos['Multiplier'] * pos['Quantity']
                 * (newPrice - pos['Theo_Price']))
    return pnl


def __define__priceshocks():
    priceShocks = [-0.20, -0.175, -0.15, -0.125, -0.10, -0.08,
                   -0.06,  -0.04, -0.03,  -0.02, -0.01,     0, 
                    0.01,   0.02,  0.03,   0.04,  0.06, 
                    0.08,   0.10, 0.125,   0.15, 0.175,  0.20]
    return priceShocks, len(priceShocks)
       
                
def calc_volrisk(portfolio, timestamp):
    """Accept a portfolio. Calculate MAC volatility and vega liquidity risk.
    Return an 1-D array: rows = positions."""
    nPos        = portfolio.shape[0]
    volRisk     = np.zeros((nPos, 1))
    vegaLiqRisk = np.zeros((nPos, 1))
    
    for row in range(0, nPos):
        pos = portfolio[:][row]
        
        if pos['CP'] == 3 or pos['CP'] == 4:
            volRisk[row]        = 0
            vegaLiqRisk[row, :] = 0
        elif pos['CP'] == 1 or pos['CP'] == 2:
            volRisk[row]     = __volrisk(pos)
            vegaLiqRisk[row] = __vegaliqrisk(pos)
        else:
            raise Exception('Position type not recognized.')
                
    np.savetxt('volRisk'     + timestamp + '.csv' , volRisk,     delimiter = ',')
    np.savetxt('vegaLiqRisk' + timestamp + '.csv' , vegaLiqRisk, delimiter = ',')
    
    portfolioVolRisk = volRisk.sum(axis = 0)
    
    netVegaLiqRisk   = abs(vegaLiqRisk.sum(axis = 0))
    grossVegaLiqRisk = np.sum(np.absolute(vegaLiqRisk), axis = 0)
    portfolioLiqRisk = max(netVegaLiqRisk, 0.2 * grossVegaLiqRisk)
    
    return portfolioVolRisk, portfolioLiqRisk


def __volrisk(pos):
    shock = __volshock(pos)
    pnl   = abs(pos['Quantity'] * pos['Vega'] * shock * pos['Implied_Vol'] * 100)
    return pnl
    
    
def __vegaliqrisk(pos):
    liqShift    = __volliquidityshift(pos)
    netLiqRisk  = pos['Quantity'] * pos['Vega'] * liqShift
    return netLiqRisk
    
    
def __volliquidityshift(pos):
    shock     = __volshock(pos)
    liqCharge = __volliquiditycharge(pos)
    liqShift  = min(2, abs(shock * pos['Implied_Vol'] * liqCharge * 100))
    return liqShift
    

def __volshock(pos):
    monToExp = 12 * pos['DTE']/365
    shock    = 0.35 * sqrt(3/monToExp)
    return shock

    
def __volliquiditycharge(pos):
    moneyness = abs((pos['Forward'] - pos['Strike'])/pos['Forward'])
    if moneyness < 0.3:
        liqvolcharge = 0.1
    elif moneyness < 0.45:
        liqvolcharge = 0.2
    else:
        liqvolcharge = 0.3
    return liqvolcharge

    
def getPortfolio(path_to_csv):
    """.csv to numpy structured array containing a portfolio."""
    if not os.path.isfile(path_to_csv):
        raise Exception('File does not exist.')
    else:
        portfolio  = np.genfromtxt(path_to_csv, dtype = float, 
                                   delimiter = ',', names = True)
        clean      = check_for_expiring(portfolio)
        fieldnames = portfolio.dtype.names
        timestamp  = path_to_csv[-18:-4]
    return clean, fieldnames, timestamp


def check_for_expiring(portfolio, daysForward = 0):
    """Converts the net delta of any options with a expired (negative DTE) to
    SPY. Expired options are removed from the position."""
    if abs(daysForward) > 0:
        portfolio['DTE'][:] = portfolio['DTE'][:] - abs(daysForward)
    
    expired      = portfolio[:][portfolio['DTE'] <= 0]
    expiredDelta = 10 * (expired['Delta'] * expired['Quantity']).sum(axis = 0)
    
    notExpired = portfolio[:][portfolio['DTE'] > 0]
    stockIdx   = np.where(notExpired['CP'] == 3.0)

    notExpired['Quantity'][stockIdx] = (notExpired['Quantity'][stockIdx]
                                        + expiredDelta)
    np.savetxt('nonExpiredPosition.csv' , notExpired, delimiter = ',')
    return notExpired
    
    
def timing(f):
    def wrap(*args):
        time1 = time.time()
        val = f(*args)
        time2 = time.time()
        print 'The %s function took %0.3f ms.' % (f.func_name, 
              (time2-time1)*1000.0)
    return wrap


@timing
def calc_margin(inFileName):
    portfolio, fieldNames, timestamp = getPortfolio(inFileName)
    
    priceRisk, worstBucket = calc_pricerisk(portfolio, timestamp)
    volRisk, vegaLiqRisk   = calc_volrisk(portfolio, timestamp)
    
    macMargin = priceRisk + volRisk[0] + vegaLiqRisk[0]
    
    print 'Portfolio Price Risk is %r from bucket %d.' % (priceRisk, worstBucket)
    print 'Portfolio Volatility Risk is %d.' % volRisk[0]
    print 'Portfolio Vega Liquidity Risk is %d.' % vegaLiqRisk[0]
    print 'Portfolio MAC Margin is %d.' % macMargin
    return macMargin


## Main
script, fileName = sys.argv
print 'Calculating margin for the position at %r.' % fileName[-17:-4]
macMargin = calc_margin(fileName)
print 'Done.'