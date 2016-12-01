"""
Black Model equations for calculating option price and greeks.
* DTE is defined as the number of TRADING days to expiration
* Uses discounted futures price in place of spot price in BS model
* Functions return None in event of an error
"""

from __future__ import division
from scipy.stats import norm
from math import log, sqrt, exp

def calc_d1(forward, strike, iv, dte):
    d1 = log(forward/strike) / (iv * sqrt(dte/252)) \
         + iv/2 * sqrt(dte/252)
    return d1

def calc_d2(forward, strike, iv, dte):
    d2 = log(forward/strike) / (iv * sqrt(dte/252)) \
         - iv/2 * sqrt(dte/252)
    return d2
    
def calc_delta(forward, strike, iv, dte, oType):
    oType = oType.upper()
    
    if iv == 0:
        d1 = 1
    else:
        d1 = calc_d1(forward, strike, iv, dte)
    
    x1 = norm.cdf(d1)
    
    if oType == 'C':
        delta = x1
    elif oType == 'P':
        delta = x1 - 1
    else:
        print 'Error - Option type not recognized: ', oType
        return None
    
    return delta
    
def calc_gamma(forward, strike, iv, dte):
    if iv <= 0 or dte <= 0:
        gamma = 0
    else:
        d1 = calc_d1(forward, strike, iv, dte)
        gamma = norm.pdf(d1) / (forward * iv * sqrt(dte/252))
    
    return gamma

def calc_theta(forward, strike, iv, dte):
    if iv <= 0 or dte <= 0:
        theta = 0
    else:
        d1 = calc_d1(forward, strike, iv, dte)
        theta = -forward * norm.pdf(d1) * iv / (2*sqrt(dte/252)) / 252
        
    return theta
    
def calc_vega(forward, strike, iv, dte):
    if iv <= 0 or dte <= 0:
        vega = 0
    else:
        d1 = calc_d1(forward, strike, iv, dte)
        vega = forward * norm.pdf(d1) * sqrt(dte/252) / 100
    
    return vega
    
def calc_charm(forward, strike, iv, dte):
    d1 = calc_d1(forward, strike, iv, dte)
    d2 = calc_d2(forward, strike, iv, dte)
    charm = -norm.pdf(d1) * (-d2 * iv * sqrt(dte/252)) \
            / (2 * dte/252 * iv* sqrt(dte/252)) / 252
    return charm
    
def calc_deltatostrike(forward, delta, iv, dte, oType):
    oType = oType.upper()
    
    if delta > 1:
        delta = delta / 100
    else:
        delta = delta
    
    if oType == 'C':
        delta = delta
    elif oType == 'P':
        delta = -delta
    else:
        print 'Error - Option Type not recognized: ', oType
        return None
    
    strike = forward * exp(-iv * sqrt(dte/252) * (norm.ppf(delta) \
             - iv * sqrt(dte/252) / 2))
    
    return strike

def calc_intrinsic(forward, strike, oType):
    oType = oType.upper()
    
    if oType == 'C':
        if forward > strike:
            intrinsic = forward - strike
        else:
            intrinsic = 0
            
    elif oType == 'P':
        if forward < strike:
            intrinsic = strike - forward
        else:
            intrinsic = 0
            
    else:
        print 'Error - Option Type not recognized: ', oType
        return None
    
    return intrinsic

def calc_price(forward, strike, iv, dte, oType):
    oType = oType.upper()
    
    if iv <= 0 or dte <= 0:
        price = calc_intrinsic(forward, strike, oType)
    else:
        d1 = calc_d1(forward, strike, iv, dte)
        d2 = calc_d2(forward, strike, iv, dte)
        x1 = norm.cdf(d1)
        x2 = norm.cdf(d2)
        
        if oType == 'C':
            price = forward * x1 - strike * x2
        elif oType == 'P':
            price = strike * (1 - x2) - forward *(1 - x1)
        else:
            print 'Error - Option Type %r not recognized.' % oType
            return None
        
    return price
    
def calc_forward(spotPrice, interestRate, divYield, dte):
    forward = spotPrice * exp((interestRate - divYield) * dte / 252)
    return forward
    
def calc_impliedvol(forward, strike, dte, oType, actualPrice):
    """Use bisection search to solve for the implied volatilty
    of a given option """
    tol = 0.0000000001
    nIter = 0
    
    volUB, volLB = 10, 0
    volMD = (volUB + volLB) / 2
    
    priceUB = calc_price(forward, strike, volUB, dte, oType)
    priceLB = calc_price(forward, strike, volLB, dte, oType)
    priceMD = calc_price(forward, strike, volMD, dte, oType)
    
    while abs(actualPrice - priceMD) > tol and nIter < 100:
        if priceMD <= actualPrice:
            volLB = volMD
        else:
            volUB = volMD
        
        volMD = (volUB + volLB) / 2
        priceMD = calc_price(forward, strike, volMD, dte, oType)
        nIter += 1
    
    if nIter >= 100:
        print '\n\t Max. iterations (%r) exceeded.' % nIter
        
    return volMD

# def test():
#     print 'Running tests...'
    
#     # forward, strike, iv, dte, oType, delta, optionPrice, spotPrice, r, y
    
#     t1 = [2209.25, 2210, 0.1002, 15, 'C',  50.43, 21.60, 0, 0, 0]
#     # t2 = [2209.25, 2210, 0.1002, 15, 'c',  50.43, 21.60, 0, 0, 0]
#     # t3 = [2209.25, 2210, 0.1002, 15, 'P', -49.57, 21.60, 0, 0, 0]
#     # t4 = [2209.25, 2210, 0.1002, 15, 'p', -49.57, 21.60, 0, 0, 0]
    
#     tests = [t1]
#     nTest = 0
    
#     for test in tests:
#         nTest = nTest + 1
#         forward = test[0]
#         strike = test[1]
#         iv = test[2]
#         dte = test[3]
#         oType = test[4]
#         delta = test[5]
#         price = test [6]
#         spotPrice = test[7]
#         interestRate = test[8]
#         divYield = test[9]
    
#         print '\nTest %r      ' % nTest
#         print 'd1:         ', calc_d1(forward, strike, iv, dte)
#         print 'd2:         ', calc_d2(forward, strike, iv, dte)
#         print 'delta:      ', calc_delta(forward, strike, iv, dte, oType)
#         print 'gamma:      ', calc_gamma(forward, strike, iv, dte)
#         print 'theta:      ', calc_theta(forward, strike, iv, dte)
#         print 'vega:       ', calc_vega(forward, strike, iv, dte)
#         print 'charm:      ', calc_charm(forward, strike, iv, dte)
#         print 'strike:     ', calc_deltatostrike(forward, delta, iv, dte, oType)
#         print 'price:      ', calc_price(forward, strike, iv, dte, oType)
#         print 'intrinsic:  ', calc_intrinsic(forward, strike, oType)
#         print 'implied vol:', calc_impliedvol(forward, strike, dte, oType, price)
#         print 'forward:    ', calc_forward(spotPrice, interestRate, divYield, dte)
        
#     # print '\nTesting complete. \n'

# ## MAIN
# test()
    

