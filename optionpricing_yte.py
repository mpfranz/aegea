"""
Black Model equations for calculating option price and greeks.
* YTE is defined as years to expiration
* Theta and Charm are in years - divide by days per year (252 or 365)
* Uses discounted futures price in place of spot price in BS model
* Functions return None in event of an error
"""

from __future__ import division
from scipy.stats import norm
from math import log, sqrt, exp

def calc_d1(forward, strike, iv, yte):
    d1 = log(forward/strike) / (iv * sqrt(yte)) \
         + iv/2 * sqrt(yte)
    return d1

def calc_d2(forward, strike, iv, yte):
    d2 = log(forward/strike) / (iv * sqrt(yte)) \
         - iv/2 * sqrt(yte)
    return d2
    
def calc_delta(forward, strike, iv, yte, oType):
    oType = oType.upper()
    
    if iv == 0:
        d1 = 1
    else:
        d1 = calc_d1(forward, strike, iv, yte)
    
    x1 = norm.cdf(d1)
    
    if oType == 'C':
        delta = x1
    elif oType == 'P':
        delta = x1 - 1
    else:
        print 'Error - Option type not recognized: ', oType
        return None
    
    return delta
    
def calc_gamma(forward, strike, iv, yte):
    if iv <= 0 or yte <= 0:
        gamma = 0
    else:
        d1 = calc_d1(forward, strike, iv, yte)
        gamma = norm.pdf(d1) / (forward * iv * sqrt(yte))
    
    return gamma

def calc_theta(forward, strike, iv, yte):
    if iv <= 0 or yte <= 0:
        theta = 0
    else:
        d1 = calc_d1(forward, strike, iv, yte)
        theta = -forward * norm.pdf(d1) * iv / (2*sqrt(yte))
        
    return theta
    
def calc_vega(forward, strike, iv, yte):
    if iv <= 0 or yte <= 0:
        vega = 0
    else:
        d1 = calc_d1(forward, strike, iv, yte)
        vega = forward * norm.pdf(d1) * sqrt(yte) / 100
    
    return vega
    
def calc_charm(forward, strike, iv, yte):
    d1 = calc_d1(forward, strike, iv, yte)
    d2 = calc_d2(forward, strike, iv, yte)
    charm = -norm.pdf(d1) * (-d2 * iv * sqrt(yte)) \
            / (2 * yte * iv* sqrt(yte))
    return charm
    
def calc_deltatostrike(forward, delta, iv, yte, oType):
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
    
    strike = forward * exp(-iv * sqrt(yte) * (norm.ppf(delta) \
             - iv * sqrt(yte) / 2))
    
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

def calc_price(forward, strike, iv, yte, oType):
    oType = oType.upper()
    
    if iv <= 0 or yte <= 0:
        price = calc_intrinsic(forward, strike, oType)
    else:
        d1 = calc_d1(forward, strike, iv, yte)
        d2 = calc_d2(forward, strike, iv, yte)
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
    
def calc_forward(spotPrice, interestRate, divYield, yte):
    forward = spotPrice * exp((interestRate - divYield) * yte)
    return forward
    
def calc_impliedvol(forward, strike, yte, oType, actualPrice):
    """Use bisection search to solve for the implied volatilty
    of a given option """
    tol = 0.0000000001
    nIter = 0
    
    volUB, volLB = 10, 0
    volMD = (volUB + volLB) / 2
    
    priceUB = calc_price(forward, strike, volUB, yte, oType)
    priceLB = calc_price(forward, strike, volLB, yte, oType)
    priceMD = calc_price(forward, strike, volMD, yte, oType)
    
    while abs(actualPrice - priceMD) > tol and nIter < 100:
        if priceMD <= actualPrice:
            volLB = volMD
        else:
            volUB = volMD
        
        volMD = (volUB + volLB) / 2
        priceMD = calc_price(forward, strike, volMD, yte, oType)
        nIter += 1
    
    if nIter >= 100:
        print '\n\t Max. iterations (%r) exceeded.' % nIter
        
    return volMD

# def test():
#     print 'Running tests...'
    
#     # forward, strike, iv, yte, oType, delta, optionPrice, spotPrice, r, y
    
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
#         yte = test[3]/252
#         oType = test[4]
#         delta = test[5]
#         price = test [6]
#         spotPrice = test[7]
#         interestRate = test[8]
#         divYield = test[9]
    
#         print '\nTest %r      ' % nTest
#         print 'd1:         ', calc_d1(forward, strike, iv, yte)
#         print 'd2:         ', calc_d2(forward, strike, iv, yte)
#         print 'delta:      ', calc_delta(forward, strike, iv, yte, oType)
#         print 'gamma:      ', calc_gamma(forward, strike, iv, yte)
#         print 'theta:      ', calc_theta(forward, strike, iv, yte)/252
#         print 'vega:       ', calc_vega(forward, strike, iv, yte)
#         print 'charm:      ', calc_charm(forward, strike, iv, yte)/252
#         print 'strike:     ', calc_deltatostrike(forward, delta, iv, yte, oType)
#         print 'price:      ', calc_price(forward, strike, iv, yte, oType)
#         print 'intrinsic:  ', calc_intrinsic(forward, strike, oType)
#         print 'implied vol:', calc_impliedvol(forward, strike, yte, oType, price)
#         print 'forward:    ', calc_forward(spotPrice, interestRate, divYield, yte)
        
#     # print '\nTesting complete. \n'

# ## MAIN
# test()
    

