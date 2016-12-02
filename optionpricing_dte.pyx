"""
Black Model equations for calculating option price and greeks.
* DTE is defined as the number of CALENDAR days to expiration
* Uses discounted futures price in place of spot price in BS model
* Functions return None in event of an error
"""

from __future__ import division
from scipy.stats import norm
from math import log, sqrt, exp

def calc_d1(forward, strike, iv, dte):
    d1 = log(forward/strike) / (iv * sqrt(dte/365)) \
         + iv/2 * sqrt(dte/365)
    return d1

def calc_d2(forward, strike, iv, dte):
    d2 = log(forward/strike) / (iv * sqrt(dte/365)) \
         - iv/2 * sqrt(dte/365)
    return d2
    
def calc_delta(forward, strike, iv, dte, oType):
    if iv == 0:
        d1 = 1
    else:
        d1 = calc_d1(forward, strike, iv, dte)
    
    x1 = norm.cdf(d1)
    
    if oType == 1:
        delta = x1
    elif oType == 2:
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
        gamma = norm.pdf(d1) / (forward * iv * sqrt(dte/365))
    
    return gamma

def calc_theta(forward, strike, iv, dte):
    if iv <= 0 or dte <= 0:
        theta = 0
    else:
        d1 = calc_d1(forward, strike, iv, dte)
        theta = -forward * norm.pdf(d1) * iv / (2*sqrt(dte/365)) / 365
        
    return theta
    
def calc_vega(forward, strike, iv, dte):
    if iv <= 0 or dte <= 0:
        vega = 0
    else:
        d1 = calc_d1(forward, strike, iv, dte)
        vega = forward * norm.pdf(d1) * sqrt(dte/365) / 100
    
    return vega
    
def calc_charm(forward, strike, iv, dte):
    d1 = calc_d1(forward, strike, iv, dte)
    d2 = calc_d2(forward, strike, iv, dte)
    charm = -norm.pdf(d1) * (-d2 * iv * sqrt(dte/365)) \
            / (2 * dte/365 * iv* sqrt(dte/365)) / 365
    return charm
    
def calc_deltatostrike(forward, delta, iv, dte, oType):
    if delta > 1:
        delta = delta / 100
    else:
        delta = delta
    
    if oType == 1:
        delta = delta
    elif oType == 2:
        delta = -delta
    else:
        print 'Error - Option Type not recognized: ', oType
        return None
    
    strike = forward * exp(-iv * sqrt(dte/365) * (norm.ppf(delta) \
             - iv * sqrt(dte/365) / 2))
    
    return strike

def calc_intrinsic(forward, strike, oType):
    if oType == 1:
        if forward > strike:
            intrinsic = forward - strike
        else:
            intrinsic = 0
            
    elif oType == 2:
        if forward < strike:
            intrinsic = strike - forward
        else:
            intrinsic = 0
            
    else:
        print 'Error - Option Type not recognized: ', oType
        return None
    
    return intrinsic

def calc_price(forward, strike, iv, dte, oType):
    if iv <= 0 or dte <= 0:
        price = calc_intrinsic(forward, strike, oType)
    else:
        d1 = calc_d1(forward, strike, iv, dte)
        d2 = calc_d2(forward, strike, iv, dte)
        x1 = norm.cdf(d1)
        x2 = norm.cdf(d2)
        
        if oType == 1:
            price = forward * x1 - strike * x2
        elif oType == 2:
            price = strike * (1 - x2) - forward *(1 - x1)
        else:
            print 'Error - Option Type %r not recognized.' % oType
            return None
        
    return price
    
def calc_forward(spotPrice, interestRate, divYield, dte):
    forward = spotPrice * exp((interestRate - divYield) * dte / 365)
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