# -*- coding: utf-8 -*-
"""
Created on Thu Dec 12 14:36:36 2019

@author: daple
"""



import numpy as np
from scipy.stats import norm
import xlwings as xw
from scipy.special import erf




@xw.func
def Kirk(S1_t, S2_t, K, T, r, vol1, vol2, rho, CallPut):
    optprice=kirk(S1_t, S2_t, K, T, r, vol1, vol2, rho, CallPut)
    return optprice.price




class kirk():
    def __init__ (self, S1_t, S2_t, K, T, r, vol1, vol2, rho, CallPut):
        #SpreadOption.__init__(self, S1_t, S2_t, K, T, r, vol1, vol2, rho, CallPut, "Kirk")
  
        self.S1_t = float(S1_t)
        self.S2_t = float(S2_t)
        self.K = float(K)
        self.T = float(T)
        self.r = float(r)
        self.vol1 = float(vol1)
        self.vol2 = float(vol2)
        self.rho = float(rho)
        self.CallPut = int(CallPut)

          
    @property
    def price(self):
        z = self.S1_t / (self.S1_t + self.K * np.exp(-1. * self.r * self.T))
        vol = np.sqrt( self.vol1 ** 2 * z ** 2 + self.vol2 ** 2 - 2 * self.rho* self.vol1 * self.vol2 * z )
        d1 = (np.log(self.S2_t / (self.S1_t + self.K * np.exp(-self.r * self.T)))
              / (vol * np.sqrt(self.T)) + 0.5 * vol * np.sqrt(self.T))
        d2 = d1 - vol * np.sqrt(self.T)
        price = (self.CallPut * (self.S2_t  * norm.cdf(self.CallPut * d1, 0, 1) 
                                 - (self.S1_t + self.K * np.exp(-self.r * self.T)) 
                                 * norm.cdf(self.CallPut * d2, 0, 1)))
        return price
    
 