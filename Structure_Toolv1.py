# coding: utf8
"""
Created on Mon Mar 30 17:19:11 2020

@author: daple
"""

  

import numpy as np
from scipy.stats import norm
import xlwings as xw
import pandas as pd
from scipy import stats
from scipy.special import erf



def PowerOpt_BS():
    sb = xw.Book.caller()   
    #sb = xw.Book(r'C:\Users\daple\OneDrive\Documents\Quant Structuring Energy\Power Origination and structuring\Structure_Toolv1.xlsm')
    #ControlSht = sb.sheets['Control']
    OutputSht = sb.sheets['PowerOpt']    
    PowerOptSht=sb.sheets('PowerOpt').range('A9').expand().options(pd.DataFrame).value    
    price=PowerOptSht['Base EU/MWh'].astype(float) 
    strike=PowerOptSht['Strike'].astype(float) 
    vols=PowerOptSht['VOL POWER'].astype(float)
    dt=PowerOptSht['T'].astype(float)    
    Days=PowerOptSht['Num_Days'].astype(float)
    powerperiod=PowerOptSht['Peak_OffPeak'].astype(str)
    capacity=PowerOptSht['Capacity_Factor'].astype(float)
    int_rate=PowerOptSht['Interest_rate'].astype(float)
    #CallPut=PowerOptSht['CallPut'].astype(str)
    CallPut=OutputSht.range('K10').value
    #simulations=PowerOptSht['Simulation'].astype(float)
    
    #Review logic
    #if powerperiod =='OP':
    #    powerprice = price * 0.35
    #if powerperiod =='Peak':
    #    powerprice = price * 0.65
    #if powerperiod =='Base':
    powerprice = price 
            
    optval = Black76_Option(powerprice, strike, dt, int_rate, vols, CallPut)
    deltaval = deltabs(powerprice, strike, dt, int_rate, vols, CallPut)    
    vegaval = vegabs(powerprice, strike, dt, int_rate, vols, CallPut)    
    gammaval = gammabs(powerprice, strike, dt, int_rate, vols, CallPut)
    intrinsic = powerprice - strike
    extrinsic = optval - intrinsic    
    frames = [optval, intrinsic, extrinsic, deltaval, vegaval,gammaval]    
    dfData = pd.concat(frames, axis=1, sort=False) 
    dfData=dfData.set_index(0)
    OutputSht.range('L9').value =  dfData
    OutputSht.range('L9').value = 'FV'
    OutputSht.range('M9').value = 'Intrinsic'
    OutputSht.range('N9').value = 'Extrinsic'
    OutputSht.range('O9').value = 'Delta Pwr'
    OutputSht.range('P9').value = 'Vega Pwr'
    OutputSht.range('Q9').value = 'Gamma Pwr'

    
def Power_Collar(): 
    sb = xw.Book.caller()  
    #sb = xw.Book(r'C:\Users\daple\OneDrive\Documents\Quant Structuring Energy\Power Origination and structuring\Structure_Toolv1.xlsm')
    #ControlSht = sb.sheets['Control']
    OutputSht = sb.sheets['PowerOpt-Collar']     
    PowerOptSht=sb.sheets('PowerOpt-Collar').range('A9').expand().options(pd.DataFrame).value  
    price=PowerOptSht['Base EU/MWh'].astype(float) 
    strikecall=PowerOptSht['Strike_1'].astype(float) 
    strikeput=PowerOptSht['Strike_2'].astype(float)         
    vols=PowerOptSht['VOL POWER'].astype(float)
    dt=PowerOptSht['T'].astype(float)    
    Days=PowerOptSht['Num_Days'].astype(float)
    powerperiod=PowerOptSht['Peak_OffPeak'].astype(str)
    capacity=PowerOptSht['Capacity_Factor'].astype(float)
    int_rate=PowerOptSht['Interest_rate'].astype(float)
    #Call=PowerOptSht['Call'].astype(float)
    #Put=PowerOptSht['Put'].astype(float)
    Call=OutputSht.range('L10').value
    Put=OutputSht.range('M10').value
    #simulations=PowerOptSht['Simulation'].astype(float)
    
    #if powerperiod =='OP':
    #    powerprice = price * 0.35
    #if powerperiod =='Peak':
    #    powerprice = price * 0.65
    #if powerperiod =='Base':
    powerprice = price 
    
    optval_call = Black76_Option(powerprice, strikecall, dt, int_rate, vols, Call)
    optval_put = Black76_Option(powerprice, strikeput, dt, int_rate, vols, Put)
    
    deltaval_call = deltabs(powerprice, strikecall, dt, int_rate, vols, Call)    
    deltaval_put = deltabs(powerprice, strikeput, dt, int_rate, vols, Put)    
    vegaval_call = vegabs(powerprice, strikecall, dt, int_rate, vols, Call)    
    vegaval_put = vegabs(powerprice, strikeput, dt, int_rate, vols, Put)        
    gammaval_call = gammabs(powerprice, strikecall, dt, int_rate, vols, Call)
    gammaval_put = gammabs(powerprice, strikeput, dt, int_rate, vols, Put)
    
    intrinsic_call = powerprice - strikecall
    intrinsic_put = powerprice - strikeput
    
    extrinsic_call = optval_call - intrinsic_call
    extrinsic_put = optval_put - intrinsic_put    
    
    frames = [optval_call, optval_put, intrinsic_call, intrinsic_put, extrinsic_call, extrinsic_put,deltaval_call, deltaval_put, vegaval_call, vegaval_put, gammaval_call, gammaval_put]
                  
    dfData = pd.concat(frames, axis=1, sort=False)
    dfData=dfData.set_index(0)
    OutputSht.range('N9').value =  dfData
    OutputSht.range('N9').value = 'FV_Call'
    OutputSht.range('O9').value = 'FV_Put'    
    OutputSht.range('P9').value = 'Intrinsic_1'
    OutputSht.range('Q9').value = 'Intrinsic_2'    
    OutputSht.range('R9').value = 'Extrinsic_1'
    OutputSht.range('S9').value = 'Extrinsic_2'    
    OutputSht.range('T9').value = 'Delta_1'
    OutputSht.range('U9').value = 'Delta_2'    
    OutputSht.range('V9').value = 'Vega_1'
    OutputSht.range('W9').value = 'Vega_2'    
    OutputSht.range('X9').value = 'Gamma_1'
    OutputSht.range('Y9').value = 'Gamma_2'
    #OutputSht.range('Z10').value = -optval_call + optavl_puth

    
def Power_Enhanced_Collar():    
    sb = xw.Book.caller()   
    #sb = xw.Book(r'C:\Users\daple\OneDrive\Documents\Quant Structuring Energy\Power Origination and structuring\Structure_Toolv1.xlsm')
    #ControlSht = sb.sheets['Control']
    OutputSht = sb.sheets['PowerOpt-EnhancedCollar']    
    PowerOptSht=sb.sheets('PowerOpt-EnhancedCollar').range('A9').expand().options(pd.DataFrame).value
    price=PowerOptSht['Base EU/MWh'].astype(float) 
    strikecall=PowerOptSht['Strike_1'].astype(float) 
    strikeputh=PowerOptSht['Strike_2'].astype(float)         
    strikeputl=PowerOptSht['Strike_3'].astype(float)         
    vols=PowerOptSht['VOL POWER'].astype(float)
    dt=PowerOptSht['T'].astype(float)    
    Days=PowerOptSht['Num_Days'].astype(float)
    powerperiod=PowerOptSht['Peak_OffPeak'].astype(str)
    capacity=PowerOptSht['Capacity_Factor'].astype(float)
    int_rate=PowerOptSht['Interest_rate'].astype(float)
    #Call=PowerOptSht['Call'].astype(float)
    #Puth=PowerOptSht['Put_H'].astype(float)
    #Putl=PowerOptSht['Put_L'].astype(float)
    
    Call=OutputSht.range('M10').value
    Puth=OutputSht.range('N10').value
    Putl=OutputSht.range('O10').value    
    simulations=PowerOptSht['Simulation'].astype(float)
    
    #if powerperiod =='OP':
    #    powerprice = price * 0.35
    #if powerperiod =='Peak':
    #    powerprice = price * 0.65
    #if powerperiod =='Base':
    powerprice = price 
    
    optval_call = Black76_Option(powerprice, strikecall, dt, int_rate, vols, Call)
    optval_puth = Black76_Option(powerprice, strikeputh, dt, int_rate, vols, Puth)
    optval_putl = Black76_Option(powerprice, strikeputl, dt, int_rate, vols, Putl)    
    
    deltaval_call = deltabs(powerprice, strikecall, dt, int_rate, vols, Call)    
    deltaval_puth = deltabs(powerprice, strikeputh, dt, int_rate, vols, Puth)
    deltaval_putl = deltabs(powerprice, strikeputl, dt, int_rate, vols, Putl)
    
    vegaval_call = vegabs(powerprice, strikecall, dt, int_rate, vols, Call)    
    vegaval_puth = vegabs(powerprice, strikeputh, dt, int_rate, vols, Puth)
    vegaval_putl = vegabs(powerprice, strikeputl, dt, int_rate, vols, Putl)
      
    gammaval_call = gammabs(powerprice, strikecall, dt, int_rate, vols, Call)
    gammaval_puth = gammabs(powerprice, strikeputh, dt, int_rate, vols, Puth)
    gammaval_putl = gammabs(powerprice, strikeputl, dt, int_rate, vols, Putl)    
    
    intrinsic_call = powerprice - strikecall
    intrinsic_puth = powerprice - strikeputh
    intrinsic_putl = powerprice - strikeputl    
    
    extrinsic_call = optval_call - intrinsic_call
    extrinsic_puth = optval_puth - intrinsic_putl    
    extrinsic_putl = optval_putl - intrinsic_putl     
    
    frames = [optval_call, optval_puth, optval_putl, intrinsic_call, intrinsic_puth, intrinsic_putl, extrinsic_call, extrinsic_puth, extrinsic_putl 
                  , deltaval_call, deltaval_puth,deltaval_putl, vegaval_call, vegaval_puth, vegaval_putl, gammaval_call, gammaval_puth, gammaval_putl]    
    dfData = pd.concat(frames, axis=1, sort=False)
    dfData=dfData.set_index(0)
    
    OutputSht.range('P9').value =  dfData
    OutputSht.range('P9').value = 'FV_1'
    OutputSht.range('Q9').value = 'FV_2'  
    OutputSht.range('R9').value = 'FV_3'    
    OutputSht.range('S9').value = 'Intrinsic_1'
    OutputSht.range('T9').value = 'Intrinsic_2' 
    OutputSht.range('U9').value = 'Intrinsic_3'    
    OutputSht.range('V9').value = 'Extrinsic_1'
    OutputSht.range('W9').value = 'Extrinsic_2' 
    OutputSht.range('X9').value = 'Extrinsic_3'    
    OutputSht.range('Y9').value = 'Delta_1'
    OutputSht.range('Z9').value = 'Delta_2'    
    OutputSht.range('AA9').value = 'Delta_3'    
    OutputSht.range('AB9').value = 'Vega_1'
    OutputSht.range('AC9').value = 'Vega_2'    
    OutputSht.range('AD9').value = 'Vega_3'    
    OutputSht.range('AE9').value = 'Gamma_1'
    OutputSht.range('AF9').value = 'Gamma_2'    
    OutputSht.range('AG9').value = 'Gamma_3'    
    #OutputSht.range('AH9').value = -optval_call + optavl_puth - optval_putl


def CorporatePPA():    
    sb = xw.Book.caller()   
    #sb = xw.Book(r'C:\Users\daple\OneDrive\Documents\Quant Structuring Energy\Power Origination and structuring\Structure_Toolv1.xlsm')
    #ControlSht = sb.sheets['Control']
    OutputSht = sb.sheets['CorporatePPA']    
    PowerOptSht=sb.sheets('CorporatePPA').range('A9').expand().options(pd.DataFrame).value
    price=PowerOptSht['Base EU/MWh'].astype(float) 
    strikecall=PowerOptSht['Strike_1'].astype(float) 
    strikeputh=PowerOptSht['Strike_2'].astype(float)         
    strikeputl=PowerOptSht['Strike_3'].astype(float)         
    vols=PowerOptSht['VOL POWER'].astype(float)
    dt=PowerOptSht['T'].astype(float)    
    Days=PowerOptSht['Num_Days'].astype(float)
    powerperiod=PowerOptSht['Peak_OffPeak'].astype(str)
    fixed_price=PowerOptSht['Fixed_Price'].astype(float)
    int_rate=PowerOptSht['Interest_rate'].astype(float)
    #Call=PowerOptSht['Call'].astype(float)
    #Puth=PowerOptSht['Put_H'].astype(float)
    #Putl=PowerOptSht['Put_L'].astype(float)
    
    Call=OutputSht.range('M10').value
    Puth=OutputSht.range('N10').value
    Putl=OutputSht.range('O10').value    
    simulations=PowerOptSht['Fixed_Price'].astype(float)
    
    #if powerperiod =='OP':
    #    powerprice = price * 0.35
    #if powerperiod =='Peak':
    #    powerprice = price * 0.65
    #if powerperiod =='Base':
    powerprice = price 
    
    optval_call = Black76_Option(powerprice, strikecall, dt, int_rate, vols, Call)
    optval_puth = Black76_Option(powerprice, strikeputh, dt, int_rate, vols, Puth)
    optval_putl = Black76_Option(powerprice, strikeputl, dt, int_rate, vols, Putl)    
    
    deltaval_call = deltabs(powerprice, strikecall, dt, int_rate, vols, Call)    
    deltaval_puth = deltabs(powerprice, strikeputh, dt, int_rate, vols, Puth)
    deltaval_putl = deltabs(powerprice, strikeputl, dt, int_rate, vols, Putl)
    
    vegaval_call = vegabs(powerprice, strikecall, dt, int_rate, vols, Call)    
    vegaval_puth = vegabs(powerprice, strikeputh, dt, int_rate, vols, Puth)
    vegaval_putl = vegabs(powerprice, strikeputl, dt, int_rate, vols, Putl)
      
    gammaval_call = gammabs(powerprice, strikecall, dt, int_rate, vols, Call)
    gammaval_puth = gammabs(powerprice, strikeputh, dt, int_rate, vols, Puth)
    gammaval_putl = gammabs(powerprice, strikeputl, dt, int_rate, vols, Putl)    
    
    intrinsic_call = powerprice - strikecall
    intrinsic_puth = powerprice - strikeputh
    intrinsic_putl = powerprice - strikeputl    
    
    extrinsic_call = optval_call - intrinsic_call
    extrinsic_puth = optval_puth - intrinsic_putl    
    extrinsic_putl = optval_putl - intrinsic_putl     
    
    frames = [optval_call, optval_puth, optval_putl, intrinsic_call, intrinsic_puth, intrinsic_putl, extrinsic_call, extrinsic_puth, extrinsic_putl 
                  , deltaval_call, deltaval_puth,deltaval_putl, vegaval_call, vegaval_puth, vegaval_putl, gammaval_call, gammaval_puth, gammaval_putl]    
    dfData = pd.concat(frames, axis=1, sort=False)
    dfData=dfData.set_index(0)
    
    OutputSht.range('P9').value =  dfData
    OutputSht.range('P9').value = 'FV_1'
    OutputSht.range('Q9').value = 'FV_2'  
    OutputSht.range('R9').value = 'FV_3'    
    OutputSht.range('S9').value = 'Intrinsic_1'
    OutputSht.range('T9').value = 'Intrinsic_2' 
    OutputSht.range('U9').value = 'Intrinsic_3'    
    OutputSht.range('V9').value = 'Extrinsic_1'
    OutputSht.range('W9').value = 'Extrinsic_2' 
    OutputSht.range('X9').value = 'Extrinsic_3'    
    OutputSht.range('Y9').value = 'Delta_1'
    OutputSht.range('Z9').value = 'Delta_2'    
    OutputSht.range('AA9').value = 'Delta_3'    
    OutputSht.range('AB9').value = 'Vega_1'
    OutputSht.range('AC9').value = 'Vega_2'    
    OutputSht.range('AD9').value = 'Vega_3'    
    OutputSht.range('AE9').value = 'Gamma_1'
    OutputSht.range('AF9').value = 'Gamma_2'    
    OutputSht.range('AG9').value = 'Gamma_3'    
    #OutputSht.range('AH9').value = -optval_call + optavl_puth - optval_putl

def SparkSpread_Kirk():
    sb = xw.Book.caller()
    #sb = xw.Book(r'C:\Users\daple\OneDrive\Documents\Quant Structuring Energy\Power Origination and structuring\Structure_Toolv1.xlsm')
    OutputSht = sb.sheets['SparkSpread-Tolling']    
    VecSht=sb.sheets('SparkSpread-Tolling').range('A9').expand().options(pd.DataFrame).value    

    powerperiod= sb.sheets('SparkSpread-Tolling').range('N10').value  
    heatrate= OutputSht.range('J3').value
    carbonrate= OutputSht.range('J4').value
    baseprice=VecSht['Base EU/MWh'].astype(float)
    Peakprice=VecSht['Peak EU/MWh'].astype(float)
    offprice=VecSht['OP EU/MWh'].astype(float)
    gasprice=VecSht['GAS EU/MWh'].astype(float) / heatrate
    carbonprice=VecSht['CO2 EU/ton'].astype(float) * carbonrate
    timemat= VecSht['T'].astype(float)
    volpower= VecSht['VOL_GAS'].astype(float)
    volgas= VecSht['VOL_PWR'].astype(float)
    correlvec= VecSht['Correlation POWER vs "GAS + CO2"'].astype(float)
    capfactor=VecSht['Capacity_Factor'].astype(float)
    #callPut= VecSht['CallPut'].astype(float)
    callPut=OutputSht.range('Q10').value    
    intrate=VecSht['Interest_rate'].astype(float)
    strike=VecSht['Strike'].astype(float) 
    gasprice = carbonprice + gasprice
    
    #if powerperiod =='OP':
    #    powerprice = offprice
    #if powerperiod =='Peak':
    #    powerprice = Peakprice
    #if powerperiod =='Base':
    powerprice = baseprice
    
    optval = KirkModel(gasprice, powerprice , strike, timemat, intrate, volpower, volgas, correlvec, callPut)# * capfactor
    deltapwr = delta(gasprice, powerprice , carbonprice, timemat, intrate, volpower, volgas, correlvec, callPut)# * capfactor
    deltagas = delta2(gasprice, powerprice , carbonprice, timemat, intrate, volpower, volgas, correlvec, callPut)# * capfactor
    vegapwr = vega(gasprice, powerprice , carbonprice, timemat, intrate, volpower, volgas, correlvec, callPut) #* capfactor
    vegagas = vega2(gasprice, powerprice , carbonprice, timemat, intrate, volpower, volgas, correlvec, callPut) #* capfactor
    gammapwr = gamma(gasprice, powerprice , carbonprice, timemat, intrate, volpower, volgas, correlvec, callPut) #* capfactor
    gammagas = gamma2(gasprice, powerprice , carbonprice, timemat, intrate, volpower, volgas, correlvec, callPut) #* capfactor
   
    
    intrinsic = powerprice - gasprice - carbonprice
    extrinsic = optval - intrinsic
    
    frames = [optval, intrinsic, extrinsic, deltapwr, deltagas, vegapwr, vegagas, gammapwr,  gammagas]    
    dfData = pd.concat(frames, axis=1, sort=False)
    dfData=dfData.set_index(0)    
    
    OutputSht.range('R9').value =  dfData
    OutputSht.range('R9').value = 'FV'
    OutputSht.range('S9').value = 'Intrinsic'
    OutputSht.range('T9').value = 'Extrinsic'
    OutputSht.range('U9').value = 'Delta Pwr'
    OutputSht.range('V9').value = 'Delta Gas'
    OutputSht.range('W9').value = 'Vega Pwr'
    OutputSht.range('X9').value = 'Vega Gas'
    OutputSht.range('Y9').value = 'Gamma Pwr'
    OutputSht.range('Z9').value = 'Gamma Gas'
    


def SparkSpread_Daily(): 

    sb = xw.Book.caller()
    #sb = xw.Book(r'C:\Users\daple\OneDrive\Documents\Quant Structuring Energy\Power Origination and structuring\Structure_Toolv1.xlsm')
    OutputSht = sb.sheets['SparkSpread-Daily']    
    VecSht=sb.sheets('SparkSpread-Daily').range('A9').expand().options(pd.DataFrame).value    

    powerperiod= sb.sheets('SparkSpread-Daily').range('N10').value  
    heatrate= OutputSht.range('J3').value
    carbonrate= OutputSht.range('J4').value
    baseprice=VecSht['Base EU/MWh'].astype(float)
    Peakprice=VecSht['Peak EU/MWh'].astype(float)
    offprice=VecSht['OP EU/MWh'].astype(float)
    gasprice=VecSht['GAS EU/MWh'].astype(float) / heatrate
    carbonprice=VecSht['CO2 EU/ton'].astype(float) * carbonrate
    timemat= VecSht['T'].astype(float)
    volpower= VecSht['VOL_GAS'].astype(float)
    volgas= VecSht['VOL_PWR'].astype(float)
    correlvec= VecSht['Correlation POWER vs "GAS + CO2"'].astype(float)
    capfactor=VecSht['Capacity_Factor'].astype(float)
    #callPut= VecSht['CallPut'].astype(float)
    callPut=OutputSht.range('Q10').value    
    intrate=VecSht['Interest_rate'].astype(float)
    strike=VecSht['Strike'].astype(float) 
    gasprice = carbonprice + gasprice
    
    #if powerperiod =='OP':
    #    powerprice = offprice
    #if powerperiod =='Peak':
    #    powerprice = Peakprice
    #if powerperiod =='Base':
    powerprice = baseprice
    
    optval = KirkModel(gasprice, powerprice , carbonprice, timemat, intrate, volpower, volgas, correlvec, callPut)# * capfactor
    deltapwr = delta(gasprice, powerprice , carbonprice, timemat, intrate, volpower, volgas, correlvec, callPut)# * capfactor
    deltagas = delta2(gasprice, powerprice , carbonprice, timemat, intrate, volpower, volgas, correlvec, callPut)# * capfactor
    vegapwr = vega(gasprice, powerprice , carbonprice, timemat, intrate, volpower, volgas, correlvec, callPut)# * capfactor
    vegagas = vega2(gasprice, powerprice , carbonprice, timemat, intrate, volpower, volgas, correlvec, callPut)# * capfactor
    gammapwr = gamma(gasprice, powerprice , carbonprice, timemat, intrate, volpower, volgas, correlvec, callPut)# * capfactor
    gammagas = gamma2(gasprice, powerprice , carbonprice, timemat, intrate, volpower, volgas, correlvec, callPut)# * capfactor
   
    
    intrinsic = powerprice - gasprice - carbonprice
    extrinsic = optval - intrinsic
    
    frames = [optval, intrinsic, extrinsic, deltapwr, deltagas, vegapwr, vegagas, gammapwr,  gammagas]    
    dfData = pd.concat(frames, axis=1, sort=False)
    dfData=dfData.set_index(0)    
    
    OutputSht.range('R9').value =  dfData
    OutputSht.range('R9').value = 'FV'
    OutputSht.range('S9').value = 'Intrinsic'
    OutputSht.range('T9').value = 'Extrinsic'
    OutputSht.range('U9').value = 'Delta Pwr'
    OutputSht.range('V9').value = 'Delta Gas'
    OutputSht.range('W9').value = 'Vega Pwr'
    OutputSht.range('X9').value = 'Vega Gas'
    OutputSht.range('Y9').value = 'Gamma Pwr'
    OutputSht.range('Z9').value = 'Gamma Gas'
 
def AsianOpt_Curran():
    sb = xw.Book.caller()   
    #sb = xw.Book(r'C:\Users\daple\OneDrive\Documents\Quant Structuring Energy\Power Origination and structuring\Structure_Toolv1.xlsm')
    #ControlSht = sb.sheets['Control']
    OutputSht = sb.sheets['AsianOpt']    
    PowerOptSht=sb.sheets('AsianOpt').range('A9').expand().options(pd.DataFrame).value    
    price=PowerOptSht['Base EU/MWh'].astype(float) 
    strike=PowerOptSht['Strike'].astype(float) 
    vols=PowerOptSht['VOL POWER'].astype(float)
    dt=PowerOptSht['T'].astype(float)    
    Days=PowerOptSht['Num_Days'].astype(float)
    powerperiod=PowerOptSht['Peak_OffPeak'].astype(str)
    capacity=PowerOptSht['Capacity_Factor'].astype(float)
    int_rate=PowerOptSht['Interest_rate'].astype(float)
    #CallPut=PowerOptSht['CallPut'].astype(float)
    CallPut=OutputSht.range('K10').value
    #simulations=PowerOptSht['Simulation'].astype(float)
    
    #Review logic
    #if powerperiod =='OP':
    #    powerprice = price * 0.35
    #if powerperiod =='Peak':
    #    powerprice = price * 0.65
    #if powerperiod =='Base':
    powerprice = price 
        
    optval = Asian_Price(powerprice, strike, dt, Days, int_rate, vols, CallPut)
    deltaval = Asian_Delta(powerprice, strike, dt, Days, int_rate, vols, CallPut)    
    vegaval = Asian_Vega(powerprice, strike, dt, Days, int_rate, vols, CallPut)    
    gammaval = Asian_Gamma(powerprice, strike, dt, Days, int_rate, vols, CallPut)
    intrinsic = powerprice - strike
    extrinsic = optval - intrinsic    
    frames = [optval, intrinsic, extrinsic, deltaval, vegaval,gammaval]    
    dfData = pd.concat(frames, axis=1, sort=False) 
    dfData=dfData.set_index(0)    
    OutputSht.range('L9').value =  dfData
    OutputSht.range('L9').value = 'FV'
    OutputSht.range('M9').value = 'Intrinsic'
    OutputSht.range('N9').value = 'Extrinsic'
    OutputSht.range('O9').value = 'Delta Pwr'
    OutputSht.range('P9').value = 'Vega Pwr'
    OutputSht.range('Q9').value = 'Gamma Pwr'
    
    
def Asian_Collar(): 
    sb = xw.Book.caller()    
    #sb = xw.Book(r'C:\Users\daple\OneDrive\Documents\Quant Structuring Energy\Power Origination and structuring\Structure_Toolv1.xlsm')
    OutputSht = sb.sheets['AsianOpt-Collar']     
    PowerOptSht=sb.sheets('AsianOpt-Collar').range('A9').expand().options(pd.DataFrame).value  
    price=PowerOptSht['Base EU/MWh'].astype(float) 
    strikecall=PowerOptSht['Strike_1'].astype(float) 
    strikeput=PowerOptSht['Strike_2'].astype(float)         
    vols=PowerOptSht['VOL POWER'].astype(float)
    dt=PowerOptSht['T'].astype(float)    
    Days=PowerOptSht['Num_Days'].astype(float)
    powerperiod=PowerOptSht['Peak_OffPeak'].astype(str)
    capacity=PowerOptSht['Capacity_Factor'].astype(float)
    int_rate=PowerOptSht['Interest_rate'].astype(float)
    #Call=PowerOptSht['Call'].astype(float)
    #Put=PowerOptSht['Put'].astype(float)
    Call=OutputSht.range('L10').value
    Put=OutputSht.range('M10').value
    #simulations=PowerOptSht['Simulation'].astype(float)
    
    #if powerperiod =='OP':
    #    powerprice = price * 0.35
    #if powerperiod =='Peak':
    #    powerprice = price * 0.65
    #if powerperiod =='Base':
    powerprice = price 
    
    optval_call = Asian_Price(powerprice, strikecall, dt, Days, int_rate, vols, Call)
    optval_put = Asian_Price(powerprice, strikeput, dt, Days, int_rate, vols, Put)
    
    deltaval_call = Asian_Delta(powerprice, strikecall, dt, Days, int_rate, vols, Call)    
    deltaval_put = Asian_Delta(powerprice, strikeput, dt, Days, int_rate, vols, Put)    
    vegaval_call = Asian_Vega(powerprice, strikecall, dt, Days, int_rate, vols, Call)    
    vegaval_put = Asian_Vega(powerprice, strikeput, dt, Days, int_rate, vols, Put)        
    gammaval_call = Asian_Gamma(powerprice, strikecall, dt, Days, int_rate, vols, Call)
    gammaval_put = Asian_Gamma(powerprice, strikeput, dt, Days, int_rate, vols, Put)
    
    intrinsic_call = powerprice - strikecall
    intrinsic_put = powerprice - strikeput
    
    extrinsic_call = optval_call - intrinsic_call
    extrinsic_put = optval_put - intrinsic_put    
    
    frames = [optval_call, optval_put, intrinsic_call, intrinsic_put, extrinsic_call, extrinsic_put 
              , deltaval_call, deltaval_put, vegaval_call, vegaval_put, gammaval_call, gammaval_put]    
    dfData = pd.concat(frames, axis=1, sort=False)
    dfData=dfData.set_index(0)    
    OutputSht.range('N9').value =  dfData
    OutputSht.range('N9').value = 'FV_1'
    OutputSht.range('O9').value = 'FV_2'    
    OutputSht.range('P9').value = 'Intrinsic_1'
    OutputSht.range('Q9').value = 'Intrinsic_2'    
    OutputSht.range('R9').value = 'Extrinsic_1'
    OutputSht.range('S9').value = 'Extrinsic_2'    
    OutputSht.range('T9').value = 'Delta_1'
    OutputSht.range('U9').value = 'Delta_2'    
    OutputSht.range('V9').value = 'Vega_1'
    OutputSht.range('W9').value = 'Vega_2'    
    OutputSht.range('X9').value = 'Gamma_1'
    OutputSht.range('Y9').value = 'Gamma_2'
    #OutputSht.range('Z10').value = -optval_call + optavl_puth
    
    
def Asian_Enhanced_Collar(): 
    sb = xw.Book.caller()   
    #sb = xw.Book(r'C:\Users\daple\OneDrive\Documents\Quant Structuring Energy\Power Origination and structuring\Structure_Toolv1.xlsm')
    #ControlSht = sb.sheets['Control']
    OutputSht = sb.sheets['AsianOpt-EnhancedCollar']    
    PowerOptSht=sb.sheets('AsianOpt-EnhancedCollar').range('A9').expand().options(pd.DataFrame).value
    price=PowerOptSht['Base EU/MWh'].astype(float) 
    strikecall=PowerOptSht['Strike_1'].astype(float) 
    strikeputh=PowerOptSht['Strike_2'].astype(float)         
    strikeputl=PowerOptSht['Strike_3'].astype(float)         
    vols=PowerOptSht['VOL POWER'].astype(float)
    dt=PowerOptSht['T'].astype(float)    
    Days=PowerOptSht['Num_Days'].astype(float)
    powerperiod=PowerOptSht['Peak_OffPeak'].astype(str)
    capacity=PowerOptSht['Capacity_Factor'].astype(float)
    int_rate=PowerOptSht['Interest_rate'].astype(float)
    #Call=PowerOptSht['Call'].astype(float)
    #Puth=PowerOptSht['Put_H'].astype(float)
    #Putl=PowerOptSht['Put_L'].astype(float)
    Call=OutputSht.range('M10').value
    Puth=OutputSht.range('N10').value
    Putl=OutputSht.range('O10').value    
    
    #simulations=PowerOptSht['Simulation'].astype(float)
    
    #if powerperiod =='OP':
    #    powerprice = price * 0.35
    #if powerperiod =='Peak':
    #    powerprice = price * 0.65
    #if powerperiod =='Base':
    powerprice = price 
    
    optval_call = Asian_Price(powerprice, strikecall, dt, Days, int_rate, vols, Call)
    optval_puth = Asian_Price(powerprice, strikeputh, dt, Days, int_rate, vols, Puth)
    optval_putl = Asian_Price(powerprice, strikeputl, dt, Days, int_rate, vols, Putl)    
    
    deltaval_call = Asian_Delta(powerprice, strikecall, dt, Days, int_rate, vols, Call)    
    deltaval_puth = Asian_Delta(powerprice, strikeputh, dt, Days, int_rate, vols, Puth)
    deltaval_putl = Asian_Delta(powerprice, strikeputl, dt, Days, int_rate, vols, Putl)
    
    vegaval_call = Asian_Vega(powerprice, strikecall, dt, Days, int_rate, vols, Call)    
    vegaval_puth = Asian_Vega(powerprice, strikeputh, dt, Days, int_rate, vols, Puth)
    vegaval_putl = Asian_Vega(powerprice, strikeputl, dt, Days, int_rate, vols, Putl)
      
    gammaval_call = Asian_Gamma(powerprice, strikecall, dt, Days, int_rate, vols, Call)
    gammaval_puth = Asian_Gamma(powerprice, strikeputh, dt, Days, int_rate, vols, Puth)
    gammaval_putl = Asian_Gamma(powerprice, strikeputl, dt, Days, int_rate, vols, Putl)    
    
    intrinsic_call = powerprice - strikecall
    intrinsic_puth = powerprice - strikeputh
    intrinsic_putl = powerprice - strikeputl    
    
    extrinsic_call = optval_call - intrinsic_call
    extrinsic_puth = optval_puth - intrinsic_putl    
    extrinsic_putl = optval_putl - intrinsic_putl     
    
    frames = [optval_call, optval_puth, optval_putl, intrinsic_call, intrinsic_puth, intrinsic_putl, extrinsic_call, extrinsic_puth, extrinsic_putl 
              , deltaval_call, deltaval_puth,deltaval_putl, vegaval_call, vegaval_puth, vegaval_putl, gammaval_call, gammaval_puth, gammaval_putl]    
    dfData = pd.concat(frames, axis=1, sort=False)
    dfData=dfData.set_index(0)     
    OutputSht.range('P9').value =  dfData
    OutputSht.range('P9').value = 'FV_1'
    OutputSht.range('Q9').value = 'FV_2'  
    OutputSht.range('R9').value = 'FV_3'    
    OutputSht.range('S9').value = 'Intrinsic_1'
    OutputSht.range('T9').value = 'Intrinsic_2' 
    OutputSht.range('U9').value = 'Intrinsic_3'    
    OutputSht.range('V9').value = 'Extrinsic_1'
    OutputSht.range('W9').value = 'Extrinsic_2' 
    OutputSht.range('X9').value = 'Extrinsic_3'    
    OutputSht.range('Y9').value = 'Delta_1'
    OutputSht.range('Z9').value = 'Delta_2'    
    OutputSht.range('AA9').value = 'Delta_3'    
    OutputSht.range('AB9').value = 'Vega_1'
    OutputSht.range('AC9').value = 'Vega_2'    
    OutputSht.range('AD9').value = 'Vega_3'    
    OutputSht.range('AE9').value = 'Gamma_1'
    OutputSht.range('AF9').value = 'Gamma_2'    
    OutputSht.range('AG9').value = 'Gamma_3'    
    #OutputSht.range('AH9').value = -optval_call + optavl_puth - optval_putl
       

        
@xw.func
def KirkModel(S1_t, S2_t, K, T, r, vol1, vol2, rho, CallPut):
         
    z = S1_t / (S1_t + K * np.exp(-1. * r * T))
    vol = np.sqrt( vol1 ** 2 * z ** 2 + vol2 ** 2 - 2 * rho* vol1 * vol2 * z )
    d1 = (np.log(S2_t / (S1_t + K * np.exp(-r * T)))
    	/ (vol * np.sqrt(T)) + 0.5 * vol * np.sqrt(T))
    d2 = d1 - vol * np.sqrt(T)
    
    #N1 = 0.5 * (1 + erf(d1 / np.sqrt(2)))
    #N2 = 0.5 * (1 + erf(d2 / np.sqrt(2)))
    
    
    if CallPut == 1 or CallPut == 'call' :
        price = ((S2_t  * stats.norm.cdf(d1,0.0,1.0) - (S1_t + K * np.exp(-r * T)) * stats.norm.cdf(d2,0.0,1.0)))
    else:
        price = ((S1_t + K * np.exp(-r * T)) * stats.norm.cdf(-d2,0.0,1.0)) - (S2_t  * stats.norm.cdf(-d1,0.0,1.0))     
    return price

@xw.func
def delta(S1_t, S2_t, K, T, r, vol1, vol2, rho, CallPut): 
    diff = S1_t * 0.01        
    myCall_1 = KirkModel(S1_t + diff,S2_t, K, T, r, vol1, vol2, rho, CallPut)
    myCall_2 = KirkModel(S1_t - diff,S2_t, K, T, r, vol1, vol2, rho, CallPut)                          
    return (myCall_1 - myCall_2) / (2. * diff) 


@xw.func
def delta2(S1_t, S2_t, K, T, r, vol1, vol2, rho, CallPut): 
    diff = S2_t * 0.01        
    myCall_1 = KirkModel(S1_t,S2_t + diff, K, T, r, vol1, vol2, rho, CallPut)
    myCall_2 = KirkModel(S1_t,S2_t - diff, K, T, r, vol1, vol2, rho, CallPut)                          
    return (myCall_1 - myCall_2) / (2. * diff)

@xw.func
def gamma(S1_t, S2_t, K, T, r, vol1, vol2, rho, CallPut): 
    diff = S1_t * 0.01
    myCall_1 = KirkModel(S1_t + diff,S2_t, K, T, r, vol1, vol2, rho, CallPut)
    myCall_2 = KirkModel(S1_t - diff,S2_t, K, T, r, vol1, vol2, rho, CallPut) 
    return (myCall_1 - myCall_2) / (2. * diff)


@xw.func
def gamma2(S1_t, S2_t, K, T, r, vol1, vol2, rho, CallPut): 
    diff = S2_t * 0.01
    myCall_1 = KirkModel(S1_t,S2_t + diff, K, T, r, vol1, vol2, rho, CallPut)
    myCall_2 = KirkModel(S1_t,S2_t - diff, K, T, r, vol1, vol2, rho, CallPut) 
    return (myCall_1 - myCall_2) / (2. * diff)


@xw.func
def vega(S1_t, S2_t, K, T, r, vol1, vol2, rho, CallPut):
    diff = vol1 * 0.01
    myCall_1 = KirkModel(S1_t,S2_t, K, T, r, vol1 + diff, vol2, rho, CallPut)
    myCall_2 = KirkModel(S1_t,S2_t, K, T, r, vol1 - diff, vol2, rho, CallPut) 
    return (myCall_1 - myCall_2) / (2. * diff)


@xw.func
def vega2(S1_t, S2_t, K, T, r, vol1, vol2, rho, CallPut):
    diff = vol2 * 0.01
    myCall_1 = KirkModel(S1_t,S2_t, K, T, r, vol1,vol2 + diff, rho, CallPut)
    myCall_2 = KirkModel(S1_t,S2_t, K, T, r, vol1, vol2 - diff, rho, CallPut) 
    return (myCall_1 - myCall_2) / (2. * diff)




    
def Black76_Option(S0, K, T, r, sigma, CallPut):

    d1 = ((np.log(S0 / K) + (r + 0.5 * sigma ** 2) * T)/(sigma * np.sqrt(T)))
    d2 = ((np.log(S0 / K) + (r - 0.5 * sigma ** 2) * T)/(sigma * np.sqrt(T)))    
        
    if CallPut == 1 or CallPut == 'call' :
        value = (S0 * stats.norm.cdf(d1,0.0,1.0) - K * np.exp(- r * T) * stats.norm.cdf(d2,0.0,1.0))
    else:
        value = (K * np.exp(-r * T) * stats.norm.cdf(-d2,0.0,1.0) - S0 * stats.norm.cdf(-d1,0.0,1.0))
    return value
    

def deltabs(S0, K, T, r, sigma, CallPut):
    diff = S0 * 0.01 
    myCall_1 = Black76_Option(S0 + diff, K, T, r, sigma, CallPut)
    myCall_2 = Black76_Option(S0 - diff, K, T, r, sigma, CallPut)
    return (myCall_1 - myCall_2) / (2.0 * diff) 
    

def vegabs(S0, K, T, r, sigma, CallPut):
    diff = sigma * 0.01
    myCall_1 = Black76_Option(S0, K, T, r, sigma + diff, CallPut)
    myCall_2 = Black76_Option(S0, K, T, r, sigma - diff, CallPut)
    return (myCall_1 - myCall_2) / (2.0 * diff)
       
    
def gammabs(S0, K, T, r, sigma, CallPut):
    diff = S0 * 0.01
    myCall_1 = Black76_Option(S0 + diff, K, T, r, sigma, CallPut)
    myCall_2 = Black76_Option(S0 - diff, K, T, r, sigma, CallPut)
    return (myCall_1 - myCall_2) / (2.0 * diff)
    
    
def thetabs(S0, K, T, r, sigma, CallPut):
    diff = 1.0 / 365
    myCall_1 = Black76_Option(S0, K, T + diff, r, sigma, CallPut)
    myCall_2 = Black76_Option(S0, K, T - diff, r, sigma, CallPut)
    return (myCall_2 - myCall_1) / (2.0 * diff)
    

def rhobs(S0, K, T, r, sigma, CallPut):
    diff = r * 0.01 
    myCall_1 = Black76_Option(S0, K, T, r + diff, sigma, CallPut)
    myCall_2 = Black76_Option(S0, K, T, r - diff, sigma, CallPut)
    return (myCall_1 - myCall_2) / (2.0 * diff)
        

def Asian_Price(S0, strike, T, M, r, sigma, CallPut):
    discount = np.exp(- r * T)
    sigsqT = ((sigma ** 2 * T * (M + 1) * (2 * M + 1))
              / (6 * M * M))
    muT = (0.5 * sigsqT + (r - 0.5 * sigma ** 2)
            * T * (M + 1) / (2 * M))
    d1 = ((np.log(S0 / strike) + (muT + 0.5 * sigsqT))
          / np.sqrt(sigsqT))
    d2 = d1 - np.sqrt(sigsqT)
    N1 = 0.5 * (1 + erf(d1 / np.sqrt(2)))
    N2 = 0.5 * (1 + erf(d2 / np.sqrt(2)))
    
    if CallPut == 1 or CallPut == 'call' :
        geometric_value = discount * (S0 * np.exp(muT) * N1 - strike * N2)
    else:   
        geometric_value = discount * (S0 * np.exp(muT) * N1 - strike * N2) + strike * np.exp(-r*(T)) - S0
    
    return geometric_value    

@xw.func
def Asian_Delta(S0, strike, T, M, r, sigma, CallPut):
    myCall_1 = Asian_Price(S0 + 0.01, strike, T, M, r, sigma, CallPut)
    myCall_2 = Asian_Price(S0 - 0.01, strike, T, M, r, sigma, CallPut)
    return ((myCall_1 - myCall_2) / 0.02)


@xw.func
def Asian_Vega(S0, strike, T, M, r, sigma, CallPut):
    myCall_1 = Asian_Price(S0, strike, T, M, r, sigma + 0.01, CallPut)
    myCall_2 = Asian_Price(S0, strike, T, M, r, sigma - 0.01, CallPut)
    return ((myCall_1 - myCall_2) / 0.02)/100

@xw.func
def Asian_Gamma(S0, strike, T, M, r, sigma, CallPut):
    myCall_1 = Asian_Price(S0 + 0.01, strike, T, M, r, sigma, CallPut)
    myCall_2 = Asian_Price(S0 - 0.01, strike, T, M, r, sigma, CallPut)
    optionval = Asian_Price(S0, strike, T, M, r, sigma, CallPut)
    return ((myCall_1 + myCall_2) - 2.0 * optionval) * 10000

@xw.func
def Asian_Theta(S0, strike, T, M, r, sigma, CallPut):
    diff = 1.0 / 365.0
    myCall_1 = Asian_Price(S0, strike, T, M, r, sigma, CallPut)
    myCall_2 = Asian_Price(S0, strike, T, M, r, sigma, CallPut)
    return (myCall_1 - myCall_2)/diff/100 
    
 
    
    

# Add lSMC for American Option/Swing, Barrier/Accumulator and choices of Stochastic process for Monte-Carlo Option pricers 
    

