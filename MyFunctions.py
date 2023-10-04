import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt 
from statsmodels.tsa.stattools import coint, adfuller, kpss
import numpy as np    
import warnings

# Printing useful DataFrame info
# =============================================================================
def dfinfo(df):
    print('\n---Dataframe Information---')
    print ("    ")
    print (df.head(3))
    print ("    ")
    print (df.tail(3))
    print (df.dtypes)
    print (df.shape)
    print (df.columns)    


# =============================================================================
# Pulling YF Data for one security
# =============================================================================
def yfdatapull(ticker,interval,startdate):
    df = yf.download(tickers=ticker,interval = interval, start=startdate)
    df.reset_index(inplace=True)    
    
    return df


# =============================================================================
# Script Variables
# =============================================================================
def ScriptVars():
    print ('\nRunning Script Variables')
    pd.set_option('display.max_rows', 10000)
    pd.set_option('display.max_columns', 15)
    pd.set_option('display.width', 1000)
    pd.set_option('max_seq_items', 2000)
    warnings.filterwarnings('ignore')
    
    
# =============================================================================
# Stationarity Function
# =============================================================================
def stationarity_test(data, cutoff):
# H_0 in adfuller is unit root exists (non-stationary)
# We must observe significant p-value to convince ourselves that the series is stationary
    pvalue = adfuller(data)[1]
    if pvalue < cutoff:
        print(f'ADF p-value = {pvalue}. {data.name} is likely stationary.')
        return True
    else:
        print(f'ADF p-value = {pvalue}. {data.name} is likely NON-stationary.')
        return False


# =============================================================================
# Trend Stationarity Function
# =============================================================================
def trend_stationarity_test(timeseries,cutoff):
    kpsstest = kpss(timeseries, regression='ct')
    kpss_pvalue = kpsstest[1]
    if kpss_pvalue < cutoff:
        print(f'KPSS p-value = {kpss_pvalue}. {timeseries.name} is likely NOT TREND stationary.')
        return False
    else:
        print(f'KPSS p-value = {kpss_pvalue}. {timeseries.name} is likely TREND stationary.')
        return True

# =============================================================================
# Cointegration Function
# =============================================================================        
def CointFunc(x,y,cutoff=0.01):
    score, pvalue, _ = coint(x,y)
    stringtemp = f'{x.name} & {y.name} p-value = {pvalue.round(3)}'
    if pvalue < cutoff:
        print(stringtemp + '. The Two assets are most likely cointegrated.')
    else:
        print(stringtemp + '. The Two assets are most likely NOT cointegrated.')
