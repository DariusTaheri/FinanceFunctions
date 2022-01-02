# -*- coding: utf-8 -*-
"""
Created on Thu Dec  9 21:13:38 2021

@author: DariusPC
"""
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from statsmodels.tsa.stattools import coint, adfuller
import yfinance as yf
from MyFunctions import dfinfo

'''
=============================================================================
IMPROVEMENTS:
    - Use Plotly instead of Matplotlib (Reuse code from Work USDCAD NN Model)
    - Pick Beta that represents the average diversion (trade time)
    - Refactor equity chart as an object
    - Include fees and Bid-Ask Spreads


=============================================================================
'''

# Script Variables
pd.set_option('display.max_rows', 10000)
pd.set_option('display.max_columns', 15)
pd.set_option('display.width', 1000)
pd.set_option('max_seq_items', 2000)
import warnings
warnings.filterwarnings('ignore')

def yfdatapull(S1,S2,interval,start,windows):
    

    df = yf.download(tickers=[S1,S2],interval = interval, start=start)
    df.reset_index(inplace=True)
    df[f'{S1}'] = df[('Adj Close',f'{S1}')]
    df[f'{S2}'] = df[('Adj Close',f'{S2}')]
    
    try: 
        df = df[['Date',f'{S2}',f'{S1}']]
    except KeyError: 
        df['Date'] = df['index'] 
        df = df[['Date',f'{S2}',f'{S1}']]
        
    df.columns = df.columns.droplevel(1)
    
    df[f'log{S1}'] = np.log(df[f'{S1}'])
    df[f'log{S2}'] = np.log(df[f'{S2}'])
    
    #Calculating log spreads and z-score
    df['LogSpread'] = df[f'log{S1}'] - df[f'log{S2}']
    df["LogSprdZScore"] = (df["LogSpread"]- df["LogSpread"].mean())/df["LogSpread"].std()
    
    #Calculating Beta (non rolling) & Beta Log Spread
    StaticBeta = (df[f'log{S1}'].cov(df[f'log{S2}']))/df[f'log{S2}'].var()
    df['StaticBetaSpread'] = df[f'log{S1}'] - (StaticBeta*df[f'log{S2}'])
    df['StaticBetaLogSprdZScore'] = (df['StaticBetaSpread']- df['StaticBetaSpread'].mean())/df['StaticBetaSpread'].std()
    
    #Calculating rolling betas and their spreads and z scores
    def rollingwrapper(df, windows):
        for x in windows:
            df[f'{x}-RollingLogBeta'] = (df[f'log{S1}'].rolling(x).cov(df[f'log{S2}']))/df[f'log{S2}'].rolling(x).var()
            df[f'{x}-RollingLogBetaSpread'] = df[f'log{S1}'] - (df[f'{x}-RollingLogBeta']*df[f'log{S2}'])
            df[f'{x}-RollingLogSprdZScore'] = (df[f'{x}-RollingLogBetaSpread']- df[f'{x}-RollingLogBetaSpread'].mean())/df[f'{x}-RollingLogBetaSpread'].std()
        return df
    df = rollingwrapper(df,windows)
    
    #Cleaning up index
    df = df.dropna()
    df.reset_index(inplace=True, drop=True)
    
    return df

def dfstats(df,windows):
    print ('\n---Correlation---')
    print(np.corrcoef(df[f'log{S1}'], df[f'log{S2}']))
    print(f'Static Beta: {StaticBeta.round(2)}')
 
    
 
    # =============================================================================
    # Stationarity Function
    # =============================================================================
    def stationarity_test(data, cutoff=0.05):
    # H_0 in adfuller is unit root exists (non-stationary)
    # We must observe significant p-value to convince ourselves that the series is stationary
        pvalue = adfuller(data)[1]
        if pvalue < cutoff:
            print(f'p-value = {pvalue}. The series {data.name} is likely stationary.')
        else:
            print(f'p-value = {pvalue}. The series {data.name} is likely NON-stationary.')    
 
            
    #Conintegration Function
    def Cointfunc(x,y,cutoff=0.01):
        score, pvalue, _ = coint(x,y)
        stringtemp = f'{x.name} & {y.name} p-value = {pvalue.round(3)}'
        if pvalue < cutoff:
            print(stringtemp + '. The Two assets are most likely cointegrated.')
        else:
            print(stringtemp + '. The Two assets are most likely NOT cointegrated.')
    

    #print ('\nNon Beta Stat Tests')
    print('\nCo-integrated Test')
    Cointfunc(df[f'log{S1}'],df[f'log{S2}'])
    
    #Testing for Stationarity
    print('\nStationarity Tests')
    stationarity_test(df[f'log{S1}'])      
    stationarity_test(df[f'log{S2}'])  
    stationarity_test(df["LogSpread"]) 
    stationarity_test(df['StaticBetaSpread'])
    for x in windows:
        stationarity_test(df[f'{x}-RollingLogBetaSpread'])
    

def equitycharts(df,windows,ztrgt):
    '''
    If Z>2: Sell S1, Buy S2
    IF Z<2: Buy S1, Sell S2
    Exit Position when z gets to 0
    '''
    
    for x in windows:
        print (f'\nCalculating {x} Equity Curve')
        df[f'{x}-Equity'] = 0
        df['DailyPnL'] = 0
        df[f'{S1}pos'] = 0 
        df[f'{S2}pos'] = 0 
        
        tradechk = 0
        S1o = 0 
        S2o = 0
        for y in range(1,len(df)):
            contracts = 0 
            
            #Checking to sell S1, Buy S2
            if df.loc[y,f'{x}-RollingLogSprdZScore']>= ztrgt and tradechk == 0:
                tradechk = -1
                tempbeta = df.loc[y,f'{x}-RollingLogBeta']
                tempbeta = tempbeta.round(1)
                S1o = df.loc[y,f'{S1}']
                S2o = df.loc[y,f'{S2}']
                
                #Finding S1 amount
                for num in range(1,200):
                    if (tempbeta * num).is_integer() is True:
                        print ('Placing Sell S1, Buy S2: '+ str(num * -1) + '  ' + str(num * tempbeta * 1 ) + ' on: ' + str(df.loc[y,'Date']))
                        df[f'{S1}pos'] = num * -1
                        df[f'{S2}pos'] = num * tempbeta * 1
                        contracts = (num * tempbeta) + num
                        break
            
            #Checking to buy S1, Sell S2
            if df.loc[y,f'{x}-RollingLogSprdZScore']<= (ztrgt*-1) and tradechk == 0:
                tradechk = 1
                tempbeta = df.loc[y,f'{x}-RollingLogBeta']
                tempbeta = tempbeta.round(1)
                S1o = df.loc[y,f'{S1}']
                S2o = df.loc[y,f'{S2}']
               
                #Finding S1 amount
                for num in range(1,200):
                    if (tempbeta * num).is_integer() is True:
                        print ('Placing Buy S1, Sell S2: '+ str(num * 1) + '  ' + str(num * tempbeta * -1) + ' on: ' + str(df.loc[y,'Date']))
                        df[f'{S1}pos'] = num * 1
                        df[f'{S2}pos'] = num * tempbeta * -1     
                        contracts = (num * tempbeta) + num
                        break
           
            #Calculating Daily PnL

            if tradechk != 0: 
                df.loc[y,'DailyPnL'] = (((df.loc[y,f'{S1}']-df.loc[y-1,f'{S1}'])*df.loc[y,f'{S1}pos'])+
                                        ((df.loc[y,f'{S2}']-df.loc[y-1,f'{S2}'])*df.loc[y,f'{S2}pos'])
                                        #+contracts*contractcost*-1
                                        )
                                     
            
           
            #Unwinding Sell S1, Buy S2 Position
            if df.loc[y,f'{x}-RollingLogSprdZScore']<= 0.05  and tradechk == -1: 
                print ('EXITING Sell S1, Buy S2 Position')
                df[f'{S1}pos'] = 0
                df[f'{S2}pos'] = 0
                S1o = 0 
                S2o = 0
                tradechk = 0
           
            #Unwinding Buy S1, Sell S2 Position
            if df.loc[y,f'{x}-RollingLogSprdZScore']>= 0.05  and tradechk == 1: 
                print ('EXITING Buy S1, Sell S2 Position')
                df[f'{S1}pos'] = 0
                df[f'{S2}pos'] = 0
                S1o = 0 
                S2o = 0
                tradechk = 0
                
            #Calculating PnL Position
            df.loc[y,f'{x}-Equity'] = df.loc[y-1,f'{x}-Equity'] + df.loc[y,'DailyPnL']
            
            
            
            

    ''' 
    df.loc[y,f'{x}-Equity'] = (df.loc[y-1,f'{x}-Equity']+
                                ((df.loc[y,f'{S1}']-S1o)*df.loc[y,f'{S1}pos'])+
                                ((df.loc[y,f'{S2}']-S2o)*df.loc[y,f'{S2}pos'])
                                
                                +
                                (contracts*contractcost*-1)
                                )
    '''
                

    '''
        df[f'{S1}pos'] = 0 
        df[f'{S2}pos'] = 0
    '''
    df.pop(f'{S1}pos')
    df.pop(f'{S2}pos')
    
    return df
    
    

def plotdf(df, windows,ztrgt):
    
    fig, axs = plt.subplots(2, 3, sharex=True)
    
    axs[0,0].set_title("Prices") 
    axs[0,0].plot(df['Date'],df[f'{S1}'], label=f'{S1}',color='red')
    axs[0,0] = axs[0,0].twinx()
    axs[0,0].plot(df['Date'],df[f'{S2}'], label=f'{S2}',color='green')
    axs[0, 0].legend(loc='upper left')
    
    axs[1, 0].set_title('Price Spread')
    axs[1, 0].plot(df['Date'],df[f'{S1}'] - df[f'{S2}'], label='Price Spread')
    #axs[1, 0].axhline(df['LogSpread'].mean())
    #axs[1, 0].axhline(ztrgt, color='red')
    #axs[1, 0].axhline(-1*ztrgt, color='green')
    
    axs[0, 1].set_title('Rolling Betas')
    for x in windows:
        axs[0, 1].plot(df['Date'],df[f'{x}-RollingLogBeta'], label=f'{x}-Beta')
    axs[0, 1].legend(loc='upper left')
    
    axs[1, 1].set_title('Rolling Beta Spreads')
    for y in windows:
        axs[1, 1].plot(df['Date'],df[f'{y}-RollingLogBetaSpread'], label=f'{y}-Spread')
    axs[1, 1].legend(loc='upper left')
    
    axs[0, 2].set_title('Rolling Beta Z-Scores')
    for z in windows:
        axs[0, 2].plot(df['Date'],df[f'{z}-RollingLogSprdZScore'], label=f'{z}-Z')
    axs[0, 2].axhline(ztrgt, color='red')
    axs[0, 2].axhline(-1*ztrgt, color='green')
    axs[0, 2].legend(loc='upper left')
    
    axs[1, 2].set_title('PnL Graph')
    for t in windows:
        axs[1, 2].plot(df['Date'],df[f'{t}-Equity'], label=f'{t}-E')
    axs[1, 2].legend(loc='upper left')
    

    
    fig.set_size_inches(20, 10)
    fig.set_tight_layout(True)
    fig.show()

# =============== Running Script ==================
if __name__ == '__main__':
    #S1 is the dependent variable (buys vs sell) against S2
    S1 = 'ETH-USD'
    S2 = 'BTC-USD'
    interval = '1d'
    start = '2018-01-01'
    windows = [15,20,30]
    ztrgt = 2
    
    #Futures Parameters -- REFACTOR: FIX FEES 
    contractnum = 1 
    S1multi = 10
    S2multi = 10
    #CostPerContract = 0 
    #contractcost = CostPerContract * S1multi
    print('---Creating DF---')
    df = yfdatapull(S1,S2,interval,start,windows)
    StaticBeta = (df[f'log{S1}'].cov(df[f'log{S2}']))/df[f'log{S2}'].var()
    
    
    print('\n---Running Statistics---')
    dfstats(df,windows) 
    
    print ('\n---Creating Equity Graphs')
    equitycharts(df,windows,ztrgt)
    
    #Exporting the CSV
    df.to_csv(r'results.csv', index = False)
    print ('\nCSV File Made')
    
    #To-Do: Refactor to plotly 
    plotdf(df,windows,ztrgt)
    
    
    print('\n---Dataframe Information---')
    print (df.head(3))
    print (df.tail(3))
    print (df.dtypes)
    print (df.shape)
    print (df.columns)
    
    