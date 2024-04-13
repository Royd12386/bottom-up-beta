import pandas as pd
import xlwings as xw
import yfinance as yf
#-----------------------------------------------------------------------------------------------
#                                   creates list of all yahoo tickers
tickers_dataFrame = pd.DataFrame(pd.read_csv("C:/Users/cc2016/Desktop/borsa/DCF/beta/tickers_list.csv"))

list_ofAllTickers = (tickers_dataFrame.Symbol).tolist()

data={}

errors = {}
#-----------------------------------------------------------------------------------------------
#                                   creates a list of sectors and industries
industries_that_exist = []
sectors_that_exist = []
for i in list_ofAllTickers:
    try:
        ticker = yf.Ticker(i)
        if((ticker.info)['industry'] not in industries_that_exist):                                       #40 min ish
            industries_that_exist.append((ticker.info)['industry'])
        if((ticker.info)['sector'] not in sectors_that_exist):
            sectors_that_exist.append((ticker.info)['sector'])

    except:
        continue
#-----------------------------------------------------------------------------------------------
#                                       sets values to [0,0,0,0]
for i in industries_that_exist:
    data[i] = [0,0,0,0,0,0]
for i in sectors_that_exist:
    data[i] = [0,0,0,0,0,0]
#-----------------------------------------------------------------------------------------------
#                                      calculate and sets data

for i in list_ofAllTickers:
    try:
        ticker = yf.Ticker(i)
        beta = (ticker.info)['beta']
        if(beta<9) and ((-9)<beta): #filter for unwanted firms
            industry = (ticker.info)['industry']                                  
            sector = (ticker.info)['sector']
            cash = (ticker.info)['totalCash']            #0: beta  1:count 2:debt 3:equity 4:cash 5:firm value
            debt = (ticker.info)['totalDebt']
            market_value_equity = ((ticker.info)['sharesOutstanding'])*((ticker.info)['currentPrice'])                                                 #30 min ish
            firm_value = ((ticker.info)['marketCap'])+((ticker.info)['totalDebt'])- cash
            cashToFirm = cash/firm_value
            data[industry] = [data[industry][0]+beta,data[industry][1]+1,data[industry][2]+debt,data[industry][3]+market_value_equity,data[industry][4]+cash,data[industry][5]+firm_value]    
            data[sector] = [data[sector][0]+beta,data[sector][1]+1,data[sector][2]+debt,data[sector][3]+market_value_equity,data[sector][4]+cash,data[sector][5]+firm_value]
        else:
            errors[i] = beta
    except:
        continue
#-----------------------------------------------------------------------------------------------
#                               sets the values in the excel sheet

wb = xw.Book(r'C:\Users\cc2016\Desktop\borsa\DCF\beta\beta.xlsx')              #s1.range(row,column).value
s1 = wb.sheets['Sheet1']
##################################################################################
#                              industry
row_num = 5
for i in industries_that_exist:
    s1.range(row_num,2).value = i
    s1.range(row_num,3).value = data[i][1]
    try:
        s1.range(row_num,4).value = (data[i][0]/data[i][1])
    except:
        s1.range(row_num,4).value = 0
    try:
        s1.range(row_num,5).value = (data[i][2]/data[i][3])
    except:
        s1.range(row_num,5).value = 0
    try:
        s1.range(row_num,8).value = (data[i][4]/data[i][5])
    except:
        s1.range(row_num,8).value = 0
    row_num = row_num+1
##################################################################################                            #15 min ish
#                              sector
row_num = 5
for i in sectors_that_exist:
    s1.range(row_num,11).value = i
    s1.range(row_num,12).value = data[i][1]
    try:
        s1.range(row_num,13).value = (data[i][0]/data[i][1])
    except:
        s1.range(row_num,13).value = 0
    try:
        s1.range(row_num,14).value = (data[i][2]/data[i][3])
    except:
        s1.range(row_num,14).value = 0
    try:
        s1.range(row_num,17).value = (data[i][4]/data[i][5])
    except:
        s1.range(row_num,17).value = 0
    row_num = row_num+1
##################################################################################


