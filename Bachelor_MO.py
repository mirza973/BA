import pandas as pd
import statsmodels.api as sm
import numpy as np
import scipy.stats as sp

trade_list = pd.read_excel(r'/Users/mirza/Desktop/BA/Inputfile.xlsx', sheet_name='Bereinigte Musterdaten')
SPI_returns = pd.read_excel(r'/Users/mirza/Desktop/BA/Inputfile.xlsx', sheet_name='SPI')
Company_returns = pd.read_excel(r'/Users/mirza/Desktop/BA/Inputfile.xlsx', sheet_name='Aktienrendite')

CARi_list_buy = []
CARi_list_sell = []

for index, element in trade_list.iterrows():
    date = element["date"]
    sec = element["sec"]
    direction = element["direction"]
    amount = element["amount"]
    executor = element["executor"]
    size = element["size"]
    sector = element["sector"]

    # calculate time deltas in format datetime64[ns]
    dateMinusDelta = date - pd.Timedelta(days=31)  # ***wie müsste man das schreiben, um nicht Tage sondern Elemente zu berücksichtigen***

    # get the security returns and select the last 250 items starting from the time delta
    data_security = (Company_returns[sec].loc[(Company_returns.date <= dateMinusDelta)].tail(250))

    # get the 250 days of the market
    market_data = (SPI_returns["SPI"].loc[(Company_returns.date <= dateMinusDelta)].tail(250))
    data_market = sm.add_constant(market_data)
    regression = sm.OLS(data_security, data_market).fit()

    alpha = regression.params[0]
    beta = regression.params.get('SPI')

    if direction == 'Buy':
        # calculate Abnormal Returns
        r_sec = (Company_returns[sec].loc[(Company_returns.date <= date)].tail(30))  # ***ist um einen Tag versetzt (letzer Run ergibt 513 anstatt 514 und date wäre 515)****
        r_mkt = (SPI_returns["SPI"].loc[(Company_returns.date <= date)].tail(30))
        ARi = r_sec - alpha - beta * r_mkt

        # calculate Cumulatice Abnormal Returns
        CARi = sum(ARi)
        CARi_list_buy.append(CARi)

    else:
        # calculate Abnormal Returns
        r_sec = (Company_returns[sec].loc[(Company_returns.date <= date)].tail(30))  # ***ist um einen Tag versetzt (letzer Run ergibt 513 anstatt 514 und date wäre 515)****
        r_mkt = (SPI_returns["SPI"].loc[(Company_returns.date <= date)].tail(30))
        ARi = r_sec - alpha - beta * r_mkt

        # calculate Cumulatice Abnormal Returns
        CARi = sum(ARi)
        CARi_list_sell.append(CARi)

# Calculate Cumulative Average Abnormal Return
CAAR_buy = np.mean(CARi_list_buy)
CAAR_sell = np.mean(CARi_list_sell)

# calculate Z1 for buy
delta_buy = 0
for i in CARi_list_buy:
    delta_buy += (i - CAAR_buy)**2

sd_CAAR_buy = (delta_buy / (len(CARi_list_buy) * (len(CARi_list_buy) - 1))) ** 0.5  # ***FORMEL PRÜFEN (IN THEORIE)***

Z1_buy = CAAR_buy / sd_CAAR_buy

# calculate Z1 for sell
delta_sell = 0
for i in CARi_list_sell:
    delta_sell += (i - CAAR_sell)**2

sd_CAAR_sell = (delta_sell / (len(CARi_list_sell) * (len(CARi_list_sell) - 1))) ** 0.5  # ***FORMEL PRÜFEN (IN THEORIE)***

Z1_sell = CAAR_sell / sd_CAAR_sell

# calculate Z2 for buy
Wilcoxon_buy = sp.wilcoxon(CARi_list_buy)
p_value_buy = Wilcoxon_buy.pvalue

Z2_buy = sp.norm.ppf(1 - p_value_buy/2)

# calculate Z2 for sell
Wilcoxon_sell = sp.wilcoxon(CARi_list_sell)
p_value_sell = Wilcoxon_sell.pvalue

Z2_sell = sp.norm.ppf(1 - p_value_sell/2)

