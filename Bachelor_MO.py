import pandas as pd
import statsmodels.api as sm
import numpy as np
import scipy.stats as sp

trade_list = pd.read_excel(r'/Users/mirza/Desktop/BA/Inputfile.xlsx', sheet_name='Bereinigte Musterdaten')
SPI_returns = pd.read_excel(r'/Users/mirza/Desktop/BA/Inputfile.xlsx', sheet_name='SPI')
Company_returns = pd.read_excel(r'/Users/mirza/Desktop/BA/Inputfile.xlsx', sheet_name='Aktienrendite')

time_frame = [30, 10, 5, 3]
CAAR_buy_list = []
CAAR_sell_list = []
T_test_buy = []
T_test_sell = []
Rank_test_buy = []
Rank_test_sell = []
CAAR_buy_list2 = []
CAAR_sell_list2 = []
T_test_buy2 = []
T_test_sell2 = []
Rank_test_buy2 = []
Rank_test_sell2 = []

for t in time_frame:
    CARi_list_buy = []
    CARi_list_sell = []
    CARi_list_buy2 = []
    CARi_list_sell2 = []

    for index, element in trade_list.iterrows():
        date = element["date"]
        sec = element["sec"]
        direction = element["direction"]
        amount = element["amount"]
        executor = element["executor"]
        size = element["size"]
        sector = element["sector"]

        # get the security returns and select the last 250 items starting from the time delta
        data_security = (Company_returns[sec].loc[(Company_returns.date <= date)].tail(281))
        data_security = data_security.iloc[:-31]

        # get the 250 days of the market
        market_data = (SPI_returns["SPI"].loc[(Company_returns.date <= date)].tail(281))
        market_data = market_data.iloc[:-31]
        data_market = sm.add_constant(market_data)
        regression = sm.OLS(data_security, data_market).fit()

        alpha = regression.params[0]
        beta = regression.params.get('SPI')

        if direction == 'Buy':
            # calculate Abnormal Returns
            r_sec = (Company_returns[sec].loc[(Company_returns.date < date)].tail(t)) # pre-event
            r_sec2 = (Company_returns[sec].loc[(Company_returns.date > date)].head(t)) # post-event
            r_mkt = (SPI_returns["SPI"].loc[(Company_returns.date < date)].tail(t)) # pre-event
            r_mkt2 = (SPI_returns["SPI"].loc[(Company_returns.date > date)].head(t)) # post-event

            ARi = r_sec - alpha - beta * r_mkt # pre-event
            ARi2 = r_sec2 - alpha - beta * r_mkt2 # post-event

            # calculate Cumulatice Abnormal Returns
            CARi = sum(ARi) # pre-event
            CARi_list_buy.append(CARi) # post-event
            CARi2 = sum(ARi2) # pre-event
            CARi_list_buy2.append(CARi2) # post-event

        else:
            # calculate Abnormal Returns
            r_sec = (Company_returns[sec].loc[(Company_returns.date < date)].tail(t)) # pre-event
            r_sec2 = (Company_returns[sec].loc[(Company_returns.date > date)].head(t)) # post-event
            r_mkt = (SPI_returns["SPI"].loc[(Company_returns.date < date)].tail(t)) # pre-event
            r_mkt2 = (SPI_returns["SPI"].loc[(Company_returns.date > date)].head(t)) # post-event

            ARi = r_sec - alpha - beta * r_mkt # pre-event
            ARi2 = r_sec2 - alpha - beta * r_mkt2 # post-event

            # calculate Cumulatice Abnormal Returns
            CARi = sum(ARi) # pre-event
            CARi_list_sell.append(CARi) # pre-event
            CARi2 = sum(ARi2) # post-event
            CARi_list_sell2.append(CARi2) # post-event

    # Calculate Cumulative Average Abnormal Return
    CAAR_buy = np.mean(CARi_list_buy)
    CAAR_buy_list.append(CAAR_buy)
    CAAR_buy2 = np.mean(CARi_list_buy2)
    CAAR_buy_list2.append(CAAR_buy2)

    CAAR_sell = np.mean(CARi_list_sell)
    CAAR_sell_list.append(CAAR_sell)
    CAAR_sell2 = np.mean(CARi_list_sell2)
    CAAR_sell_list2.append(CAAR_sell2)

    # calculate Z1 for buy
    delta_buy = 0
    for i in CARi_list_buy:
        delta_buy += (i - CAAR_buy) ** 2

    sd_CAAR_buy = CAAR_buy / np.sqrt((1 / len(CARi_list_buy)) *(len(CARi_list_buy) - 1) * delta_buy)

    Z1_buy = CAAR_buy / sd_CAAR_buy
    T_test_buy.append(Z1_buy)

    delta_buy2 = 0
    for i in CARi_list_buy2:
        delta_buy2 += (i - CAAR_buy2) ** 2

    sd_CAAR_buy2 = CAAR_buy2 / np.sqrt((1 / len(CARi_list_buy2)) *(len(CARi_list_buy2) - 1) * delta_buy2)

    Z1_buy2 = CAAR_buy2 / sd_CAAR_buy2
    T_test_buy2.append(Z1_buy2)

    # calculate Z1 for sell
    delta_sell = 0
    for i in CARi_list_sell:
        delta_sell += (i - CAAR_sell) ** 2

    sd_CAAR_sell = CAAR_sell / np.sqrt((1 / len(CARi_list_sell)) *(len(CARi_list_sell) - 1) * delta_sell)

    Z1_sell = CAAR_sell / sd_CAAR_sell
    T_test_sell.append(Z1_sell)

    delta_sell2 = 0
    for i in CARi_list_sell2:
        delta_sell2 += (i - CAAR_sell2) ** 2

    sd_CAAR_sell2 = CAAR_sell2 / np.sqrt((1 / len(CARi_list_sell2)) *(len(CARi_list_sell2) - 1) * delta_sell2)

    Z1_sell2 = CAAR_sell2 / sd_CAAR_sell2
    T_test_sell2.append(Z1_sell2)

summary1 = pd.DataFrame([CAAR_buy_list, T_test_buy, Rank_test_buy, CAAR_sell_list, T_test_sell, Rank_test_sell])
summary2 = pd.DataFrame([CAAR_buy_list2, T_test_buy2, Rank_test_buy2, CAAR_sell_list2, T_test_sell2, Rank_test_sell2])
summary2 = summary2.reindex(columns = [3, 2, 1, 0])

results = pd.concat([summary1, summary2], axis=1)
print(results)
results.to_excel('output.xlsx', index=False)

    #  calculate Z2 for buy
    # Wilcoxon_buy = sp.wilcoxon(ARi_list_buy[1])
    # p_value_buy = Wilcoxon_buy.pvalue

    # Z2_buy = sp.norm.ppf(1 - p_value_buy/2)

    # calculate Z2 for sell
    # Wilcoxon_sell = sp.wilcoxon(CARi_list_sell)
    # p_value_sell = Wilcoxon_sell.pvalue

    # Z2_sell = sp.norm.ppf(1 - p_value_sell/2)



