import pandas as pd
import statsmodels.api as sm
import numpy as np
import scipy.stats as sp

trade_list = pd.read_excel(r'/Users/mirza/Desktop/BA/Inputfile1.xlsx', sheet_name='Bereinigte Musterdaten')
SPI_returns = pd.read_excel(r'/Users/mirza/Desktop/BA/Inputfile1.xlsx', sheet_name='SPI')
Company_returns = pd.read_excel(r'/Users/mirza/Desktop/BA/Inputfile1.xlsx', sheet_name='Aktienrendite')
writer = pd.ExcelWriter('output.xlsx', engine='xlsxwriter')

# get the samples in trade list
total_list = trade_list["dummy"].unique()
amount_list = trade_list["amount"].unique()
size_list = trade_list["size"].unique()
executor_list = trade_list["executor"].unique()
sector_list = trade_list["sector"].unique()
subcriteria_list = np.concatenate((total_list, amount_list, size_list, executor_list, sector_list), axis=None)

# run calculations for all samples
for subcriteria in subcriteria_list:

    # set up necessary lists for later calculations
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

    # run calculations for different time frames
    for t in time_frame:
        CARi_list_buy = []
        CARi_list_sell = []
        CARi_list_buy2 = []
        CARi_list_sell2 = []

        # specification of elements in trade list
        for index, element in trade_list.iterrows():
            date = element["date"]
            sec = element["sec"]
            direction = element["direction"]
            amount = element["amount"]
            executor = element["executor"]
            size = element["size"]
            sector = element["sector"]
            dummy = element["dummy"]

            # specification of samples
            if not (dummy == subcriteria or amount == subcriteria or
                    executor == subcriteria or size == subcriteria or sector == subcriteria):
                continue

            # security returns and selection of  last 250 items starting from the transaction date
            data_security = (Company_returns[sec].loc[(Company_returns.date <= date)].tail(281))
            data_security = data_security.iloc[:-31]

            # 250 days of the market and regression parameters
            market_data = (SPI_returns["SPI"].loc[(Company_returns.date <= date)].tail(281))
            market_data = market_data.iloc[:-31]
            data_market = sm.add_constant(market_data)
            regression = sm.OLS(data_security, data_market).fit()

            alpha = regression.params[0]
            beta = regression.params.get('SPI')

            # differentiation between buy and sell
            if direction == 'Buy':
                # calculate Abnormal Returns for buy transactions
                r_sec = (Company_returns[sec].loc[(Company_returns.date < date)].tail(t))  # blank is pre-event
                r_sec2 = (Company_returns[sec].loc[(Company_returns.date > date)].head(t))  # '2' is post-event
                r_mkt = (SPI_returns["SPI"].loc[(Company_returns.date < date)].tail(t))
                r_mkt2 = (SPI_returns["SPI"].loc[(Company_returns.date > date)].head(t))

                ARi = r_sec - alpha - beta * r_mkt
                ARi2 = r_sec2 - alpha - beta * r_mkt2

                # calculate Cumulatice Abnormal Returns for buy transactions
                CARi = sum(ARi)
                CARi_list_buy.append(CARi)
                CARi2 = sum(ARi2)
                CARi_list_buy2.append(CARi2)

            else:
                # calculate Abnormal Returns for sell transactions
                r_sec = (Company_returns[sec].loc[(Company_returns.date < date)].tail(t))
                r_sec2 = (Company_returns[sec].loc[(Company_returns.date > date)].head(t))
                r_mkt = (SPI_returns["SPI"].loc[(Company_returns.date < date)].tail(t))
                r_mkt2 = (SPI_returns["SPI"].loc[(Company_returns.date > date)].head(t))

                ARi = r_sec - alpha - beta * r_mkt
                ARi2 = r_sec2 - alpha - beta * r_mkt2

                # calculate Cumulative Abnormal Returns sell transactions
                CARi = sum(ARi)
                CARi_list_sell.append(CARi)
                CARi2 = sum(ARi2)
                CARi_list_sell2.append(CARi2)

        # Calculate Cumulative Average Abnormal Return for buy and sell transactions
        CAAR_buy = np.mean(CARi_list_buy)
        CAAR_buy_list.append(CAAR_buy)
        CAAR_buy2 = np.mean(CARi_list_buy2)
        CAAR_buy_list2.append(CAAR_buy2)

        CAAR_sell = np.mean(CARi_list_sell)
        CAAR_sell_list.append(CAAR_sell)
        CAAR_sell2 = np.mean(CARi_list_sell2)
        CAAR_sell_list2.append(CAAR_sell2)

        # calculate Z1 for buy transactions pre event date
        delta_buy = 0
        for i in CARi_list_buy:
            delta_buy += (i - CAAR_buy) ** 2

        sd_CAAR_buy = CAAR_buy / np.sqrt((1 / len(CARi_list_buy)) * (len(CARi_list_buy) - 1) * delta_buy)

        Z1_buy = CAAR_buy / sd_CAAR_buy
        T_test_buy.append(Z1_buy)

        # calculate Z1 for buy transactions post event date
        delta_buy2 = 0
        for i in CARi_list_buy2:
            delta_buy2 += (i - CAAR_buy2) ** 2

        sd_CAAR_buy2 = CAAR_buy2 / np.sqrt((1 / len(CARi_list_buy2)) * (len(CARi_list_buy2) - 1) * delta_buy2)

        Z1_buy2 = CAAR_buy2 / sd_CAAR_buy2
        T_test_buy2.append(Z1_buy2)

        # calculate Z1 for sell transactions  pre event date
        delta_sell = 0
        for i in CARi_list_sell:
            delta_sell += (i - CAAR_sell) ** 2

        sd_CAAR_sell = CAAR_sell / np.sqrt((1 / len(CARi_list_sell)) * (len(CARi_list_sell) - 1) * delta_sell)

        Z1_sell = CAAR_sell / sd_CAAR_sell
        T_test_sell.append(Z1_sell)

        # calculate Z1 for sell transactions post event date
        delta_sell2 = 0
        for i in CARi_list_sell2:
            delta_sell2 += (i - CAAR_sell2) ** 2

        sd_CAAR_sell2 = CAAR_sell2 / np.sqrt((1 / len(CARi_list_sell2)) * (len(CARi_list_sell2) - 1) * delta_sell2)

        Z1_sell2 = CAAR_sell2 / sd_CAAR_sell2
        T_test_sell2.append(Z1_sell2)

        # calculate Z2 for buy transactions pre event date
        Wilcoxon_buy = sp.wilcoxon(CARi_list_buy)
        p_value_buy = Wilcoxon_buy.pvalue

        Z2_buy = sp.norm.ppf(1 - p_value_buy / 2)
        Rank_test_buy.append(Z2_buy)

        # calculate Z2 for buy transactions post event date
        Wilcoxon_buy2 = sp.wilcoxon(CARi_list_buy2)
        p_value_buy2 = Wilcoxon_buy2.pvalue

        Z2_buy2 = sp.norm.ppf(1 - p_value_buy2 / 2)
        Rank_test_buy2.append(Z2_buy2)

        # calculate Z2 for sell transactions pre event date
        Wilcoxon_sell = sp.wilcoxon(CARi_list_sell)
        p_value_sell = Wilcoxon_sell.pvalue

        Z2_sell = sp.norm.ppf(1 - p_value_sell / 2)
        Rank_test_sell.append(Z2_sell)

        # calculate Z2 for sell transactions pre event date
        Wilcoxon_sell2 = sp.wilcoxon(CARi_list_sell2)
        p_value_sell2 = Wilcoxon_sell2.pvalue

        Z2_sell2 = sp.norm.ppf(1 - p_value_sell2 / 2)
        Rank_test_sell2.append(Z2_sell2)

    s1 = pd.DataFrame([CAAR_buy_list, T_test_buy, Rank_test_buy, CAAR_sell_list, T_test_sell, Rank_test_sell],)

    s2 = pd.DataFrame([CAAR_buy_list2, T_test_buy2, Rank_test_buy2, CAAR_sell_list2, T_test_sell2, Rank_test_sell2],)
    s2 = s2.reindex(columns=[3, 2, 1, 0])

    results = pd.concat([s1, s2], axis=1)
    results.to_excel(writer, sheet_name=subcriteria, index=False)
writer.save()
