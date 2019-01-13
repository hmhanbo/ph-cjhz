from WindPy import *
import pandas as pd

w.start()
def wind_to_dataframe(wfunc, *args, **kargs):
    # print(wfunc, args, kargs)
    if wfunc == 'wss' or wfunc == 'w.wss':
        r = w.wss(*args, **kargs)
        return pd.DataFrame(
            index=r.Fields,
            columns=r.Codes,
            data=r.Data
        ).T
    elif wfunc == 'wsd' or wfunc == 'w.wsd':
        r = w.wsd(*args)
        return pd.DataFrame(
            index=r.Codes,
            data=r.Data
        )

tradeday = '20181225'
last_tradeday = '20181224'
ID = ['180210.IB', '180205.IB']

r = wind_to_dataframe('wss',
                      ID,
                      "couponrate2,sec_name,latestissurercreditrating,ptmyear,calc_mduration,eobspecialinstrutions,nature1,windl1type,ipo_date",
                      tradeDate=tradeday)

r_1 = wind_to_dataframe('wsd',
                            ID,
                            'yield_cnbd',
                            last_tradeday,
                            tradeday,
                            'credibility=1')

print(1)