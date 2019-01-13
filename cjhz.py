# -*- coding: utf-8 -*-
"""
Created on Sat Jan  5 19:25:17 2019

@author: hanbo
"""

import re
from WindPy import *
import os
import numpy as np
import pandas as pd
from openpyxl import load_workbook


def pick_id(num):
    pat = re.compile(r'\d{5,}(\.SH)?(\.SZ)?(\.IB)?')
    if re.search(pat, num):
        return re.search(pat, num).group()


def pick_rate(num):
    pat = re.compile(r'\d+\.\d+')
    res = pat.findall(num)
    if res:
        return res[-1]


def getcode(id_list, suffix):
    name = w.wss(id_list, "sec_name1", tradeDate=tradeday)
    name_bool = np.array(list(bool(val) for val in name.Data[0]))
    return np.where(
        name_bool,
        np.array(name.Codes),
        np.char.add(id_distinct, [suffix])
    ).tolist()


def make_id_complete(id):
    id_ib = getcode(id, '.IB')  # 试一试加后缀".IB"
    id_sh = getcode(id_ib, '.SH')  # 试一试加加后缀".SH"
    id_sz = getcode(id_sh, '.SZ')  # 试一试加加后缀".SZ"
    return getcode(id_sz, '')


def add_sheet(data_frame, excel_path):
    excel_writer = pd.ExcelWriter(excel_path, engine='openpyxl')
    excel_writer.book = load_workbook(excel_writer.path)
    data_frame.to_excel(excel_writer=excel_writer, sheet_name=tradeday_cn, index=True)
    excel_writer.close()


def main():
    id, rate = [], []
    global id_distinct, tradeday, tradeday_cn, addr_txt

    # 获取文件地址
    for root, dirs, files in os.walk(os.getcwd()):
        f = list(filter(lambda x: x[-3:] == 'txt', files))[-1]
        addr_txt = os.path.join(root, f)
        tradeday_cn = f.split('周')[0]
        break  # break的意思是不再循环子文件夹里面的文件

    file_obj = open(addr_txt)
    try:
        file_context = file_obj.read().splitlines()  # 按行分割
    finally:
        file_obj.close()

    # 迭代每行，pick出该行的id和rate
    for item in file_context:
        p_id = pick_id(item)
        p_rate = pick_rate(item)
        if bool(p_id) * bool(p_rate):
            id.append(p_id)
            rate.append(p_rate)

    w.start()
    tradeday = datetime.strptime(tradeday_cn, '%Y年%m月%d日')
    last_tradeday = w.tdaysoffset(-1, tradeday, "").Data[0][0]

    id_digit = list(filter(lambda x: x.isdigit(), id))  # 只把没有后缀的提取出来
    id_distinct = list(set(id_digit))  # 去重
    id_pd1 = pd.DataFrame({'key': id})
    id_pd2 = pd.DataFrame({'key': id_distinct, 'id_complete': make_id_complete(id_distinct)})

    id_w = pd.merge(id_pd1, id_pd2, on=['key'], how='left')
    id_whole = np.where(
        id_w['id_complete'].isnull(),
        id_w['key'],
        id_w['id_complete']
    )

    # id_suffix_distinct 去重
    id_whole_distinct = list(set(id_whole))

    # 获取横截面数据
    wind_data_1 = w.wss(id_whole_distinct,
                        "couponrate2,sec_name,latestissurercreditrating,ptmyear,calc_mduration,eobspecialinstrutions,nature1,windl1type,ipo_date",
                        tradeDate=tradeday
                        )
    # 获取估值数据，用w.wsd时间序列函数
    wind_data_2 = w.wsd(id_whole_distinct,
                        'yield_cnbd',
                        last_tradeday,
                        tradeday,
                        'credibility=1')

    w.stop
    res_1 = pd.DataFrame(
        index=wind_data_1.Fields,
        columns=wind_data_1.Codes,
        data=wind_data_1.Data
    ).T
    res_2 = pd.DataFrame(
        index=wind_data_2.Codes,
        columns=['yield_cnbd_lastday', 'yield_cnbd'],
        data=wind_data_2.Data
    )

    res_3 = pd.merge(res_1, res_2, left_index=True, right_index=True, how='left')
    id_whole_pd = pd.DataFrame({'id': list(id_whole), 'rate': rate})
    res_4 = pd.merge(id_whole_pd, res_3, left_on=['id'], right_index=True, how='left')

    add_sheet(res_4, os.getcwd() + '/成交汇总.xlsx')


if __name__ == '__main__':
    main()
