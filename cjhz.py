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


def pick_ID(num):
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
        np.char.add(ID_quchong, [suffix])
    ).tolist()


def add_sheet(data_frame, excel_path):
    excel_writer = pd.ExcelWriter(excel_path, engine='openpyxl')
    excel_writer.book = load_workbook(excel_writer.path)
    data_frame.to_excel(excel_writer=excel_writer, sheet_name=tradeday_cn, index=True)
    excel_writer.close()


ID, rate = [], []
global tradeday_cn, addr_txt

#获取文件地址
for root, dirs, files in os.walk(os.getcwd()):
    f = list(filter(lambda x: x[-3:] == 'txt', files))[-1]
    addr_txt = os.path.join(root, f)
    tradeday_cn = f.split('周')[0]
    break  # break的意思是不再循环子文件夹里面的文件


file_obj = open(addr_txt)
try:
    file_context = file_obj.read().splitlines()#按行分割
finally:
    file_obj.close()

# 迭代每行，pick出该行的ID和rate
for item in file_context:
    p_ID = pick_ID(item)
    p_rate = pick_rate(item)
    if bool(p_ID)*bool(p_rate):
        ID.append(p_ID)
        rate.append(p_rate)

w.start()
tradeday = datetime.strptime(tradeday_cn, '%Y年%m月%d日')
last_tradeday = w.tdaysoffset(-1, tradeday, "").Data[0][0]
ID_nosuffix = list(filter(lambda x: x.isdigit(), ID))
ID_quchong = list(set(ID_nosuffix))
ID_addIB = getcode(ID_quchong, '.IB')
ID_addSH = getcode(ID_addIB, '.SH')
ID_addSZ = getcode(ID_addSH, '.SZ')
ID_final = getcode(ID_addSZ, '')

id_pd1 = pd.DataFrame({'key': ID})
id_pd2 = pd.DataFrame({'key': ID_quchong, 'id_final': ID_final})

res = pd.merge(id_pd1, id_pd2, on=['key'], how='left')

ID_suffix = np.where(
    res['id_final'].isnull(),
    res['key'],
    res['id_final']
    )

# ID_suffix_quchong 去重
ID_suf = list(set(ID_suffix))

res_1 = w.wss(ID_suf,
              "couponrate2,sec_name,latestissurercreditrating,ptmyear,calc_mduration,eobspecialinstrutions,nature1,windl1type,ipo_date",
              tradeDate=tradeday
              )

# 获取估值数据，用w.wsd时间序列函数
res_1_5 = w.wsd(ID_suf,
                'yield_cnbd',
                last_tradeday,
                tradeday,
                'credibility=1')

w.stop


res_2 = pd.DataFrame(
    index=res_1.Fields,
    columns=res_1.Codes,
    data=res_1.Data
).T

res_2_5 = pd.DataFrame(
    index=res_1_5.Codes,
    columns=['yield_cnbd_lastday', 'yield_cnbd'],
    data=res_1_5.Data
)

res_2_6 = pd.merge(res_2, res_2_5, left_index=True, right_index=True, how='left')
ID_suffix_pd = pd.DataFrame({'id': list(ID_suffix), 'rate': rate})
res_3 = pd.merge(ID_suffix_pd, res_2_6, left_on=['id'], right_index=True, how='left')


add_sheet(res_3, os.getcwd()+'/成交汇总.xlsx')
print(1)

# def __main__():
#
#


