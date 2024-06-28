"""
@Author : evelyn
Created on June/22 23:12:19 2024
"""

import yfinance as yf
import pandas as pd
import xlwings as xw
import numpy as np
import psycopg2
import requests
import time
import random
from datetime import date,timedelta
from bs4 import BeautifulSoup

# Yahoo Finance價量資料
def get_pv_datas(symbol):
    pv_datas=yf.download(f'{symbol}',start='2023-01-01',end='2023-12-31')
    return pv_datas

# Money DJ配息資料
def get_dividend(symbol):
    header={'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36'}
    url=f'https://www.moneydj.com/ETF/X/Basic/Basic0005.xdjhtm?etfid={symbol}.TW'
    r=requests.get(url,headers=header)
    r.encoding='UTF-8'
    soup=BeautifulSoup(r.text,'html.parser')
    datas=soup.select('table.datalist')[0]
    rows=datas.select('tr')
    list_rows=list()
    for row in rows:
        row_td=[i.text for i in row.select('td')]
        if len(row_td)>1:
             list_rows.append(np.array(row_td)[[1,2,6]])
    df = pd.DataFrame(list_rows, columns =['ex_div_date','pay_date','div_amount'] )
    df.set_index(df.columns[0],inplace=True)
    return df

# google trend資料(excel)
def trend():
    df=pd.read_excel('D:\python\CCclub\multiTimeline.xlsx')
    df.set_index(df.columns[0],inplace=True)
    return df
    

# goodinfo區間漲跌幅資料(excel)
def goodinfo():
    df=pd.read_excel('D:\python\CCclub\CompareDetail.xlsx')
    df.set_index(df.columns[0],inplace=True)
    return df

# 淨值(excel)
def nw(symbol):
    df=pd.read_excel(fr'D:\python\CCclub\nw\{symbol}.xlsx')
    df.set_index(df.columns[0],inplace=True)
    return df

# 淨值postgreSQL抓檔(供參)
def fund(symbol):
    host=input('host:')
    dbname="fund"
    user=input('username:')
    password=input('password:')
    sslmode="allow"

    conn=psycopg2.connect("host={0} user={1} dbname={2} password={3} sslmode={4}".format(host,user,dbname,password,sslmode))
    print("Connection established")

    conn_cursor = conn.cursor()
    conn_cursor.execute(f'''SELECT zdate,val
                    FROM fund.fdnav_txdsa
                    WHERE comp_id = {symbol}
                    ''')
    rows=conn_cursor.fetchall()

    Data_row=[]
    for row in rows:
        Data_row += [[row[0],row[1]]]

    df=pd.DataFrame(Data_row[0:])
    return df

# TWSE三大法人資料(sleep時間設較長，大約會抓20~30min)
def get_all_inv(start_year, start_month, start_day, end_year, end_month, end_day, symbol):
    start_date = str(date(start_year, start_month, start_day))
    end_date = str(date(end_year, end_month, end_day))
    date_list = pd.date_range(start_date, end_date, freq='D').strftime("%Y%m%d").tolist()
    result_df = pd.DataFrame()
    header={'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36'}
    for day in date_list:
        url = 'https://www.twse.com.tw/rwd/zh/fund/T86?date=' + day + '&selectType=0099P&response=html'
        r=requests.get(url,headers=header)
        soup=BeautifulSoup(r.text,'html.parser')
        datas=soup.select('div')
        word=''
        for data in datas:
            if "很抱歉，沒有符合條件的資料!" in data.get_text(strip=True):
                word=data.get_text(strip=True)
        if word!='':
            continue
        else:
            print(day)
            try:
                df=pd.read_html(url.format(symbol))
                df=df[0]
                df.insert(0, '日期', day)
                df.reset_index(drop=True,inplace=True)
                df.columns= ['日期', '證券代號', '證券名稱', '外陸資買進股數(不含外資自營商)', '外陸資賣出股數(不含外資自營商)', '外陸資買賣超股數(不含外資自營商)', '外資自營商買進股數', '外資自營商賣出股數', '外資自營商買賣超股數', '投信買進股數', '投信賣出股數', '投信買賣超股數', '自營商買賣超股數', '自營商買進股數(自行買賣)', '自營商賣出股數(自行買賣)', '自營商買賣超股數(自行買賣)', '自營商買進股數(避險)', '自營商賣出股數(避險)', '自營商買賣超股數(避險)', '三大法人買賣超股數']
                result_df=pd.concat([result_df,df])
            except:
                pass
            time.sleep(10*random.random())
    if symbol == None:
        pass
    else:
        result_df = result_df[result_df['證券代號'] == str(symbol)]
    result_df.set_index(df.columns[0],drop=True,inplace=True)
    return result_df


# 匯入工作表
def input_wb(wb,result_df,name,index):
    sheet = wb.sheets[index]
    sheet.name = f'{name}'
    sheet.range("A1").value = result_df

# 將所有資料整理至同一個excel檔
wb = xw.Book()
#for i in range(0, 5):
for i in range(0, 3):
    wb.sheets.add(f'Sheet{i}')
wb.save(r'D:\python\CCclub\00923.xlsx')

#lst=[get_pv_datas('00896.TW') , get_dividend('00896') , goodinfo() , trend() , nw('00896') , get_all_inv(2023,1,1,2023,12,31,'00896')]
lst=[get_pv_datas('00923.TW') , nw('00923') , get_all_inv(2023,1,1,2023,12,31,'00923')]
#name_lst=['價量資料','配息資料','漲跌幅','Google_trend','淨值','三大法人']
name_lst=['價量資料','淨值','三大法人']

for index,df in enumerate(lst):
    input_wb(wb,df,name_lst[index],index)

#直接在這裡整理