# 1.連結至TWSE, Get 抽籤資訊
# 2.Get現貨價,算出溢價差
# 3.利用Line 通知

import pandas as pd
# 連至TWSE 爬蟲
url='https://www.twse.com.tw/announcement/publicForm?response=html&yy=2022'
list_str = pd.read_html(url)  #list
df=list_str[0]

df2=df.copy()
row_len,col_len=df2.shape

import datetime
Date_Condition=datetime.date.today()  #2022-10-19 <class 'datetime.date'>
#Date_Condition = datetime.date(2022, 10, 10)

list_data=[]
for m in range(row_len): #Row index 每次跑一筆Data
    date_arr=df2.iloc[m, 6].split('/')
    input_date = datetime.date(int(date_arr[0])+1911, int(date_arr[1]), int(date_arr[2]))
    #print(df2.iloc[m, 6])
    #if df2.iloc[m,6]             #申購結束日
    if Date_Condition >= input_date :
        break

df2=df2.drop(list(range(m,row_len)))

list_columns=[]

#df2的Col index
#df.columns
#for n in range(df.columns.size):  #Column index 每次跑一個
for n in range(col_len):  # Column index 每次跑一個
    list_columns.append(df2.columns[n][1])
    #print(df2.columns[n][1])

df2.columns=list_columns
df2=df2.set_index('序號')

# df2.to_excel(r'D:\temp\out.xlsx')    # debug 用

# 處理Line 問題
import requests
import twstock
import time


def lineNotify(token, msg):
    headers = {
        "Authorization": "Bearer " + token,  # 注意"Bearer ", 有一個空白
        "Content-Type": "application/x-www-form-urlencoded"
    }

    payload = {'message': msg}
    notify = requests.post("https://notify-api.line.me/api/notify", headers=headers, params=payload)
    return notify.status_code


row_len, col_len = df2.shape
token = 'VZSWE7IPJ1NTiTZ9lAcHUUvHBvHLC3yKbHb3JpMNG5E'  # 權杖

for m in range(row_len):  # Row index 每次跑一筆Data
    #    print(df2.iloc[m, 1]+','+df2.iloc[m, 2]+','+df2.iloc[m, 4]+','+df2.iloc[m, 5])
    #    print(df2.columns[1]+','+df2.columns[2]+','+df2.columns[4]+','+df2.columns[5])

    # df2.iloc[m, 1] #證券名稱
    # df2.iloc[m, 2] #證券代號
    # df2.iloc[m, 4] #申購開始日
    # df2.iloc[m, 5] #申購結束日
    # df2.iloc[m, 7] #實際承銷股數
    # df2.iloc[m, 9] #實際承銷價(元)

    number_stock= int(df2.iloc[m, 7]/1000)  #  股票張數
    message = df2.columns[1] + '=' + df2.iloc[m, 1] + ',' + \
              df2.columns[2] + '=' + df2.iloc[m, 2] + ',' + \
              df2.columns[4] + '=' + df2.iloc[m, 4] + ',' + \
              df2.columns[5] + '=' + df2.iloc[m, 5] + ',' + \
              "承銷張數" + '=' + str(number_stock) + ',' + \
              df2.columns[9] + '=' + str(df2.iloc[m, 9])

    latest_trade_price = ''
    price_gap = ''

    while True :
        time.sleep(3)
        stock = twstock.realtime.get(df2.iloc[m, 2])  # 查詢上市股票之即時資料
        if stock['success']:  # 上市股票
            latest_trade_price = stock['realtime']['latest_trade_price']
            if latest_trade_price != '-' :
                # print(stock)
                # print('latest_trade_price=' + latest_trade_price)
                price_gap = str(eval(latest_trade_price + '-' + str(df2.iloc[m, 9])))  # 價差
                underwriting_price = float(df2.iloc[m, 9])  # 承銷價
                Premium_ratio = (float(price_gap) / underwriting_price) * 100  # 溢價比
                message = message + ' ,最新價格=' + "{:.2f}".format(float(latest_trade_price))  + ' ,價差=' + "{:.2f}".format(float(price_gap))  + ' ,溢價差(%)=' + "{:.2f}".format(Premium_ratio)
                break
        else :
            break

    # print(message+'\n')

    code = lineNotify(token, message)
    if code == 200:
        print('發送 Line Notify 成功')
    else:
        print('發送 Line Notify 失敗')