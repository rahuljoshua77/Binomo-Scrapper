import os
import requests
import warnings 
warnings.filterwarnings("ignore")
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import openpyxl
from time import sleep

time_to_fetch_again = 5  # in seconds
excel_file_name = "data.xlsx"
#coin_name = "NZD%2FJPY"  # EURO, AUD%2FJPY, EUR%2FNZD-HSFX, NZD%2FJPY, Z-CRY%2FIDX
timeframe = 1  # 1-> 5s, 2->30s, 3->1m

print("[*] Automation Scrapping BINOMO")
print("[*] Choose Coin")
print("[*] 1. EURO USD")
print("[*] 2. CRY IDX")
coin_choose = int(input("[*] Please input your choice (1/2): "))
date_your = input("[*] Input Date (example 2022-04-27): ")
timeframe = int(input("[*] Input timeframe (1=5s, 2=30s, 3=1m): "))
time_api = "https://worldtimeapi.org/api/timezone/Etc/UTC"
file_date = "https://worldtimeapi.org/api/timezone/Asia/Jakarta"
coin_name =  ["EURO","Z-CRY%2FIDX"]
file_coin = ["EURO - USD", "CRYPTO IDX"]
time_frame = [5,30,60][timeframe-1]
file_coin = file_coin[coin_choose-1]
coin_name =  coin_name[coin_choose-1]
def getCurrentTime():
    response = requests.get(time_api)
    current_time = response.json()["datetime"]
    if timeframe == 1:
        current_time = current_time.split(".")[0].split(":")
        current_time = current_time[0] + ":00:00"
    elif timeframe == 2:
        current_time = datetime.strptime(current_time.split('.')[0], "%Y-%m-%dT%H:%M:%S")
        if current_time.hour >= 12:
            current_time = current_time.replace(hour=12, minute=0, second=0)
        else:
            current_time = current_time.replace(hour=0, minute=0, second=0)
        # print(current_time)
        current_time = current_time.strftime("%Y-%m-%dT%H:%M:%S")
    elif timeframe == 3:
        current_time = current_time.split("T")[0] + "T00:00:00"

   
    # print(current_time)
    return current_time

def fileDate():
    response = requests.get(file_date)
    current_time = response.json()["datetime"]
    if timeframe == 1:
        current_time = current_time.split(".")[0].split(":")
        current_time = current_time[0] + ":00:00"
    elif timeframe == 2:
        current_time = datetime.strptime(current_time.split('.')[0], "%Y-%m-%dT%H:%M:%S")
        if current_time.hour >= 12:
            current_time = current_time.replace(hour=12, minute=0, second=0)
        else:
            current_time = current_time.replace(hour=0, minute=0, second=0)
        # print(current_time)
        current_time = current_time.strftime("%Y-%m-%dT%H:%M:%S")
    elif timeframe == 3:
        current_time = current_time.split("T")[0] + "T00:00:00"

   
    # print(current_time)
    current_time = current_time.split("T")
    get_date = current_time[0]
    get_hour = current_time[1]
    return current_time

def toIndonesiaTime(utc_time):
    time = datetime.strptime(utc_time.split('.')[0], "%Y-%m-%dT%H:%M:%S") + timedelta(hours=7)
    # print(time)
    # print(time.strftime("%Y-%m-%dT%H:%M:%S"))
    return time.strftime("%Y-%m-%d %H:%M:%S") 
def getCoinData(i):
    if len(str(i)) == 1:
        i = "0" + str(i)
        
    header = {
        'user-timezone': 'Asia/Jakarta'
    }
    url = f"https://api.binomo-investment.com/candles/v1/{coin_name}/{date_your}T{i}:00:00/{[5, 30, 60][timeframe - 1]}?locale=en"
    print(url)
    response = requests.get(url,headers=header).json()
    #print(response)
 
    df = pd.DataFrame(response["data"])
    df['open'] = df['open'].astype(str)
    df['close'] = df['close'].astype(str)
    df['colour'] = np.where(df['open'] <= df['close'], "GREEN", "RED")
    df['created_at'] = df['created_at'].apply(toIndonesiaTime)
    
    # print(df)
    df = df[::-1]
    toExcelSheet(df)


def toExcelSheet(df):
    
    excel_sheet_name = file_coin + " " + ["5s", "30s", "1m"][timeframe - 1] 
    try:
        df2 = pd.read_excel(excel_file_name, sheet_name=excel_sheet_name)
        # df = pd.concat([df2, df])
        df = df.append(df2)
        df.drop_duplicates(subset="created_at", keep='first', inplace=True)
    except Exception as e:
        print(e)
  
    df = df[::-1]
    df = df.apply(np.roll, shift = len(df))
    print(df)
    workbook = openpyxl.load_workbook(excel_file_name)
    writer = pd.ExcelWriter(excel_file_name, engine='openpyxl')
    writer.book = workbook
    writer.sheets = dict((ws.title, ws) for ws in workbook.worksheets)
    df.to_excel(writer, excel_sheet_name, index=False)
    writer.save()
    writer.close()


if time_frame == 5:
    for i in range(1,23):
        try:
            
            getCoinData(i)
        except Exception as e:
            print(e)
        print("Ok")
        sleep(time_to_fetch_again)

elif time_frame == 30:
    list_time = ["00","12"]
    for i in list_time:
        try:
            getCoinData(i)
        except Exception as e:
            print(e)
        print("Ok")
        sleep(time_to_fetch_again)
        
elif time_frame == 60:
    list_time = ["00"]
    for i in list_time:
        try:
            getCoinData(i)
        except Exception as e:
            print(e)
        print("Ok")
        sleep(time_to_fetch_again)