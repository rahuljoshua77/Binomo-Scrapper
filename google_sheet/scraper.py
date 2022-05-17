import os
import requests
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import openpyxl
from time import sleep
import gspread


sa = gspread.service_account(filename="creds.json")


time_to_fetch_again = 5  # in seconds
excel_file_name = "data.xlsx"
# coin_name = "NZD%2FJPY"  # EURO, AUD%2FJPY, EUR%2FNZD-HSFX, NZD%2FJPY, Z-CRY%2FIDX
timeframe = 1  # 1-> 5s, 2->30s, 3->1m

print("[*] Automation Scrapping BINOMO")
print("[*] Choose Coin")
print("[*] 1. EURO USD")
print("[*] 2. CRY IDX")
coin_choose = int(input("[*] Please input your choice (1/2): "))
timeframe = int(input("[*] Input timeframe (1=5s, 2=30s, 3=1m): "))
time_api = "https://worldtimeapi.org/api/timezone/Etc/UTC"
file_date = "https://worldtimeapi.org/api/timezone/Asia/Jakarta"
coin_name = ["EURO", "Z-CRY%2FIDX"]
file_coin = ["EURO - USD", "CRYPTO IDX"]
file_coin = file_coin[coin_choose - 1]
coin_name = coin_name[coin_choose - 1]


def getCurrentTime():
    response = requests.get(time_api)
    current_time = response.json()["datetime"]
    if timeframe == 1:
        current_time = current_time.split(".")[0].split(":")
        current_time = current_time[0] + ":00:00"
    elif timeframe == 2:
        current_time = datetime.strptime(
            current_time.split(".")[0], "%Y-%m-%dT%H:%M:%S"
        )
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
        current_time = datetime.strptime(
            current_time.split(".")[0], "%Y-%m-%dT%H:%M:%S"
        )
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


def toIndonesiaTime(utc_time):
    time = datetime.strptime(utc_time.split(".")[0], "%Y-%m-%dT%H:%M:%S") + timedelta(
        hours=7
    )
    # print(time)
    # print(time.strftime("%Y-%m-%dT%H:%M:%S"))
    return time.strftime("%Y-%m-%d %H:%M:%S")


def To_SpreadSheet(df):
    sheet_name = file_coin + " " + ["5s", "30s", "1m"][timeframe - 1]
    sh = sa.open("new_data")
    try:
        wks = sh.worksheet(sheet_name)
    except:
        print("error")
        wks2 = sh.add_worksheet(title=sheet_name, rows="1", cols="6")
        wks2.update("A1", "open")
        wks2.update("B1", "high")
        wks2.update("C1", "low")
        wks2.update("D1", "close")
        wks2.update("E1", "created_at")
        wks2.update("F1", "colour")
        wks = sh.worksheet(sheet_name)
    already = [item for item in wks.col_values(5) if item]
    new = []
    for row in df.iterrows():
        if (row[1]["created_at"]) not in already:
            new.append(
                [
                    row[1]["open"],
                    row[1]["high"],
                    row[1]["low"],
                    row[1]["close"],
                    row[1]["created_at"],
                    row[1]["colour"],
                ]
            )
    wks.append_rows(new)


def getCoinData():
    header = {"user-timezone": "Asia/Jakarta"}
    url = f"https://api.binomo-investment.com/candles/v1/{coin_name}/{getCurrentTime()}/{[5, 30, 60][timeframe - 1]}?locale=en"
    print(url)
    response = requests.get(url, headers=header).json()
    # print(response)

    df = pd.DataFrame(response["data"])
    df["open"] = df["open"].astype(str)
    df["close"] = df["close"].astype(str)
    df["colour"] = np.where(df["open"] <= df["close"], "GREEN", "RED")
    df["created_at"] = df["created_at"].apply(toIndonesiaTime)

    df = df[::-1]
    To_SpreadSheet(df)


while True:
    try:
        getCoinData()
    except Exception as e:
        print(e)
    print("Ok")
