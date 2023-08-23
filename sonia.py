import datetime
from dateutil import parser
import openpyxl as xl
import pandas as pd
import requests

cur_date = '22.08.23'
cur_month = 'AUG'
cur_year = 2023

cur_date_datetime = parser.parse(cur_date).date()

key = "bf5204b93cd4e38625e4d899fc6d5e9f"

# shares = ['AAPL', 'AMZN', 'META', 'MSFT', 'NFLX', 'NVDA', 'QQQ', 'TSLA']

shares = ['AAPL']

for share in shares:
    url = rf'https://financialmodelingprep.com/api/v3/historical-chart/1min/{share}?apikey={key}'
    path = rf'E:\sonia daily data\1 min cash\{cur_year}\{cur_month}\{cur_date}\{share} 1 min csh.xlsx'

    response = requests.get(url)

    data = response.json()

    df = pd.DataFrame(data)
    df["date"] = pd.to_datetime(df["date"])
    df["Date"] = df["date"].dt.date
    df["Time"] = df["date"].dt.time

    df.drop(df[df.Date < cur_date_datetime].index, inplace=True)

    df = df.iloc[:, [7, 1, 3, 2, 4, 5]]
    df = df.sort_values(by='Time')

    df.to_excel(path, index=False)
    # print(df)

