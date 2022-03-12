import requests
from bs4 import BeautifulSoup
import pandas as pd

excelStocks = 'Data.xlsx'
df = pd.read_excel(excelStocks, sheet_name='Ticker')
dfTicker = pd.DataFrame(df, columns=['Ticker'])

df = pd.read_excel(excelStocks, sheet_name='Fields')
dfFields = pd.DataFrame(df, columns=['Fields'])

listFields=['Ticker'] + dfFields['Fields'].tolist()

df = pd.DataFrame(
     columns=listFields)

print(df)

symbolnr = 0
for index, row in dfTicker.iterrows():
    symbol = row['Ticker']
    print(symbol)

    url = f'https://finance.yahoo.com/quote/{symbol}/key-statistics?p={symbol}'
    headers = {
        'User-agent': 'Mozilla/5.0',
    }

    print(url)

    r = requests.get(url, headers=headers)

    soup = BeautifulSoup(r.text, features="lxml")

    tables = soup.findAll("table")

    df = df.append(pd.Series(), ignore_index=True)
    df.iloc[-1, df.columns.get_loc('Ticker')] = symbol

    for table in tables:
        for tr in table.find_all('tr'):
            row = [td.text for td in tr.find_all('td')]
            # print(f'{row[0]}xxx{row[1]}')
            if row[0].rstrip() in df.columns:
                df.iloc[-1, df.columns.get_loc(row[0].rstrip())] = row[1].rstrip()
                # print(f'{row[0]:30}  {row[1]}')
    symbolnr = ++symbolnr

df['Total Debt/Equity (mrq)'] = df['Total Debt/Equity (mrq)'] + '%'
df.to_excel("FinancialData.xlsx")
