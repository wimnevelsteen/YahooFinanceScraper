import requests
from bs4 import BeautifulSoup
import pandas as pd


listSymbols = ['eli.br',
'rio'
'solb.br',
'rand.as',
'eurn.br',
]

df = pd.DataFrame(
  columns=['Tickr', 'Trailing P/E', 'Forward P/E', 'Trailing Annual Dividend Yield 3', 'Enterprise Value/EBITDA', 'Return on Equity (ttm)', 'Current Ratio (mrq)', 'Enterprise Value/EBITDA', 'Total Debt/Equity (mrq)'])

symbolnr=0
for symbol in listSymbols:
  print(symbol)

  url = f'https://finance.yahoo.com/quote/{symbol}/key-statistics?p={symbol}'
  headers = {
      'User-agent': 'Mozilla/5.0',
  }

  print(url)

  r= requests.get(url, headers=headers)

  data=r.text
  soup=BeautifulSoup(data, features="lxml")

  tables = soup.findAll("table")

  df = df.append(pd.Series(), ignore_index=True)
  df.iloc[-1, df.columns.get_loc('Tickr')] = symbol

  for table in tables:
    for tr in table.find_all('tr'):
      row = [td.text for td in tr.find_all('td')]
      #print(f'{row[0]}xxx{row[1]}')
      if row[0].rstrip() in df.columns:
        df.iloc[-1, df.columns.get_loc(row[0].rstrip())] = row[1].rstrip()
        # print(f'{row[0]:30}  {row[1]}')
  symbolnr=++symbolnr

#print(df.to_markdown)
df.to_excel("output.xlsx")