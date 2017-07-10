#02035
from bs4 import BeautifulSoup
import requests
import pandas as pd
import itertools
from tabulate import tabulate
import datetime


def pretty_print(table_title, table_to_print):
    df = pd.read_html(str(table_to_print))
    print table_title, "\n", (tabulate(df[0], headers='keys', tablefmt='psql'))

date = "2017-07-12"

url = "finance.yahoo.com/calendar/earnings?day=" + date
r  = requests.get("https://" + url)
data = r.text
soup = BeautifulSoup(data, "html5lib" )
data_table = soup.find('table', {'class':'data-table'})

headers = []
cells = []
tickers = []

cols = data_table.find_all('td')
ths = data_table.find_all('th')

for th in ths[1::]:
    headers.append(th.get_text())

for col in cols[1::]:
    if col != "-":
        cells.append(col.get_text())
    else:
        cells.append("NULL")

w = ''
spl = [list(y) for x, y in itertools.groupby(cells, lambda z: z == w) if not x]

print "Earnings announcements on " + date, "\n", headers, "\n"
for line in spl:
    tickers.append(line[0])
    print line

# Financials (marketwatch.com)----------------------------------------------------------------------------------------
ticker_urls = []

# INCOME STATEMENT-----------------------------------------------------
for ticker in tickers:
    ticker_urls.append("http://www.marketwatch.com/investing/stock/" + ticker + "/financials")

# Run for each URL
for ticker_url in ticker_urls:
    r = requests.get(ticker_url)
    data = r.text
    soup = BeautifulSoup(data, "html5lib")

    try:
        # Company Name
        c_name = soup.find('h1', id='instrumentname')
        print c_name.get_text()
        print "Income Statement"
        fin_tables = soup.find_all('table', ['class', 'crDataTable'])
        for fin_table in fin_tables:
            pretty_print("", fin_table)
    except:
        "Stock Financials Unavailable"

# BALANCE SHEET ----------------------------------------------------
ticker_urls = []

for ticker in tickers:
    ticker_urls.append("http://www.marketwatch.com/investing/stock/" + ticker + "/financials/balance-sheet")

# Run for each URL
for ticker_url in ticker_urls:
    r = requests.get(ticker_url)
    data = r.text
    soup = BeautifulSoup(data, "html5lib")
    fin_heads = []
    x = 0
    y = 0
    try:
        # Company Name
        c_name = soup.find('h1', id='instrumentname')
        print c_name.get_text()
        print "Balance Sheet"
        fin_headers = soup.find_all('h2')
        for fin_header in fin_headers:
            fin_heads.append(fin_header.get_text())
        # only want the 2nd and 3rd
        fin_heads.pop(0)

        fin_tables = soup.find_all('table', ['class', 'crDataTable'])
        for fin_table in fin_tables:
            if len(fin_tables) == 3:
                if x == 0 or x == 2:
                    pretty_print(fin_heads[y], fin_table)
                    y += 1
                else:
                    pretty_print("", fin_table)
            elif len(fin_tables) == 2:
                pretty_print(fin_heads[x], fin_table)
            else:
                pretty_print("", fin_table)
            x += 1
    except:
        "Stock Financials Unavailable"

# CASH FLOW STATEMENT -----------------------------------------------

ticker_urls = []

for ticker in tickers:
    ticker_urls.append("http://www.marketwatch.com/investing/stock/" + ticker + "/financials/cash-flow")

# Run for each URL
for ticker_url in ticker_urls:
    r = requests.get(ticker_url)
    data = r.text
    soup = BeautifulSoup(data, "html5lib")
    fin_heads = []
    x = 0
    y = 0
    try:
        # Company Name
        c_name = soup.find('h1', id='instrumentname')
        print c_name.get_text()
        print "Cash Flow Statement"
        fin_headers = soup.find_all('h2')
        for fin_header in fin_headers:
            fin_heads.append(fin_header.get_text())
        # only want the 2nd and 3rd
        fin_heads.pop(0)

        fin_tables = soup.find_all('table', ['class', 'crDataTable'])
        for fin_table in fin_tables:
            pretty_print(fin_heads[x], fin_table)
            x += 1

    except:
        "Stock Financials Unavailable"
