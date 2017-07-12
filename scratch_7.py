
#02035
from bs4 import BeautifulSoup, Comment
import requests
import pandas as pd
import numpy as np
import itertools
from tabulate import tabulate
import types
import datetime
import time
import xlsxwriter

all_data = []

d_valid = False
#---------------------------- Get earnings announcements for a given date ----------------------------------------------
date = ""
while d_valid == False:
    date = str(raw_input("Enter a date or type 'exit' (Format: YYYY-MM-DD): "))
    try:
        datetime.datetime.strptime(date, '%Y-%m-%d')
        d_valid = True
    except:
        if date =='exit':
            print "bye"
            exit()
        else:
            print "Nah" + "\n"

start_time = time.time()
time_one = time.time() #----------------------------------------------------------------------------------------------------------------------

def pretty_print(table_title, table_to_print):
    df = pd.read_html(str(table_to_print))
    print table_title, "\n", (tabulate(df[0], headers='keys', tablefmt='psql'))
    all_data.append(pd.DataFrame(df))

def replace_with_newlines(element):
    text = ''
    for elem in element.recursiveChildGenerator():
        if isinstance(elem, types.StringTypes):
            text += elem.strip()
        elif elem.name == 'br':
            text += '\n'
    return text

#date = "2017-09-19"

url = "finance.yahoo.com/calendar/earnings?day=" + date
r  = requests.get("https://" + url)
data = r.text
soup = BeautifulSoup(data, "html5lib" )
data_table = soup.find('table', {'class':'data-table'})

headers = []
cells = []
tickers = []

pretty_print('Earnings announcements on ' + date, data_table)

cols = data_table.find_all('td')
#ths = data_table.find_all('th')
#for th in ths[1::]:
#    headers.append(th.get_text())
for col in cols[1::]:
    if col != "-":
        cells.append(col.get_text())
    else:
        cells.append("NULL")

w = ''
spl = [list(y) for x, y in itertools.groupby(cells, lambda z: z == w) if not x]

for line in spl:
    tickers.append(line[0])
print len(tickers)
time_two = time.time() #----------------------------------------------------------------------------------------------------------------------

#-----------------------add more info for known tickers-----------------------------------------------------------------
#https://finance.yahoo.com/quote/TICKER?p=TICKER - stock page

print "-----------------------------SUMMARY-----------------------------"

ticker_urls = []

#Stock Quote Page
for ticker in tickers:
    ticker_urls.append("https://finance.yahoo.com/quote/" + ticker + "?p=" + ticker)

x = -1
#Run for each URL
for ticker_url in ticker_urls:
    x += 1
    r  = requests.get(ticker_url)
    data = r.text
    soup = BeautifulSoup(data, "html5lib")
    summary_table = soup.find_all('table',['class','W(100%)'])



    #Summary------------------------------------------------------------------------------------------------------------

    #Error if the url is bad
    try:

        cur_price = soup.find('span',['class','Trsdu(0.3s) Fw(b) Fz(36px) Mb(-4px) D(ib)'])
        # Filter out stocks under $20
        if float(cur_price.get_text()) < 20:
            print tickers[x], "was removed from search (price below $20)"
            tickers.remove(tickers[x])
            x -= 1
            continue
        # Company Name
        c_name = soup.find('h1', ['class', 'D(ib)'])
        print "Company: " + c_name.get_text()
        print "Current Price: " + cur_price.get_text()
        pretty_print("Summary", summary_table)
        #General Data
        # nums = summary_table[0].findChildren()
        # nums_two = summary_table[1].findChildren()
        #
        # for child in nums:
        #     print "Company: " + c_name.get_text()
        #     print "Current Price: " + cur_price.get_text()
        #     print nums[0].get_text()
        #     print nums_two[0].get_text()
        #     break



    except:
        print "Stock Quote Unavailable"
        print tickers[x], "removed from search"
        tickers.remove(tickers[x])
        x -= 1

writer = pd.ExcelWriter('Scraper_Data.xlsx', engine='xlsxwriter')
for frame in all_data:
    x += 1
    frame.to_excel(writer,sheet_name='Sheet' + str(x))

writer.save()

