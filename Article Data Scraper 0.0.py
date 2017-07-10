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

def pretty_print(table_title, table_to_print):
    df = pd.read_html(str(table_to_print))
    print table_title, "\n", (tabulate(df[0], headers='keys', tablefmt='psql'))

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

#-----------------------add more info for known tickers-----------------------------------------------------------------
#https://finance.yahoo.com/quote/TICKER?p=TICKER - stock page

print "-----------------------------SUMMARY-----------------------------"

ticker_urls = []

#Stock Quote Page
for ticker in tickers:
    ticker_urls.append("https://finance.yahoo.com/quote/" + ticker + "?p=" + ticker)

#Run for each URL
for ticker_url in ticker_urls:
    r  = requests.get(ticker_url)
    data = r.text
    soup = BeautifulSoup(data, "html5lib")
    summary_table = soup.find_all('table',['class','W(100%)'])



    #Summary------------------------------------------------------------------------------------------------------------

    #Error if the url is bad
    try:
        #Company Name
        c_name = soup.find('h1', ['class', 'D(ib)'])
        cur_price = soup.find('span',['class','Trsdu(0.3s) Fw(b) Fz(36px) Mb(-4px) D(ib)'])
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

print "-----------------------------STATISTICS-----------------------------"
#Statistics---------------------------------------------------------------------------------------------------------
ticker_urls = []
stat_heads = []
# Stock Quote Page
for ticker in tickers:
    ticker_urls.append("https://finance.yahoo.com/quote/" + ticker + "/key-statistics?p=" + ticker)

# Run for each URL
for ticker_url in ticker_urls:
    r = requests.get(ticker_url)
    data = r.text
    soup = BeautifulSoup(data, "html5lib")
    x = 0
    try:
        # Company Name
        c_name = soup.find('h1', ['class', 'D(ib)'])
        stat_tables = soup.find_all('table', ['class', 'table-qsp-stats Mt(10px)'])

        print c_name.get_text()
        stat_headers = soup.find_all(True, {'class':['Pt(20px)','Pt(6px) Pstart(20px)','Fz(s) Mt(20px)']})
        for stat_header in stat_headers:
            stat_heads.append(stat_header.get_text())


        for stat_table in stat_tables:
            if stat_heads[x] == "Financial Highlights" or stat_heads[x] == "Trading Information":
                pretty_print(stat_heads[x] + "\n" + "\n" + stat_heads[x+1],stat_table)
                x += 1
            else:
                pretty_print(stat_heads[x], stat_table)
            x += 1
    except:
        "Stock Stats Unavailable"

print "-----------------------------PROFILE-----------------------------"
#Profile------------------------------------------------------------------------------------------------------------
ticker_urls = []

# Stock Quote Page
for ticker in tickers:
    ticker_urls.append("https://finance.yahoo.com/quote/" + ticker + "/profile?p=" + ticker)

profile_info = []

# Run for each URL
for ticker_url in ticker_urls:
    r = requests.get(ticker_url)
    data = r.text
    soup = BeautifulSoup(data, "html5lib")

    try:
        # strip comments
        comments = soup.find_all(text=lambda text: isinstance(text, Comment))
        [comment.extract() for comment in comments]

        # Company Name
        c_name = soup.find('h1', ['class', 'D(ib)'])
        geo_con_web = replace_with_newlines(soup.find('p', ['class', 'D(ib) W(47.727%) Pend(40px)']))
        sec_ind_emp = replace_with_newlines(soup.find('p', ['class', 'D(ib) Va(t)']))
        execs_table = soup.find_all('table', ['class', 'W(100%)'])
        #execs = execs_table[0].findChildren()
        descript = replace_with_newlines(soup.find('p',['class','Mt(15px) Lh(1.6)']))

        print c_name.get_text()
        pretty_print("Executives", execs_table)
        print geo_con_web.splitlines()
        print sec_ind_emp.splitlines()
        print "Description: " + descript

    except:
        "Stock Profile Unavailable"

print "-----------------------------FINANCIALS-----------------------------"
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



#Historical Data ----- (potentially download) ----------------------------------------------------------------------
print  "-----------------------------HISTORICAL DATA (download file)-----------------------------"
print "To be completed"


print "-----------------------------ANALYSTS-----------------------------"
#Analysts------------------------------------------------------------------------------------------------------------

ticker_urls = []

# Stock Quote Page
for ticker in tickers:
    ticker_urls.append("https://finance.yahoo.com/quote/" + ticker + "/analysts?p=" + ticker)

# Run for each URL
for ticker_url in ticker_urls:
    r = requests.get(ticker_url)
    data = r.text
    soup = BeautifulSoup(data, "html5lib")

    try:
        # Company Name
        c_name = soup.find('h1', ['class', 'D(ib)'])
        print c_name.get_text()

        analyst_est_sect = soup.find('section', id='quote-leaf-comp')
        analyst_tables = analyst_est_sect.find_all('table', ['class', 'W(100%)'])
        for a_table in analyst_tables:
            pretty_print("",a_table)

            # here is an alternative if the  4 lines above break
            # all_tables = soup.find_all('table', class_='W(100%)')
            # for table in all_tables:
            #     pretty_print("This Table", table

        #bs can't find the Upgrades and Downgrades section for some reason
        # grade_section = soup.find('section',['data-test','upgrade-downgrade-history'])
        # grade_table = soup.find('table', ['class', "W(100%)"])
        # cols = grade_table.find_all('td')
        # for c in cols:
        #     print c.get_text()

    except:
        "Stock Profile Unavailable"

print "Program took", (time.time() - start_time)/60 , "minutes to run"

