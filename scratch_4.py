#02035
from bs4 import BeautifulSoup, Comment
import requests
import pandas as pd
import numpy as np
import itertools
from tabulate import tabulate
import types
import datetime

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


date = "2017-07-12"

url = "finance.yahoo.com/calendar/earnings?day=" + date
r  = requests.get("https://" + url)
data = r.text
soup = BeautifulSoup(data, "lxml" )
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
    #print line

ticker_urls = []
#--------------------------------------Paste Test Section Here----------------------------------------------------------

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

        # geo_con_web_list = []
        # geo_con_web_list.append(geo_con_web.splitlines())
        # profile_info.append(geo_con_web_list)
        #
        # #profile_info.append(c_name.get_text())
        # #profile_info.append(geo_con_web)
        # #profile_info.append(sec_ind_emp)
        # #profile_info.append(descript)
        #
        # df_ = pd.DataFrame(profile_info, columns=['Address',''])
        # print df_
    except:
        "Stock Profile Unavailable"