# x = 0
# z = 0
# i = 1
#
# for j in range(0,300):
#     z = z + 1
#     if z % 7 == 0:
#         x = 49 * i
#         i = i + 1
#     print z % 7 * 7 + 1 + x
#
#
# #for col in cols:
# #    z = z + 1
# #    if  z % 49 == 0:
# #        i = i + 1
# #        x = 49 * i
# #    print cols[(z % 7 * 7 + 1) + x].get_text(), earnDict[cols[(z % 7 * 7 + 1) + x].get_text()]#.update({})
#
# print 22*7
# import datetime
#
#
#
#
#
# d_valid = False
# # Get earnings announcements for a given date
# while d_valid == False:
#     date = str(raw_input("Enter a date or type 'exit' (Format: YYYY-MM-DD): "))
#     try:
#         datetime.datetime.strptime(date, '%Y-%m-%d')
#         print "Nailed it"
#         d_valid = True
#     except:
#         if date =='exit':
#             print "bye"
#             d_valid = True
#         else:
#             print "Nah" + "\n"

# import urllib
#
# urllib.urlretrieve("https://finance.yahoo.com/quote/FB/history?/p=FB")

#02035
from bs4 import BeautifulSoup
import requests
import pandas as pd
import itertools
import datetime

d_valid = False
#---------------------------- Get earnings announcements for a given date ----------------------------------------------
#WORKS - JUST ANNOYING TO INPUT THE DATE FOR DEBUGGING PURPOSES

# while d_valid == False:
#     date = str(raw_input("Enter a date or type 'exit' (Format: YYYY-MM-DD): "))
#     try:
#         datetime.datetime.strptime(date, '%Y-%m-%d')
#         d_valid = True
#     except:
#         if date =='exit':
#             print "bye"
#             exit()
#         else:
#             print "Nah" + "\n"

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
    print line

ticker_urls = []




#--------------------------------------Paste Test Section Here----------------------------------------------------------






# Stock Quote Page
for ticker in tickers:
    ticker_urls.append("https://finance.yahoo.com/quote/" + ticker + "/analysts?p=" + ticker)

# Run for each URL
for ticker_url in ticker_urls:
    r = requests.get(ticker_url)
    data = r.text
    soup = BeautifulSoup(data, "lxml")

    try:
        # Company Name
        c_name = soup.find('h1', ['class', 'D(ib)'])
        print c_name.get_text()

        analyst_est_sect = soup.find('section',id='quote-leaf-comp')
        analyst_tables = analyst_est_sect.find_all('table', ['class','W(100%)'])
        for table in analyst_tables:
            est_heads = table.find_all('th')
            print "Column Names"
            for est_head in est_heads:
                print est_head.get_text()
            ests = table.find_all('td')
            print "\n"
            for est in ests:
                print est.get_text()
            print "\n"

        #all_tables = soup.find('table',class_='W(100%)')
        #for ind_table in all_tables:

        #grades_table = all_tables

        # grades_div = soup.find('div',id='Col2-5-QuoteModule-Proxy')
        # step_down = grades_div.findChildren()
        # grades_tables = step_down[0].find_all('table', class_='W(100%)')
        # for grades_table in grades_tables:
        #     grades = grades_table.find_all('td')
        #     for grade in grades:
        #         print grade.get_text()


        output = soup.findall('td', class_= 'W(70px) Pend(12px) Fw(500) Bxz(bb)')
        for td in output:
            print td.get_text()
        break








    except:
        "Stock Profile Unavailable"