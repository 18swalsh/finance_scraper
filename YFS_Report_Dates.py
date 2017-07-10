#02035
from bs4 import BeautifulSoup
import requests
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




#-----------------------add more info for known tickers-----------------------------------------------------------------
#https://finance.yahoo.com/quote/TICKER?p=TICKER - stock page
print tickers
ticker_urls = []

#Stock Quote Page
for ticker in tickers:
    ticker_urls.append("https://finance.yahoo.com/quote/" + ticker + "?p=" + ticker)

#Run for each URL
for ticker_url in ticker_urls:
    r  = requests.get(ticker_url)
    data = r.text
    soup = BeautifulSoup(data, "lxml")
    summary_table = soup.find_all('table',['class','W(100%)'])


    #Summary------------------------------------------------------------------------------------------------------------

    #Error if the url is bad
    try:
        #Company Name
        c_name = soup.find('h1', ['class', 'D(ib)'])
        #Current Price
        cur_price = soup.find('span',['class','Trsdu(0.3s) Fw(b) Fz(36px) Mb(-4px) D(ib)'])

        #General Data
        nums = summary_table[0].findChildren()
        nums_two = summary_table[1].findChildren()

        for child in nums:
            print "Company: " + c_name.get_text()
            print "Current Price: " + cur_price.get_text()
            print nums[0].get_text()
            print nums_two[0].get_text()
            break

    except:
        print "Stock Quote Unavailable"

print "!!!!!!!!!!!!!!!!         STATISTICS          !!!!!!!!!!!!!!!!!!!!!"
#Statistics---------------------------------------------------------------------------------------------------------
ticker_urls = []

# Stock Quote Page
for ticker in tickers:
    ticker_urls.append("https://finance.yahoo.com/quote/" + ticker + "/key-statistics?p=" + ticker)

# Run for each URL
for ticker_url in ticker_urls:
    r = requests.get(ticker_url)
    data = r.text
    soup = BeautifulSoup(data, "lxml")

    try:
        # Company Name
        c_name = soup.find('h1', ['class', 'D(ib)'])
        stat_tables = soup.find_all('table', ['class', 'table-qsp-stats Mt(10px)'])
        print c_name.get_text()
        for stat_table in stat_tables:
            print stat_table.get_text()

    except:
        "Stock Stats Unavailable"

print "!!!!!!!!!!!!!!!!         PROFILE          !!!!!!!!!!!!!!!!!!!!!"
#Profile------------------------------------------------------------------------------------------------------------
ticker_urls = []

# Stock Quote Page
for ticker in tickers:
    ticker_urls.append("https://finance.yahoo.com/quote/" + ticker + "/profile?p=" + ticker)

# Run for each URL
for ticker_url in ticker_urls:
    r = requests.get(ticker_url)
    data = r.text
    soup = BeautifulSoup(data, "lxml")

    try:
        # Company Name
        c_name = soup.find('h1', ['class', 'D(ib)'])
        geo_con_web = soup.find('p', ['class', 'D(ib) W(47.727%) Pend(40px)'])
        sec_ind_emp = soup.find('p', ['class', 'D(ib) Va(t)'])
        execs_table = soup.find_all('table', ['class', 'W(100%)'])
        execs = execs_table[0].findChildren()
        descript = soup.find('p',['class','Mt(15px) Lh(1.6)'])

        print c_name.get_text()
        for executive in execs:
            print executive.get_text()
        print geo_con_web.get_text()
        print sec_ind_emp.get_text()
        print descript.get_text()

    except:
        "Stock Profile Unavailable"

print "!!!!!!!!!!!!!!!!         FINANCIALS (broken)         !!!!!!!!!!!!!!!!!!!!!"
#Financials seem to be broken---------------------------------------------------------------------------------------



#Historical Data ----- (potentially download) ----------------------------------------------------------------------
print  "!!!!!!!!!!!!!!           HISTORICAL DATA (download file)       !!!!!!!!!!!!!!!!!!"



print "!!!!!!!!!!!!!!!!              ANALYSTS                      !!!!!!!!!!!!!!!!!!!!!"
#Analysts------------------------------------------------------------------------------------------------------------

ticker_urls = []

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

        analyst_est_sect = soup.find('section', id='quote-leaf-comp')
        analyst_tables = analyst_est_sect.find_all('table', ['class', 'W(100%)'])
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

        #bs can't find the Upgrades and Downgrades section for some reason
        # grade_section = soup.find('section',['data-test','upgrade-downgrade-history'])
        # grade_table = soup.find('table', ['class', "W(100%)"])
        # cols = grade_table.find_all('td')
        # for c in cols:
        #     print c.get_text()



    except:
        "Stock Profile Unavailable"



