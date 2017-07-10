from bs4 import BeautifulSoup
import requests
from IPython.display import Image

url = "finance.yahoo.com/calendar/earnings?day=2017-07-03"
r  = requests.get("https://" + url)

data = r.text
soup = BeautifulSoup(data, "lxml" )

data_table = soup.find('table', {'class':'data-table'})

headerDict = {}
earnDict = {}
results = []

cols = data_table.find_all('td')
ths = data_table.find_all('th')

y = 0
for th in ths[1::]:
    print th.get_text()
    headerDict.update({th.get_text() : y})
    y = y + 1

print headerDict


#get tickers
for col in cols[1::7]:
    earnDict.update({col.get_text() : headerDict})
    print col.get_text()
    #print earnDict[col.get_text()][headerDict[x]]

# #x = 1
# #z = 0
# #i = 0
# #for col in cols:
# #    z = z + 1
# #    if  z % 49 == 0:
# #        i = i + 1
# #        x = 49 * i
# #    print cols[(z % 7 * 7 + 1) + x].get_text(), earnDict[cols[(z % 7 * 7 + 1) + x].get_text()]#.update({})
#
#
# #x = 0
# #z = 0
# #i = 1
# #j = 0
#
# #for col in cols:
#     #print cols[z % 7 * 7 + 1 + x].get_text(), earnDict[cols[z % 7 * 7 + 1 + x].get_text()]
# #    z = z + 1
# #    if z % 7 == 0:
# #        x = 49 * i
# #        i = i + 1
#
# x = 0
# z = 0
# i = 1
j = 1
#
for col in cols:
     if j < 154:
#         print j
#         print z % 7 * 7 + 1 + x
#         print cols[j].get_text()
#         #earnDict[cols[z % 7 * 7 + 1 + x].get_text()]['Symbol'] = cols[j].get_text()
         results.append(cols[j].get_text())
         j = j + 1
#         #earnDict[cols[z % 7 * 7 + 1 + x].get_text()]['Company'] = cols[j].get_text()
         results.append(cols[j].get_text())
         j = j + 1
#         #earnDict[cols[z % 7 * 7 + 1 + x].get_text()]['Earnings Call Time'] = cols[j].get_text()
         results.append(cols[j].get_text())
         j = j + 1
#         #earnDict[cols[z % 7 * 7 + 1 + x].get_text()]['EPS Estimate'] = cols[j].get_text()
         results.append(cols[j].get_text())
         j = j + 1
#         #earnDict[cols[z % 7 * 7 + 1 + x].get_text()]['Reported EPS'] = cols[j].get_text()
         results.append(cols[j].get_text())
         j = j + 1
#         #earnDict[cols[z % 7 * 7 + 1 + x].get_text()]['Surprise(%)'] = cols[j].get_text()
         results.append(cols[j].get_text())
         j = j + 2
#     else:
#         break
#     z = z + 1
#     if z % 7 == 0:
#         x = 49 * i
#         i = i + 1
#     print i, j, x, z

#added after-------
y = 0
for th in ths[1::]:
    print th.get_text()
    headerDict.update({th.get_text() : results[y]})
    y = y + 1

for col in cols[1::7]:
    earnDict.update({col.get_text() : headerDict})
    print col.get_text()
#------------------

print earnDict

print len(earnDict)

print results
#nested dictionary for table values
