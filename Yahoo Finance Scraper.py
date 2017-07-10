from bs4 import BeautifulSoup
import requests
import itertools

#Get earnings announcements for a given date

url = "finance.yahoo.com/calendar/earnings?day=2017-07-03"
r  = requests.get("https://" + url)
data = r.text
soup = BeautifulSoup(data, "lxml" )
data_table = soup.find('table', {'class':'data-table'})

headers = []
cells = []
earnDict = {}

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

print headers, "\n"
for line in spl:
    print line


