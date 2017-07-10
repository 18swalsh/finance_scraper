from bs4 import BeautifulSoup

import requests

#url = raw_input("Enter a website to extract the URL's from: ")

url = "finance.yahoo.com/calendar/earnings?day=2017-07-03"

r  = requests.get("https://" +url)

data = r.text

soup = BeautifulSoup(data, "lxml" )

type(soup)

for link in soup.find_all('a'):
    print(link.get('href'))