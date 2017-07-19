from bs4 import BeautifulSoup
import requests
import pandas as pd
import numpy as np
import requests.packages.urllib3
requests.packages.urllib3.disable_warnings()

table = ""

ticker = "ACU"
ticker_url = "http://ycharts.com/companies/" + ticker
r = requests.get(ticker_url)
data = r.text
soup = BeautifulSoup(data, "html5lib")

print soup.prettify()

#print soup

try:
    comp_table = soup.find('table', ['class','relCompSect'])
    table = pd.read_html(str(comp_table))
except:
    print "no comps found"

print table