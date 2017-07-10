from bs4 import BeautifulSoup
import re
import requests

#url = raw_input("Enter a website to extract the URL's from: ")
url = "www.goodreads.com/search?q=the+big+short"
r  = requests.get("https://" +url)

data = r.text

soup = BeautifulSoup(data, "lxml" )

for rating in soup.find_all('span', {"class" : "minirating"}):
    rated = str(re.findall("\d\.\d\d",rating.text))[3:7]
    print(rated)




