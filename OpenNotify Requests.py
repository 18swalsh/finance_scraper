import requests

p = {"lat": 40.71, "lon": -74}

response = requests.get("http://api.open-notify.org/iss-pass.json", params = p)

print(response.status_code)
print(response.content.decode("utf-8"))

data = response.json()
print(data)
print (response.headers)
print (data)

response2 = requests.get("http://api.open-notify.org/astros.json")
data2 = response2.json()

print(data2["number"])
print(data2)