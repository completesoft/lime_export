import requests


url = 'http://hive.product.in.ua/api/test/'
headers = {'content-type': 'application/json'}

status = "ok"

json = {"id": 123, "status": status}

r = requests.post(url, json=json, headers=headers)

print(r.text)
