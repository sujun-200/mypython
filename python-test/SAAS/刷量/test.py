from urllib import request, parse

import time

url = "https://page.om.qq.com/page/Okj3Ij1AYt5YS8DN5iJYKMRA0"

handler = request.ProxyHandler({"HTTP": "223.241.78.43:8010"});
opener = request.build_opener(handler)

resp = opener.open(url)
time.sleep(100)
print(resp.read())