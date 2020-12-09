#!/usr/bin/python
# -*- coding:utf-8 -*-

import requests
from urllib.parse import urlencode

import json
import hashlib
import time



header = {
    "authToken":"",
    "userId": "jianye.yin",
    "BaSource": "kf",
    "modav": "true"
}

payload = {'pageNo': '0', 'pageSize': '10','query': '{"prodId":"914010298"}'}

r = requests.get(url="https://www.tgenie.cn/api/v2/smart/sadmin/customer/api/v1/sceneMess",headers=header, verify=False,params=payload)
print(r.text)
#cookie = {
#    "authToken":"eyJhbGciOiJIUzUxMiJ9.eyJzdWIiOiJqaWFueWUueWluI-WwueW7uuS4miMxNjA2Mjg4OTM5MDgxIn0.YFZi3xVWmzm_E10iN-1M8nZ8ByLeB0wLs5DGohuIAA_ybarwZuXi3j62C5zKsZITjnBknPNhFTn82QHRRkcT4w"
#}
r = requests.get(url="https://www.tgenie.cn/api/v2/smart/sadmin/customer/api/v1/sceneMess?pageNo=1&pageSize=10",headers=header, verify=False)

print(r.text)  # 查看打印结果headers中的Cookie和User-Agent的values

r = requests.get(url="https://www.tgenie.cn/api/v2/smart/sadmin/voice/api/v1/tts?pageNo=1&pageSize=10&query=%7B%22ttsType%22%3A%220%22%7D",headers=header, verify=False)

print(r.text)  # 查看打印结果headers中的Cookie和User-Agent的values