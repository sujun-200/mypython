#!/usr/bin/python
# -*- coding: UTF-8 -*-
import string
import time
import requests
import random
import json
import hashlib
import threading


class TokenGenerator(object):
    '''
         缓存token，获取token之后缓存，然后每次使用有效期内的token，设置定机器清空过期token
    '''
    token = ""

    def __init__(self, _productId, _publicKey, _secretKey):
        self.productId = _productId
        self.publicKey = _publicKey
        self.secretKey = _secretKey
        self._value_lock = threading.Lock()

    def clearToken(self):
        print("clear token")
        TokenGenerator.token = ""

    def getToken(self):
        # global token
        if TokenGenerator.token:
            return TokenGenerator.token
        with self._value_lock:
            if TokenGenerator.token:
                return TokenGenerator.token
            print("generate token......")
            url = "https://api.talkinggenie.com/api/v2/public/authToken"
            timeStamp = str(int(round(time.time() * 1000)))
            sign = hashlib.md5(
                (self.publicKey + self.productId + timeStamp + self.secretKey).encode('utf8')).hexdigest()
            # 请求头
            data = {
                "productId": self.productId,
                "publicKey": self.publicKey,
                "sign": sign,
                "timeStamp": timeStamp
            }
            headers = {
                'Content-Type': "application/json"
            }
            print(data)
            response = requests.post(url=url, headers=headers, data=json.dumps(data))
            print(response.json())
            TokenGenerator.token = response.json()["result"].get("token")
            print(TokenGenerator.token)
            expireTime = response.json()["result"].get("expireTime")
            print(expireTime)
            timer = threading.Timer(int(int(expireTime) / 1000) - int(round(time.time())) - 60, self.clearToken)
            timer.start()
            return TokenGenerator.token


if __name__ == "__main__":
    tokenGenerator = TokenGenerator("914013393", "05457dfe443546eba4c63badff456d05", "9D48A3670193DCD63DA4C125F4FD11FA")
    print(tokenGenerator.getToken())
    print("end")