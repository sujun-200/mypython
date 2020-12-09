import requests
import time

# pyinstaller -c -F 刷量/csdn.py

url = [
    'https://blog.csdn.net/four_three/article/details/107017137',
    'https://blog.csdn.net/four_three/article/details/110668344',
    'https://blog.csdn.net/four_three/article/details/110670191',
    'https://blog.csdn.net/four_three/article/details/107014962',
    'https://blog.csdn.net/four_three/article/details/110876058',
    'https://blog.csdn.net/four_three/article/details/110918127'
    ]

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.111 Safari/537.36'}

count = 0
countUrl = len(url)

# 访问次数设置
for i in range(1, 1000):

    if count < 5000:
        try:  # 正常运行
            for i in range(countUrl):
                response = requests.get(url[i], headers=headers)
                if response.status_code == 200:
                    count = count + 1
                    print('Success ' + str(count), 'times')
            time.sleep(70)

        except Exception:  # 异常
            print('Failed and Retry')
            time.sleep(60)




