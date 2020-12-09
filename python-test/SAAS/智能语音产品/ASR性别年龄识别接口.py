import string
import requests
from requests_toolbelt import MultipartEncoder
import random

def speech_synthe():
    url = "https://api.talkinggenie.com/portal/api/v1/ba/asr/emotion"
    # 请求头
    multipart_encoder = MultipartEncoder(
        fields={
            'param': '''{"audio": {"audioType": "wav","channel": 1,"sampleBytes": 2,"sampleRate": 16000}}''',
            'file': ('lastSound.wav', open("C:/Users/wwh05/Documents/录音/test.wav", "rb"), "audio/x-wav")
        },
        boundary=''.join(random.sample(string.ascii_letters + string.digits, 30))
    )
    session = requests.session()
    session.headers = {
        'X-AISPEECH-PRODUCT-ID': "914013393",
        'X-AISPEECH-TOKEN': "1d91b739-6910-4743-b524-6f00ff303164"
    }
    session.headers["Content-Type"] = multipart_encoder.content_type
    response = session.post(url=url, data=multipart_encoder, timeout=20)
    data = response.json()
    print(data)

if __name__ == "__main__":
    speech_synthe()
