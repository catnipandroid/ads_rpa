# Authorization data
import base64
import requests
import json

"""
kauth.kakao.com/oauth/authorize?client_id=0a41f3b0076a902e89bdeb0612bd8fba&redirect_uri=http://zqkrwogusx12.godomall.com/main/html.php?htmid=main/moment.html&response_type=code
"""

kakao_auth_uri = 'https://kauth.kakao.com/oauth/token'
kakao_restAPI_key = '0a41f3b0076a902e89bdeb0612bd8fba'
authorize_code = '7gDueWDJh_D6bplIgJdwspV4OBSPWA9z2ojWw3IzxDWFHVmSA1X_JQMEKWiavRtsTPWVKgo9dVsAAAF9obDnmA'
redirect_uri = 'http://zqkrwogusx12.godomall.com/main/html.php?htmid=main/moment.html'

data = {
    'grant_type': 'authorization_code',
    'client_id': kakao_restAPI_key,
    'redirect_uri': redirect_uri,
    'code': authorize_code,
}

response = requests.post(kakao_auth_uri, data=data)
token = response.json()

print(token)
# access_token = token.access_token
# refresh_token = token.refresh_token
