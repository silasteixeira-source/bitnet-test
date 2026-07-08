import urllib.request, urllib.parse
client_id='177636338434-2pq6n7enn93fmqsif4tp6hvdhb5cpkf4.apps.googleusercontent.com'
redirect_uri='https://bitnet-colaboradores.streamlit.app'
scope='https://www.googleapis.com/auth/gmail.send'
url=f'https://accounts.google.com/o/oauth2/auth?response_type=code&client_id={client_id}&redirect_uri={urllib.parse.quote(redirect_uri)}&scope={urllib.parse.quote(scope)}&access_type=offline&prompt=consent'
print('URL:', url)
req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
try:
    urllib.request.urlopen(req)
except urllib.error.HTTPError as e:
    print('Error Code:', e.code)
    print('Error Reason:', e.reason)
    print(e.read().decode('utf-8')[:1000])
