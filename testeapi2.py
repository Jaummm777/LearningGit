import urllib.request

headers = {
    'Content-Type': 'application/json',
    'access-token': 'private',
    'secret-access-token': 'private'
}

request = urllib.request.Request('https://api.gestaoclick.com/vendas', headers=headers)

try:
    with urllib.request.urlopen(request) as response:
        response_body = response.read()
        print(response_body.decode('utf-8'))
except urllib.error.HTTPError as e:
    print(f'HTTPError: {e.code} - {e.reason}')
except urllib.error.URLError as e:
    print(f'URLError: {e.reason}')
