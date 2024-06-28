import urllib.request

headers = {
    'Content-Type': 'application/json',
    'access-token': 'bf07732c8f55a4a4a4891dca2513f9b9e4514136',
    'secret-access-token': 'c6aa5784b15c110c7dc85a8835f32b98870b10f1'
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