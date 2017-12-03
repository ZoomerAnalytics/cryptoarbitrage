import xlwings as xw
import requests

cache = dict()

def hello_xlwings():
    wb = xw.Book.caller()
    wb.sheets[0].range("A1").value = "Hello xlwings!"


@xw.func
def requests_get(url, path=None, reload=False):
    if path and '#' in url:
        raise ValueError('You cannot use fully qualified path in url and explicit path')

    if '#' in url:
        url, path = url.split('#')

    if reload or url not in cache:
        response = requests.get(url)
        j = response.json()
        cache[url] = j
    
    j = cache[url]
    if path:
        for p in path.split('/'):
            j = j[int(p)] if isinstance(j, list) else j[p] 

    if isinstance(j, dict):
        return url if not path else f'{url}#{path}'
    
    if isinstance(j, list):
        for item in j:
            if isinstance(item, list):
                for col in item:
                    if isinstance(col, (list, dict)):
                        return url if not path else f'{url}#{path}'

    return j
