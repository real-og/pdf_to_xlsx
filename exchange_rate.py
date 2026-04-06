import requests

def get_exchange_rate():
    r = requests.get('https://api.nbrb.by/exrates/rates/456')
    return float(r.json()['Cur_OfficialRate'])

