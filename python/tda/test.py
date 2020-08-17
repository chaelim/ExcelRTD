import json
import requests

API_KEY="W5LOW5PIXKAIOFJ6VPC62DADFPGJJZ60"

endpoint = 'https://api.tdameritrade.com/v1/marketdata/{stock_ticker}/quotes?'

full_url = endpoint.format(stock_ticker='MSFT')
page = requests.get(url=full_url,
                    params={'apikey' : API_KEY})
content = json.loads(page.content)

print(content)