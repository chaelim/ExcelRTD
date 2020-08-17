import os
from tda.auth import easy_client, client_from_login_flow
from tda.client import Client

def make_webdriver():
    # Import selenium here because it's slow to import
    from selenium import webdriver

    driver = webdriver.Chrome()
    atexit.register(lambda: driver.quit())
    return driver

TOKEN_PATH = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'token.pickle')

if os.path.isfile(TOKEN_PATH):
    c = easy_client(
            api_key='W5LOW5PIXKAIOFJ6VPC62DADFPGJJZ60',
            redirect_uri='https://localhost',
            token_path=TOKEN_PATH)
else:
    with make_webdriver() as driver:
        client_from_login_flow(
                driver,
                api_key='W5LOW5PIXKAIOFJ6VPC62DADFPGJJZ60',
                redirect_uri='https://localhost',
                token_path=TOKEN_PATH)

resp = c.get_price_history('AAPL',
        period_type=Client.PriceHistory.PeriodType.YEAR,
        period=Client.PriceHistory.Period.TWENTY_YEARS,
        frequency_type=Client.PriceHistory.FrequencyType.DAILY,
        frequency=Client.PriceHistory.Frequency.DAILY)
assert resp.ok
history = resp.json()