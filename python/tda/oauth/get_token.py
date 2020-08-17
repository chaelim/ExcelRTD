import atexit
import os
from requests_oauthlib import OAuth2Session
import pickle

token_ptah = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'token.pickle')
client_id = "W5LOW5PIXKAIOFJ6VPC62DADFPGJJZ60"
client_id_auth: str = client_id + '@AMER.OAUTHAP'
redirect_uri = "https://localhost"
authorization_base_url = 'https://auth.tdameritrade.com/auth'
token_url = 'https://api.tdameritrade.com/v1/oauth2/token'

def make_webdriver():
    # Import selenium here because it's slow to import
    from selenium import webdriver

    driver = webdriver.Chrome()
    atexit.register(lambda: driver.quit())
    return driver

oauth = OAuth2Session(
    client_id=client_id_auth,
    redirect_uri=redirect_uri
)

authorization_url, state = oauth.authorization_url(authorization_base_url)

with make_webdriver() as webdriver:
    webdriver.get(authorization_url)

    if redirect_url.startswith('http://'):
        print(('WARNING: Your redirect URL ({}) will transmit data over HTTP, ' +
                'which is a potentially severe security vulnerability. ' +
                'Please go to your app\'s configuration with TDAmeritrade ' +
                'and update your redirect URL to begin with \'https\' ' +
                'to stop seeing this message.').format(redirect_url))

        redirect_urls = (redirect_url, 'https' + redirect_url[4:])
    else:
        redirect_urls = (redirect_url,)

    # Wait until the current URL starts with the callback URL
    current_url = ''
    num_waits = 0
    while not any(current_url.startswith(r_url) for r_url in redirect_urls):
        current_url = webdriver.current_url

        if num_waits > max_waits:
            raise RedirectTimeoutError('timed out waiting for redirect')
        time.sleep(redirect_wait_time_seconds)
        num_waits += 1

    token = oauth.fetch_token(
        'https://api.tdameritrade.com/v1/oauth2/token',
        authorization_response=current_url,
        access_type='offline',
        client_id=api_key,
        include_client_id=True)

    with open(token_path, 'wb') as f:
        pickle.dump(token, f)
