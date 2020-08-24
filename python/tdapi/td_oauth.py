from requests_oauthlib import OAuth2Session
from flask import Flask, request, redirect, session, url_for
from flask.json import jsonify
import logging
import datetime
import json
import os
import pickle
import requests
import time
import win32crypt

from typing import Dict
from typing import List

if __package__:
    from .td_config import APPDATA_PATH, CLIENT_ID, CLIENT_ID_AUTH, REDIRECT_URI, AUTHORIZATION_BASE_URL, TOKEN_URL, TOKEN_FILE_NAME
    from .td_config import USERPRINCIPALS_FILE_NAME, CREDENTIALS_FILE_NAME, API_ENDPOINT, API_VERSION, TOKEN_ENDPOINT
else:
    from td_config import APPDATA_PATH, CLIENT_ID, CLIENT_ID_AUTH, REDIRECT_URI, AUTHORIZATION_BASE_URL, TOKEN_URL, TOKEN_FILE_NAME
    from td_config import USERPRINCIPALS_FILE_NAME, CREDENTIALS_FILE_NAME, API_ENDPOINT, API_VERSION, TOKEN_ENDPOINT

logger = logging.getLogger(__name__)
logger.setLevel(logging.ERROR)
logger.addHandler(logging.StreamHandler())

app = Flask(__name__)

if not os.path.exists(APPDATA_PATH):
    os.makedirs(APPDATA_PATH)

@app.route("/")
def demo():
    """Step 1: User Authorization.

    Redirect the user/resource owner to the OAuth provider (i.e. Github)
    using an URL with a few key OAuth parameters.
    """
    td_session = OAuth2Session(
        client_id=CLIENT_ID_AUTH,
        redirect_uri=REDIRECT_URI
    )

    authorization_url, state = td_session.authorization_url(AUTHORIZATION_BASE_URL)

    # State is used to prevent CSRF, keep this for later.
    session['oauth_state'] = state
    return redirect(authorization_url)


# Step 2: User authorization, this happens on the provider.

@app.route("/callback", methods=["GET"])
def callback():
    """ Step 3: Retrieving an access token.

    The user has been redirected back from the provider to your registered
    callback URL. With this redirection comes an authorization code included
    in the redirect URL. We will use that to obtain an access token.
    """
    td_session = OAuth2Session(
        client_id=CLIENT_ID_AUTH,
        redirect_uri=REDIRECT_URI,
        state=session['oauth_state']
    )
    token = td_session.fetch_token(
        TOKEN_URL,
        access_type='offline',
        authorization_response=request.url,
        include_client_id=True
    )

    # At this point you can fetch protected resources but lets save
    # the token and show how this is done from a persisted token
    # in /profile.
    session['oauth_token'] = token

    save_token(token)

    # Grab the Streamer Info.
    #userPrincipalsResponse = get_user_principals(
    #    token,
    #    fields=['streamerConnectionInfo', 'streamerSubscriptionKeys', 'preferences', 'surrogateIds'])

    #if userPrincipalsResponse:
    #    save_credentials(userPrincipalsResponse)

    return redirect(url_for('shutdown'))

def shutdown_server():
    func = request.environ.get('werkzeug.server.shutdown')
    if func is None:
        raise RuntimeError('Not running with the Werkzeug Server')
    func()

@app.route('/shutdown', methods=['GET'])
def shutdown():
    shutdown_server()
    return '<html><head>Server shutting down...</head><body>Now you can close this and go back to Excel</body></html>'
    
@app.route("/profile", methods=["GET"])
def profile():
    """Fetching a protected resource using an OAuth 2 token.
    """
    td_session = OAuth2Session(CLIENT_ID, token=session['oauth_token'])
    return jsonify(td_session.get('https://api.td_session.com/user').json())

def get_token():
    pass

def get_user_principals(token, fields: List[str]) -> Dict:
    """Returns User Principal details.

    Documentation:
    ----
    https://developer.tdameritrade.com/user-principal/apis/get/userprincipals-0

    Arguments:
    ----

    fields: A comma separated String which allows one to specify additional fields to return. None of 
        these fields are returned by default. Possible values in this String can be:

            1. streamerSubscriptionKeys
            2. streamerConnectionInfo
            3. preferences
            4. surrogateIds

    Usage:
    ----
        >>> td_client.get_user_principals(fields=['preferences'])
        >>> td_client.get_user_principals(fields=['preferences','streamerConnectionInfo'])
    """

    # define the endpoint
    endpoint = 'userprincipals'

    # build the params dictionary
    params = {
        'fields': ','.join(fields)
    }

    parts = [self.API_ENDPOINT, self.API_VERSION, endpoint]
    url = '/'.join(parts)

    headers = {
        'Authorization': 'Bearer {token}'.format(token=token['access_token'])
    }

    # Define a new session.
    request_session = requests.Session()
    request_session.verify = True

    # Define a new request.
    request_request = requests.Request(
        method='GET',
        headers=headers,
        url=url,
        params=params,
    ).prepare()
    
    # Send the request.
    response: requests.Response = request_session.send(request=request_request)

    request_session.close()

    # grab the status code
    status_code = response.status_code

    # grab the response headers.
    response_headers = response.headers

    if response.ok:
        return response.json()
    else:
        return None

def load_token():
    try:
        with open(TOKEN_FILE_NAME, 'rb') as encoded_file:
            encoded_data = encoded_file.read()
            token_data = json.loads(win32crypt.CryptUnprotectData(encoded_data)[1].decode())

        return token_data
    except Exception as e:
        return None

def save_token(token_dict: dict) -> bool:
    # make sure there is an access token before proceeding.
    if 'access_token' not in token_dict:
        return False

    token_data = {}

    # save the access token and refresh token
    token_data['access_token'] = token_dict['access_token']
    token_data['refresh_token'] = token_dict['refresh_token']

    # store token expiration time
    access_token_expire = time.time() + int(token_dict['expires_in'])
    refresh_token_expire = time.time() + int(token_dict['refresh_token_expires_in'])
    token_data['access_token_expires_at'] = access_token_expire
    token_data['refresh_token_expires_at'] = refresh_token_expire
    token_data['access_token_expires_at_date'] = datetime.datetime.fromtimestamp(access_token_expire).isoformat()
    token_data['refresh_token_expires_at_date'] = datetime.datetime.fromtimestamp(refresh_token_expire).isoformat()
    token_data['logged_in'] = True

    token_json = json.dumps(token_data)
    try:
        with open(TOKEN_FILE_NAME, 'wb') as encoded_file:
            enc = win32crypt.CryptProtectData(token_json.encode())
            encoded_file.write(enc)
    except Exception as e:
        return False

    return True

def save_credentials(userPrincipalsResponse):
    # Grab the timestampe.
    tokenTimeStamp = userPrincipalsResponse['streamerInfo']['tokenTimestamp']

    # Grab socket
    socket_url = userPrincipalsResponse['streamerInfo']['streamerSocketUrl']

    # Parse the token timestamp.
    token_timestamp = datetime.datetime.strptime(tokenTimeStamp, "%Y-%m-%dT%H:%M:%S%z")
    tokenTimeStampAsMs = int(token_timestamp.timestamp()) * 1000


    # Define our Credentials Dictionary used for authentication.
    credentials = {
        "userid": userPrincipalsResponse['accounts'][0]['accountId'],
        "token": userPrincipalsResponse['streamerInfo']['token'],
        "company": userPrincipalsResponse['accounts'][0]['company'],
        "segment": userPrincipalsResponse['accounts'][0]['segment'],
        "cddomain": userPrincipalsResponse['accounts'][0]['accountCdDomainId'],
        "usergroup": userPrincipalsResponse['streamerInfo']['userGroup'],
        "accesslevel": userPrincipalsResponse['streamerInfo']['accessLevel'],
        "authorized": "Y",
        "timestamp": tokenTimeStampAsMs,
        "appid": userPrincipalsResponse['streamerInfo']['appId'],
        "acl": userPrincipalsResponse['streamerInfo']['acl']
    }

    with open(file=USERPRINCIPALS_FILE_NAME, mode='w+') as json_file:
        json.dump(obj=userPrincipalsResponse, fp=json_file, indent=4)

    with open(file=CREDENTIALS_FILE_NAME, mode='w+') as json_file:
        json.dump(obj=credentials, fp=json_file, indent=4)

def _token_seconds(token_data, token_type: str = 'access_token') -> int:
    """Determines time till expiration for a token.
    
    Return the number of seconds until the current access token or refresh token
    will expire. The default value is access token because this is the most commonly used
    token during requests.

    Arguments:
    ----
    token_type {str} --  The type of token you would like to determine lifespan for. 
        Possible values are ['access_token', 'refresh_token'] (default: {access_token})
    
    Returns:
    ----
    {int} -- The number of seconds till expiration.
    """

    # if needed check the access token.
    if token_type == 'access_token':

        # if the time to expiration is less than or equal to 0, return 0.
        if not token_data['access_token'] or time.time() + 60 >= token_data['access_token_expires_at']:
            return 0

        # else return the number of seconds until expiration.
        token_exp = int(token_data['access_token_expires_at'] - time.time() - 60)

    # if needed check the refresh token.
    elif token_type == 'refresh_token':

        # if the time to expiration is less than or equal to 0, return 0.
        if not token_data['refresh_token'] or time.time() + 60 >= token_data['refresh_token_expires_at']:
            return 0

        # else return the number of seconds until expiration.
        token_exp = int(token_data['refresh_token_expires_at'] - time.time() - 60)

    return token_exp

def grab_refresh_token(access_token, refresh_token) -> bool:
    """Refreshes the current access token.
    
    This takes a  valid refresh token and refreshes
    an expired access token.

    Returns:
    ----
    {bool} -- `True` if successful, `False` otherwise.
    """

    # build the parameters of our request
    data = {
        'client_id': CLIENT_ID_AUTH,
        'grant_type': 'refresh_token',
        'access_type': 'offline',
        'refresh_token': refresh_token
    }

    # build url: https://api.tdameritrade.com/v1/oauth2/token
    parts = [API_ENDPOINT, API_VERSION, TOKEN_ENDPOINT]
    url = '/'.join(parts)

    # Define a new session.
    request_session = requests.Session()
    request_session.verify = True

    headers = { 'Content-Type': 'application/x-www-form-urlencoded' }

    # Define a new request.
    request_request = requests.Request(
        method='POST',
        headers=headers,
        url=url,
        data=data
    ).prepare()

    # Send the request.
    response: requests.Response = request_session.send(request=request_request)

    request_session.close()

    if response.ok:
        save_token(response.json())
        return True

    return False

def silent_sso() -> bool:
    try:
        token_data = load_token()

        # if the current access token is not expired then we are still authenticated.
        if _token_seconds(token_data, token_type='access_token') > 0:
            return True

        # if the refresh token is expired then you have to do a full login.
        elif _token_seconds(token_data, token_type='refresh_token') <= 0:
            return False

        # if the current access token is expired then try and refresh access token.
        elif token_data['refresh_token'] and grab_refresh_token(token_data['access_token'], token_data['refresh_token']):
            return True
    except Exception as e:
        print(repr(e))
        return False

    return True

def _run_full_oauth() -> None:
    import webbrowser
    webbrowser.open_new_tab('https://localhost:8080/')

    app.secret_key = os.urandom(24)
    app.run(ssl_context='adhoc', host="localhost", port=8080, debug=False)

def run_full_oauth_subprocess() -> None:
    from subprocess import run
    run(["python", os.path.realpath(__file__)], cwd= os.path.dirname(os.path.realpath(__file__)))

if __name__ == "__main__":
    import sys

    # Check if current token is valid
    if silent_sso():
        sys.exit(0)
    else:
        _run_full_oauth()