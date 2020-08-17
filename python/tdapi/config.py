# This information is obtained upon registration of app (https://developer.tdameritrade.com/user/me/apps)
# Consumer Key of ExcelRTD app 
import os

# This information is obtained upon registration of app (https://developer.tdameritrade.com/user/me/apps)
# Consumer Key of ExcelRTD app 
CLIENT_ID = "W5LOW5PIXKAIOFJ6VPC62DADFPGJJZ60"
CLIENT_ID_AUTH: str = CLIENT_ID + '@AMER.OAUTHAP'

REDIRECT_URI = "https://localhost:8080/callback"

APPDATA_PATH = os.path.join(os.getenv('APPDATA'), 'TD_ExcelRTD')

TOKEN_URL = 'https://api.tdameritrade.com/v1/oauth2/token'
TOKEN_FILE_NAME = os.path.join(APPDATA_PATH, 'token.json')
USERPRINCIPALS_FILE_NAME = os.path.join(APPDATA_PATH, 'user_principals.json')
CREDENTIALS_FILE_NAME = os.path.join(APPDATA_PATH, 'credentials.json')

AUTHORIZATION_BASE_URL = 'https://auth.tdameritrade.com/auth'
TOKEN_ENDPOINT = 'oauth2/token'
API_ENDPOINT = 'https://api.tdameritrade.com'
API_VERSION = 'v1'
