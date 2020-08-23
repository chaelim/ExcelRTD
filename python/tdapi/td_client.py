import logging
import asyncio
import threading
import json
import requests
import time
import urllib

from datetime import datetime
from typing import List
from typing import Dict
from typing import Set
from typing import Union
from collections.abc import Iterable

import websockets

from .fields import CSV_FIELD_KEYS
from .fields import CSV_FIELD_KEYS_LEVEL_2
from .fields import STREAM_FIELD_IDS
from .fields import LEVEL_ONE_QUOTE_KEY_LIST
from .fields import LEVEL_ONE_QUOTE_VALUE_LIST

from .config import API_ENDPOINT
from .config import API_VERSION
from .config import TOKEN_FILE_NAME

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
logger.addHandler(logging.StreamHandler())

class TDWSConn(object):
    def __init__(self, websocket_url):
        """Constructor"""
        super(TDWSConn, self).__init__()
        self.ws = None

    async def send(self, message) -> None:
        if self.ws:
            await self.ws.send(message)

    async def recv_message(self) -> str:
        message = await self.ws.recv()
        return message

    async def start(self, url, initial_request) -> None:
        logger.debug("start")
        try:
            self.ws = await websockets.connect(url)
            await self.send(initial_request)
            await asyncio.sleep(0.5)
        except Exception as e:
            logger.error("Failed to connect the server: {}".format(repr(e)))

    async def close(self) -> None:
        if self.ws:
            await self.ws.close()
            self.ws = None

    @property
    def closed(self):
        if self.ws:
            return self.ws.closed
        else:
            return True

class TDQuote():
    def __init__(self, ticker: str, fields: dict, msg_timestamp: int) -> None:
        self._ticker = ticker
        self._fields = fields
        self._timestamp = msg_timestamp

    @property
    def time_recieved(self, as_datetime: bool = False) -> Union[datetime, int]:
        if as_datetime:
            return datetime.fromtimestamp(t=self._timestamp)
        else:
            return self._timestamp

    @property
    def ticker(self) -> str:
        return self._ticker

    @property
    def fields(self) -> Dict:
        return self._fields

def _create_quotes_from_content(msg_content: List[dict], msg_timestamp: int) -> List[TDQuote]:
    tdquotes = []
    for c in msg_content:
        ticker = None
        fields_dict = {}
        
        for k, v in c.items():
            if k == 'key':
                ticker = v
                continue
            
            if k in LEVEL_ONE_QUOTE_KEY_LIST:
                field_key = LEVEL_ONE_QUOTE_VALUE_LIST[LEVEL_ONE_QUOTE_KEY_LIST.index(k)]
            else:
                field_key = k

            fields_dict[field_key] = v

        tdquotes.append(TDQuote(ticker, fields_dict, msg_timestamp))

    return tdquotes


class TDClient():
    """
        TD Ameritrade Streaming API Client Class.

        Implements a Websocket object that connects to the TD Streaming API, submits requests,
        handles messages, and streams data back to the user.
    """

    def __init__(self, websocket_url: str, user_principal_data: dict, credentials: dict) -> None:     
        """Initalizes the Streaming Client.
        
        Initalizes the Client Object and defines different components that will be needed to
        make a connection with the TD Streaming API.

        Arguments:
        ----
        websocket_url {str} -- The websocket URL that is returned from a Get_User_Prinicpals Request.

        user_principal_data {dict} -- The data that was returned from the "Get_User_Principals" request. 
            Contains the info need for the account info.

        credentials {dict} -- A credentials dictionary that is created from the "create_streaming_session"
            method.
        
        Usage:
        ----

            >>> td_session = TDClient(
                client_id='<CLIENT_ID>',
                redirect_uri='<REDIRECT_URI>',
                credentials_path='<CREDENTIALS_PATH>'
            )
            >>> td_session.login()
            >>> td_stream_session = td_session.create_streaming_session()

        """

        self.websocket_url = "wss://{}/ws".format(websocket_url)
        self.credentials = credentials
        self.user_principal_data = user_principal_data

        self.conn = TDWSConn(self.websocket_url)
        self.async_thread = None
        self.async_loop = None
        self.message_queue = None
        self.sub_queue = None
        self.shutdown = False
        self.connect_event = threading.Event()
        self.on_quote_received = None

        # this will hold all of our requests
        self.data_requests = {"requests": []}
        self.fields_ids_dictionary = STREAM_FIELD_IDS

        self._requestid = 0
        self.subscriptions = {}
        self._outstanding_requests = set()

    def _async_thread_handler(self) -> None:
        logger.debug("_async_thread_handler")
        try:
            asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            self.async_loop = loop
            self.message_queue = asyncio.Queue()
            self.sub_queue = asyncio.Queue()
            login_request = self._build_login_request()
            loop.run_until_complete(self.conn.start(self.websocket_url, login_request))
            self.connect_event.set()

            send_messages_coro = self._send_message_async()
            send_sub_coro = self._send_sub_message_async()
            recv_messages_coro = self._recv_message_async()
            loop.run_until_complete(asyncio.gather(send_messages_coro, send_sub_coro, recv_messages_coro))
            loop.close()
        except Exception as e:
            logger.error("_async_thread_handler: {}".format(repr(e)), exc_info=True)
            self.connect_event.set()

    async def _send_message_async(self) -> None:
        while not self.shutdown:
            message = await self.message_queue.get()
            if message:
                await self.conn.send(message)
            else:
                await self.conn.close()

    async def _send_sub_message_async(self) -> None:
        await asyncio.sleep(1)
        while not self.shutdown:
            while len(self._outstanding_requests) > 0:
                await asyncio.sleep(0.5)
            message = await self.sub_queue.get()
            if not message:
                await self.conn.close()
                return
            
            symbols = set()
            while True:
                if message['type'] == 'subscribe':
                    symbol = message['symbol']
                    field = message['field']

                    if symbol in self.subscriptions:
                        self.subscriptions[symbol].add(field)
                    else:
                        self.subscriptions[symbol] = set(['0', field])

                    symbols.add(symbol)
                    if len(symbols) == 100:
                        break
                    if self.sub_queue.empty():
                        break

                    message = await self.sub_queue.get()

            payload = self._level_one_quotes(symbols)

            if payload:
                await self.conn.send(payload)
    
    async def _recv_message_async(self) -> None:
        # message types: "notify", "response", "data"
        # {"response":[{"service":"ADMIN","requestid":"0","command":"LOGIN","timestamp":1597561577408,"content":{"code":0,"msg":"08-4"}}]}
        while not self.shutdown:
            message = await self.conn.recv_message()
            decoded_message = message.encode(
                'utf-8').replace(
                    b'\xef\xbf\xbd',
                    bytes('"None"', 'utf-8')
                ).decode('utf-8')
            decoded_message = json.loads(decoded_message)

            if 'data' in decoded_message:
                self._handle_data(decoded_message['data'])
            elif 'response' in decoded_message:
                for r in decoded_message['response']:
                    self._outstanding_requests.remove(int(r['requestid']))
                logging.debug(f'MsgRecv: response - {decoded_message}')
            elif 'notify' in decoded_message:
                logging.debug(f'MsgRecv: notify - {decoded_message}')
            else:
                logging.warn(f'MsgRecv: unknown - {decoded_message}')

    def _new_request_template(self) -> dict:
        """Serves as a template to build new service requests.

        This takes the Request template and populates the required fields
        for a subscription request.

        Returns:
        ----
        {dict} -- The service request with the standard fields filled out.
        """

        # first get the current service request count
        #service_count = len(self.data_requests['requests']) + 1
        self._requestid += 1
        self._outstanding_requests.add(self._requestid)

        request = {
            "service": None, 
            "requestid": str(self._requestid),
            "command": None,
            "account": self.user_principal_data['accounts'][0]['accountId'],
            "source": self.user_principal_data['streamerInfo']['appId'],
            "parameters": {
                "keys": None, 
                "fields": None
            }
        }

        return request

    def _validate_argument(self, argument, endpoint: str) -> Union[List[str], str]:
        """Validate field arguments before submitting request.

        Arguments:
        ---
        argument {Union[str, int]} -- Either a single argument or a list of arguments that are
            fields to be requested.
        
        endpoint {str} -- The subscription service the request will be sent to. For example,
            "level_one_quote".

        Returns:
        ----
        Union[List[str], str] -- The field or fields that have been validated.
        """        

        # initalize a new list.
        arg_list = []

        # see if the argument is a list or not.
        if isinstance(argument, Iterable):

            for arg in argument:

                arg_str = str(arg)

                if arg_str in LEVEL_ONE_QUOTE_KEY_LIST:
                    arg_list.append(arg_str)
                elif arg_str in LEVEL_ONE_QUOTE_VALUE_LIST:
                    key_value = LEVEL_ONE_QUOTE_KEY_LIST[LEVEL_ONE_QUOTE_VALUE_LIST.index(arg_str)]
                    arg_list.append(key_value)                  

            return arg_list

        else:

            arg_str = str(argument)
            key_list = list(self.fields_ids_dictionary[endpoint].keys())
            val_list = list(self.fields_ids_dictionary[endpoint].values())

            if arg_str in key_list:
                return arg_str
            elif arg_str in val_list:
                key_value = key_list[val_list.index(arg_str)]
                return key_value

    def _build_login_request(self) -> str:
        """Builds the Login request for the streamer.

        Builds the login request dictionary that will 
        be used as the first service request with the 
        streaming API.

        Returns:
        ----
        [str] -- A JSON string with the login details.

        """        

        self._outstanding_requests.add(0)

        # define a request
        login_request = {
            "requests": [
                {
                    "service": "ADMIN",
                    "requestid": "0",
                    "command": "LOGIN",
                    "account": self.user_principal_data['accounts'][0]['accountId'],
                    "source": self.user_principal_data['streamerInfo']['appId'],
                    "parameters": {
                        "credential": urllib.parse.urlencode(self.credentials),
                        "token": self.user_principal_data['streamerInfo']['token'],
                        "version": "1.0",
                        "qoslevel": "0" # 0 = Express (500 ms)
                    }
                },
            ]
        }

        return json.dumps(login_request, separators=(',', ':'))

    def _handle_data(self, data_components: List) -> None:
        try:
            for component in data_components:
                if component['service'] == 'QUOTE' and component['command'] == 'SUBS':
                    quotes = _create_quotes_from_content(component['content'], component['timestamp'])
                    self.on_quote_received(quotes)
        except Exception as e:
            logging.error(repr(e))

    def connect(self, on_quote_received):
        logging.debug('connect')
        self.on_quote_received = on_quote_received
        self.async_thread = threading.Thread(target=self._async_thread_handler)
        self.async_thread.start()
        self.connect_event.wait()

    def quality_of_service(self, qos_level: str) -> None:
        """Quality of Service Subscription.
        
        Allows the user to set the speed at which they recieve messages
        from the TD Server.

        Arguments:
        ----
        qos_level {str} -- The Quality of Service level that you wish to set. 
            Ranges from 0 to 5 where 0 is the fastest and 5 is the slowest.

        Raises:
        ----
        ValueError: Error if no field is passed through.

        Usage:
        ----
            >>> td_session = TDClient(
                client_id='<CLIENT_ID>',
                redirect_uri='<REDIRECT_URI>',
                credentials_path='<CREDENTIALS_PATH>'
            )

            >>> td_session.login()
            >>> td_stream_session = td_session.create_streaming_session()
            >>> td_stream_session.quality_of_service(qos_level='express')
            >>> td_stream_session.stream()
        """
        # valdiate argument.
        qos_level = self._validate_argument(argument=qos_level, endpoint='qos_request')

        if qos_level is not None:

            # Build the request
            request = self._new_request_template()
            request['service'] = 'ADMIN'
            request['command'] = 'QOS'
            request['parameters']['qoslevel'] = qos_level
            self.async_loop.call_soon_threadsafe(lambda: self.message_queue.put_nowait(json.dumps({"requests":[request]}, separators=(',', ':'))))

        else:
            raise ValueError('No Quality of Service Level provided.')

    def _level_one_quote(self, symbol: str, fields: Union[List[str], List[int]]) -> str:
        """
            Represents the LEVEL ONE QUOTES endpoint for the TD Streaming API. This
            will return quotes for a given list of symbols along with specified field information.

            NAME: symbols
            DESC: A List of symbols you wish to stream quotes for.
            TYPE: List<String>

            NAME: fields
            DESC: The fields you want returned from the Endpoint, can either be the numeric representation
                  or the key value representation. For more info on fields, refer to the documentation.
            TYPE: List<Integer> | List<Strings>
        """

        # valdiate argument.
        fields = self._validate_argument(
            argument=fields, endpoint='level_one_quote')

        # Build the request
        request = self._new_request_template()
        request['service'] = 'QUOTE'
        request['command'] = 'SUBS'
        request['parameters']['keys'] = symbol
        request['parameters']['fields'] = ','.join(fields)

        return json.dumps({"requests":[request]}, separators=(',', ':'))

    def _level_one_quotes(self, symbols: Set) -> None:
        #quote_requests = []
        all_fields = set()
        for symbol in symbols:
            all_fields.update(self.subscriptions[symbol])

        fields = self._validate_argument(
            argument=all_fields, endpoint='level_one_quote')
        fields.sort(key=int)

        # Build the request
        request = self._new_request_template()
        request['service'] = 'QUOTE'
        request['command'] = 'SUBS'
        request['parameters']['keys'] = ','.join(symbols)
        request['parameters']['fields'] = ','.join(fields)
        
        return json.dumps({"requests":[request]}, separators=(',', ':'))
        #return json.dumps({"requests":quote_requests}, separators=(',', ':'))

    def send(self, msg):
        if not msg:
            self.async_loop.call_soon_threadsafe(lambda: self.message_queue.put_nowait(None))
            return

        if msg['type'] == 'subscribe':
            self.async_loop.call_soon_threadsafe(lambda: self.sub_queue.put_nowait(msg))

    def close(self):
        self.shutdown = True
        self.send(None)
        while not self.conn.closed:
            time.sleep(0.3)
        self.async_loop.stop()
        self.async_loop.close()
        if self.async_thread:
            self.async_thread.join()

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

    parts = [API_ENDPOINT, API_VERSION, endpoint]
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


def _create_token_timestamp(token_timestamp: str) -> int:
    """Parses the token and converts it to a timestamp.
    
    Arguments:
    ----
    token_timestamp {str} -- The timestamp returned from the get_user_principals endpoint.
    
    Returns:
    ----
    int -- the token timestamp as an integer.
    """

    token_timestamp = datetime.strptime(token_timestamp, "%Y-%m-%dT%H:%M:%S%z")
    token_timestamp = int(token_timestamp.timestamp()) * 1000

    return token_timestamp

def create_td_client() -> TDClient:
    """Creates a new streaming session with the TD API.

    Grab the token to authenticate a stream session, builds
    the credentials payload, and initalizes a new instance
    of the TDStream client.

    Usage:
    ----
        >>> td_session = TDClient(
            client_id='<CLIENT_ID>',
            redirect_uri='<REDIRECT_URI>',
            credentials_path='<CREDENTIALS_PATH>'
        )
        >>> td_session.login()
        >>> td_stream_session = td_session.create_streaming_session()

    Returns:
    ----
    TDStreamerClient -- A new instance of a Stream Client that can be
        used to subscribe to different streaming services.
    """
    try:
        with open(TOKEN_FILE_NAME, 'r') as json_file:
            token_data = json.load(json_file)

        # Grab the Streamer Info.
        userPrincipalsResponse = get_user_principals(
            token_data,
            fields=['streamerConnectionInfo', 'streamerSubscriptionKeys', 'preferences', 'surrogateIds'])

        # Grab the timestampe.
        tokenTimeStamp = userPrincipalsResponse['streamerInfo']['tokenTimestamp']

        # Grab socket
        socket_url = userPrincipalsResponse['streamerInfo']['streamerSocketUrl']

        # Parse the token timestamp.
        tokenTimeStampAsMs = _create_token_timestamp(
            token_timestamp=tokenTimeStamp)

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

        # Create the session
        streaming_session = TDClient(
            websocket_url=socket_url,
            user_principal_data=userPrincipalsResponse, 
            credentials=credentials
        )
    except Exception as e:
        logging.debug(e)

    return streaming_session
