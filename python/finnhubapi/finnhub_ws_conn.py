import logging
import asyncio
import threading
import websockets
import json
from datetime import datetime

import signal
import sys
import time

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
logger.addHandler(logging.StreamHandler())

class FinnhubWSConn(object):

    def __init__(self):
        """Constructor"""
        super(FinnhubWSConn, self).__init__()
        self.ws = None

    async def send(self, message) -> None:
        if self.ws:
            await self.ws.send(message)

    async def recv_message(self) -> str:
        message = await self.ws.recv()
        return message

    async def start(self, token) -> None:
        logger.debug("start")
        try:
            uri = f"wss://ws.finnhub.io?token={token}"
            self.ws = await websockets.connect(uri)
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

class FinnhubClient():
    def __init__(self):
        self.conn = FinnhubWSConn()
        self.async_thread = None
        self.async_loop = None
        self.message_queue = None
        self.shutdown = False
        self.token = None
        self.connect_event = threading.Event()

    def connect(self, on_recv_msg, token):
        self.on_recv_msg = on_recv_msg
        self.token = token
        self.async_thread = threading.Thread(target=self._async_thread_handler)
        self.async_thread.start()
        self.connect_event.wait()

    def _async_thread_handler(self) -> None:
        logger.debug("_async_thread_handler")
        try:
            asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            self.async_loop = loop
            self.message_queue = asyncio.Queue()
            loop.run_until_complete(self.conn.start(self.token))
            self.connect_event.set()

            send_messages_coro = self._send_message_async()
            recv_messages_coro = self._recv_message_async()
            loop.run_until_complete(asyncio.gather(send_messages_coro, recv_messages_coro))
            loop.close()
        except Exception as e:
            logger.error("_async_thread_handler: {}".format(repr(e)))
            self.connect_event.set()

    async def _send_message_async(self) -> None:
        while not self.shutdown:
            message = await self.message_queue.get()
            if message:
                await self.conn.send(message)
            else:
                await self.conn.close()

    async def _recv_message_async(self) -> None:
        while not self.shutdown:
            msg = await self.conn.recv_message()
            self.on_recv_msg(msg)

    def send(self, msg):
        self.async_loop.call_soon_threadsafe(lambda: self.message_queue.put_nowait(msg))
    
    def close(self):
        self.shutdown = True
        self.send(None)
        while not self.conn.closed:
            time.sleep(0.3)
        self.async_loop.stop()
        self.async_loop.close()
        if self.async_thread:
            self.async_thread.join()

def signal_handler(sig, frame):
    print('You pressed Ctrl+C!')
    global finnhub_client
    finnhub_client.close()
    sys.exit(0)

def on_recv_message(message) -> None:
    # e.g. {"data":[{"p":379.6,"s":"AAPL","t":1594228987324,"v":21}],"type":"trade"}
    response = json.loads(message)
    if response['type'] == 'trade':
        trades = response['data']
        for trade in trades:
            dt = datetime.fromtimestamp(trade['t']/1000)
            tr_time = dt.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]
            tr_ticker = trade['s']
            tr_price = trade['p']
            tr_volume = trade['v']
            print(f'{tr_time}: Ticker={tr_ticker}, Price={tr_price}, Volume={tr_volume}')
    else:
        # e.g. {"type":"ping"}
        print(message)

def finnhub_test_main():
    signal.signal(signal.SIGINT, signal_handler)

    global finnhub_client
    finnhub_client = FinnhubClient()
    finnhub_client.connect(on_recv_message, "YOUR_FINNHUB_TOKEN")
    finnhub_client.send('{"type":"subscribe","symbol":"MSFT"}')
    finnhub_client.send('{"type":"subscribe","symbol":"BINANCE:BTCUSDT"}')
    while True:
        time.sleep(1)

################################################################################
#   __main__
################################################################################
if __name__ == '__main__':
    finnhub_test_main()
