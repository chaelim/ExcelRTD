"""Excel RTD (RealTimeData) Server sample for real-time stock quote.

Periodically polling stock quote from the stockrow.com.
"""

import excel_rtd as rtd
from datetime import datetime
import threading
import pythoncom
import win32api
import win32com.client
from win32com.server.exception import COMException
import logging
import os
import time
import asyncio
import json
from aiohttp import ClientSession

LOG_FILE_FOLDER = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'logs')
LOG_FILENAME = os.path.join(LOG_FILE_FOLDER, 'Stockrow_{:%Y%m%d_%H%M%S}.log'.format(datetime.now()))

if not os.path.exists(LOG_FILE_FOLDER):
    os.makedirs(LOG_FILE_FOLDER)

logging.basicConfig(
    filename=LOG_FILENAME,
    level=logging.INFO,
    format="%(asctime)s:%(levelname)s:%(message)s"
)

# Update frequency values
_DEF_UPDATE_FREQ_SEC = 30
_MIN_UPDATE_FREQ_SEC = 10
_MAX_UPDATE_FREQ_SEC = 60 * 10

# Max tickers in one request.
_STOCKROW_TICKERS_IN_ONE_REQUESTS = 30

class StockrowServer(rtd.RTDServer):
    _reg_clsid_ = '{C38E586E-C6B9-41B2-9C99-88180E2B9DB8}'
    _reg_progid_ = 'STOCKROW'
    _reg_desc_ = "RTD server for realtime stock quote using stockrow.com"

    # other class attributes...

    def __init__(self):
        super(StockrowServer, self).__init__()
        self.topics_by_key = {}

        self.update_thread = threading.Thread(target = self.update_thread_handler)
        self.shutdown = False
        self.async_loop = None
        self.update_freq_sec = _DEF_UPDATE_FREQ_SEC

    def OnServerStart(self):
        logging.info("OnServerStart Begin")

        self.update_thread.start()
        while not self.async_loop:
            time.sleep(0.1)

    def OnServerTerminate(self):
        logging.info("OnServerTerminate Begin")
        self.shutdown = True

    #def OnRefreshData(self):
    #    """Called when excel has requested refresh topic data."""

    def update_thread_handler(self) -> None:
        logging.info("update_thread_handler start")
        try:
            pythoncom.CoInitializeEx(pythoncom.COINIT_MULTITHREADED)

            asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            self.async_loop = loop
            self.update_message_queue = asyncio.Queue(loop=self.async_loop)
            self.stock_quote_queue = asyncio.Queue(loop=self.async_loop)

            # Following call can cause deadlock if mainthread is not pumping Windows message.
            self.SetCallbackThread()

            update_msg_coro = self._update_msg_handler()
            auto_update_coro = self._auto_update_worker()
            stock_quote_coro = self._get_stock_quote_worker()
            loop.run_until_complete(asyncio.gather(update_msg_coro, auto_update_coro, stock_quote_coro))
            loop.close()
        except Exception as e:
            logging.error("update_thread_handler: {}".format(repr(e)))
        finally:
            pythoncom.CoUninitialize()

    #
    # _update_msg_handler coro
    #
    async def _update_msg_handler(self) -> None:
        logging.debug("_update_msg_handler: start")

        if self.shutdown:
            return
        
        while not self.shutdown:
            msgs = await self.update_message_queue.get()
            
            try:
                # Check if any of our topics have new info to pass on
                if not len(self.topics):
                    pass

                # {"MSFT|last_price":166.79,"MSFT|volume":21266168,"MSFT|last_update_time":"09:48:54","CSCO|last_price":40.519,"CSCO|volume":11838414,"CSCO|last_update_time":"09:48:54","QCOM|volume":4211871,"HPQ|bid":21.12,"COF|volume":1096158,"DIS|volume":6656830}
                if msgs:
                    for k, v in msgs.items():
                        logging.debug(f"dequeue: {k} {v}")
                        ticker, field = k.split('|')
                        if (ticker, field) in self.topics_by_key:
                            topic = self.topics_by_key[(ticker, field)]
                            topic.Update(v)

                            if topic.HasChanged():
                                self.updatedTopics[topic.topicID] = topic.GetValue()

                if self.updatedTopics:
                    # Retry when com_error occurs 
                    # e.g. RPC_E_SERVERCALL_RETRYLATER = com_error(-2147417846, 'The message filter indicated that the application is busy.', None, None)
                    while True:
                        try:
                            self.SignalExcel()
                            break
                        except pythoncom.com_error as error:
                            await asyncio.sleep(0.01)

            except Exception as e:
                logging.error("Update: {}".format(repr(e)))
                #raise COMException(desc=repr(e))
    
    async def _auto_update_worker(self):
        while not self.shutdown:
            await asyncio.sleep(self.update_freq_sec)
            while not self.stock_quote_queue.empty():
                await asyncio.sleep(0.5)

            for (ticker, field) in self.topics_by_key.keys():
                if field:
                    self.stock_quote_queue.put_nowait(ticker)

    @staticmethod
    async def _fetch(url, session) -> str:
        async with session.get(url) as response:
            return await response.read()

    async def _get_stock_quote_worker(self) -> None:
        async with ClientSession() as session:
            while not self.shutdown:
                tickers = []
                tickers.append(await self.stock_quote_queue.get())

                for _ in range(_STOCKROW_TICKERS_IN_ONE_REQUESTS - 1):
                    if self.stock_quote_queue.empty():
                        break
                    tickers.append(await self.stock_quote_queue.get())

                url_params = str()
                for ticker in tickers:
                    url_params += f"tickers[]={ticker}&"
                
                url = "https://stockrow.com/api/price_changes.json?" + url_params.rstrip("&")

                while True:
                    logging.debug(f"fetch: {url}")
                    response_str = await self._fetch(url, session)
                    # [{"ticker":"AAPL","price":383.6800,"ohlc":[381.3400,383.8800,378.8310,383.6800,21827181.0000],"absolute_change":0.6700,"relative_change":0.0017,"date":"10 Jul"},
                    logging.debug(response_str)
                    try:
                        updates = {}
                        price_changes = json.loads(response_str)
                        for price_change in price_changes:
                            #dt = datetime.fromtimestamp(response['t'])
                            #sq_time = dt.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]
                            ticker = price_change["ticker"]
                            price = price_change["price"]
                            ohlc = price_change["ohlc"]
                            absolute_change = price_change["absolute_change"]
                            relative_change = price_change["relative_change"]
                            update_date = price_change["date"]

                            # Generate updates
                            updates[f"{ticker}|last_price"] = price
                            updates[f"{ticker}|open"] = ohlc[0]
                            updates[f"{ticker}|high"] = ohlc[1]
                            updates[f"{ticker}|low"] = ohlc[2]
                            updates[f"{ticker}|close"] = ohlc[3]
                            updates[f"{ticker}|volume"] = ohlc[4]
                            updates[f"{ticker}|absolute_change"] = absolute_change
                            updates[f"{ticker}|relative_change"] = relative_change
                            updates[f"{ticker}|last_update_date"] = update_date
                            #updates[f"{ticker}|last_update_time"] = sq_time
                        self.async_loop.call_soon_threadsafe(lambda: self.update_message_queue.put_nowait(updates))
                        break
                    except Exception as e:
                        logging.error(f"{repr(e)}: {response_str}")
                        await asyncio.sleep(0.5)
                        continue

    def CreateTopic(self, TopicId,  TopicStrings=None):
        """Topic factory. Builds a StockTickTopic object out of the given TopicStrings."""
        if len(TopicStrings) >= 2:
            ticker, field = TopicStrings
            logging.debug(f"CreateTopic {TopicId}, {ticker}|{field}")
            if not ticker:
                return None

            if ticker == "set_update_frequency":
                self.update_freq_sec = max(min(float(field), _MAX_UPDATE_FREQ_SEC), _MIN_UPDATE_FREQ_SEC)
                logging.info(f"set_update_frequency: {self.update_freq_sec}")

                new_topic = SimpeVarTopic(TopicId, TopicStrings)
                self.topics_by_key[(ticker, None)] = new_topic
            else:
                new_topic = StockTickTopic(TopicId, TopicStrings)
                ticker = ticker.upper()
                self.topics_by_key[(ticker, field)] = new_topic

                self.async_loop.call_soon_threadsafe(lambda: self.stock_quote_queue.put_nowait(ticker))
        else:
            logging.error(f"Unknown param: CreateTopic {TopicId}, {TopicStrings}")
            return None
        return new_topic

class SimpeVarTopic(rtd.RTDTopic):
    def __init__(self, topicID, TopicStrings):
        super(SimpeVarTopic, self).__init__(TopicStrings)
        try:
            cmd, var = self.TopicStrings
            self.topicID = topicID
        except Exception as e:
            raise ValueError("Invalid topic strings: %s" % str(TopicStrings))

        # setup our initial value
        self.checkpoint = self.timestamp()
        self.SetValue(var)

    def timestamp(self):
        return datetime.now()

    def Update(self, value):
        self.SetValue(value)
        self.checkpoint = self.timestamp()

class StockTickTopic(rtd.RTDTopic):
    """Stock quote topic
    """
    def __init__(self, topicID, TopicStrings):
        super(StockTickTopic, self).__init__(TopicStrings)
        try:
            ticker, field = self.TopicStrings
            self.topicID = topicID
            self.ticker = ticker
            self.field = field
        except Exception as e:
            raise ValueError("Invalid topic strings: %s" % str(TopicStrings))

        # setup our initial value
        self.checkpoint = self.timestamp()
        self.SetValue("N/A")

    def __key(self):
        return (self.ticker, self.field)

    def __hash__(self):
        return hash(self.__key())

    def __eq__(self, other):
        if isinstance(other, StockTickTopic):
            return self.__key() == other.__key()
        return NotImplemented

    def timestamp(self):
        return datetime.now()

    def Update(self, value):
        self.SetValue(value)
        self.checkpoint = self.timestamp()

if __name__ == "__main__":
    import win32com.server.register

    # Register/Unregister StockrowServer example
    # eg. at the command line: py stockrow_rtd.py --register
    # Then type in an excel cell something like:
    # =RTD("STOCKROW","","MSFT","last_price")
    
    win32com.server.register.UseCommandLine(StockrowServer)
