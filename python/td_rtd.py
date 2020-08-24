"""Excel RTD (RealTimeData) Server sample for real-time stock quote.
"""
import excel_rtd as rtd
import tdapi as ta
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

from typing import List

LOG_FILE_FOLDER = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'logs')
LOG_FILENAME = os.path.join(LOG_FILE_FOLDER, 'TD_{:%Y%m%d_%H%M%S}.log'.format(datetime.now()))

if not os.path.exists(LOG_FILE_FOLDER):
    os.makedirs(LOG_FILE_FOLDER)

logging.basicConfig(
    filename=LOG_FILENAME,
    level=logging.ERROR,
    format="%(asctime)s:%(levelname)s:%(message)s"
)

class TDServer(rtd.RTDServer):
    _reg_clsid_ = '{E28CFA65-CC94-455E-BF49-DCBCEBD17154}'
    _reg_progid_ = 'TD.RTD'
    _reg_desc_ = "RTD server for realtime stock quote"

    # other class attributes...

    def __init__(self):
        super(TDServer, self).__init__()
        self.td_cli = ta.create_td_client()
        self.start_conn_event = threading.Event()
        self.async_loop = None

        self.topics_by_key = {}

        self.update_thread = threading.Thread(target=self.update_thread_handler)
        self.shutdown = False

    def OnServerStart(self):
        logging.info("OnServerStart Begin")

        self.update_thread.start()
        while not self.async_loop:
            time.sleep(0.1)

    def OnServerTerminate(self):
        logging.info("OnServerTerminate Begin")

        self.shutdown = True

        if self.td_cli:
            self.td_cli.close()
            self.td_cli = None

        if not self.start_conn_event.is_set():
            self.start_conn_event.set()

        if not self.ready_to_send.is_set():
            self.ready_to_send.set()

        self.start_conn_event.clear()
        self.ready_to_send.clear()

    def _on_quote_received(self, quotes: List[ta.TDQuote]) -> None:
        self.async_loop.call_soon_threadsafe(lambda: self.update_message_queue.put_nowait(quotes))

    def update_thread_handler(self) -> None:
        logging.info("update_thread_handler start")
        try:
            pythoncom.CoInitializeEx(pythoncom.COINIT_MULTITHREADED)

            asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            self.async_loop = loop
            self.update_message_queue = asyncio.Queue(loop=self.async_loop)
            self.send_message_queue = asyncio.Queue(loop=self.async_loop)
            self.ready_to_send = asyncio.Event(loop=self.async_loop)

            # Following call can cause deadlock if mainthread is not pumping Windows message.
            self.SetCallbackThread()

            update_msg_coro = self._update_msg_handler()
            send_msg_coro = self._send_msg_handler()
            loop.run_until_complete(asyncio.gather(update_msg_coro, send_msg_coro))
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

        self.start_conn_event.wait()
        if self.shutdown:
            return
        
        self.td_cli.connect(self._on_quote_received)
        self.ready_to_send.set()
        logging.debug("_update_msg_handler: ready_to_send.set()")

        while not self.shutdown:
            quotes = await self.update_message_queue.get()
            
            try:
                # Check if any of our topics have new info to pass on
                if not len(self.topics):
                    pass

                for quote in quotes:
                    ticker = quote.ticker

                    for k, v in quote.fields.items():
                        if (ticker, k) in self.topics_by_key:
                            topic = self.topics_by_key[(ticker, k)]
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
    
    async def _send_msg_handler(self) -> None:
        self.ready_to_send.wait()
        logging.debug(f"_send_msg_handler: ready_to_send signalled")
        if self.shutdown:
            return

        while not self.shutdown:
            msg = await self.send_message_queue.get()
            if msg:
                self.td_cli.send(msg)

    def CreateTopic(self, TopicId,  TopicStrings=None):
        """Topic factory. Builds a StockTickTopic object out of the given TopicStrings."""
        if len(TopicStrings) >= 2:
            ticker, field = TopicStrings
            logging.info(f"CreateTopic {TopicId}, {ticker}|{field}")
            if not ticker:
                return None
            
            if not self.start_conn_event.is_set():
                self.start_conn_event.set()

            new_topic = StockTickTopic(TopicId, TopicStrings)
            ticker = ticker.upper()
            self.topics_by_key[(ticker, field)] = new_topic
            
            subscribe_msg = {
                "type": "subscribe",
                "symbol": ticker,
                "field": field
            }

            logging.debug(subscribe_msg)
            try:
                self.async_loop.call_soon_threadsafe(lambda: self.send_message_queue.put_nowait(subscribe_msg))
            except Exception as e:
                logging.error("CreateTopic: {}".format(repr(e)))
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
        self.SetValue("#WatingDataForData")

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

    # Register/Unregister TDServer example
    # eg. at the command line: td_rtd.py --register
    # Then type in an excel cell something like:
    # =RTD("TD.RTD","","MSFT","last-price")
    
    win32com.server.register.UseCommandLine(TDServer)
