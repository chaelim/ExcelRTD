# Originally from https://github.com/mhammond/pywin32/blob/master/com/win32com/demos/excelRTDServer.py and slightly modified.
# - Added support for calling SignalExcel from a different python worker thread, not from the Excel's main thread.
# - RefreshData in RTDServer only sends new values in "self.updatedTopics" list to Excel.

"""Excel IRTDServer implementation.

This module is a functional example of how to implement the IRTDServer interface
in python, using the pywin32 extensions. Further details, about this interface
and it can be found at:
     http://msdn.microsoft.com/library/default.asp?url=/library/en-us/dnexcl2k2/html/odc_xlrtdfaq.asp
"""

# Copyright (c) 2003-2004 by Chris Nilsson <chris@slort.org>
#
# By obtaining, using, and/or copying this software and/or its
# associated documentation, you agree that you have read, understood,
# and will comply with the following terms and conditions:
#
# Permission to use, copy, modify, and distribute this software and
# its associated documentation for any purpose and without fee is
# hereby granted, provided that the above copyright notice appears in
# all copies, and that both that copyright notice and this permission
# notice appear in supporting documentation, and that the name of
# Christopher Nilsson (the author) not be used in advertising or publicity
# pertaining to distribution of the software without specific, written
# prior permission.
#
# THE AUTHOR DISCLAIMS ALL WARRANTIES WITH REGARD
# TO THIS SOFTWARE, INCLUDING ALL IMPLIED WARRANTIES OF MERCHANT-
# ABILITY AND FITNESS.  IN NO EVENT SHALL THE AUTHOR
# BE LIABLE FOR ANY SPECIAL, INDIRECT OR CONSEQUENTIAL DAMAGES OR ANY
# DAMAGES WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS,
# WHETHER IN AN ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS
# ACTION, ARISING OUT OF OR IN CONNECTION WITH THE USE OR PERFORMANCE
# OF THIS SOFTWARE.

import pythoncom
import threading
import win32api
import win32com.client
import logging
from win32com import universal
from win32com.client import gencache
from win32com.server.exception import COMException

# Typelib info for version 10 - aka Excel XP.
# This is the minimum version of excel that we can work with as this is when
# Microsoft introduced these interfaces.
EXCEL_TLB_GUID = '{00020813-0000-0000-C000-000000000046}' # LIBID_Excel
EXCEL_TLB_LCID = 0
EXCEL_TLB_MAJOR = 1
EXCEL_TLB_MINOR = 4

IID_RTDUpdateEvent = '{A43788C1-D91B-11D3-8F39-00C04F3651B8}'

# Import the excel typelib to make sure we've got early-binding going on.
# The "ByRef" parameters we use later won't work without this.
gencache.EnsureModule(EXCEL_TLB_GUID, EXCEL_TLB_LCID, \
                      EXCEL_TLB_MAJOR, EXCEL_TLB_MINOR)

# Tell pywin to import these extra interfaces.
# --
# QUESTION: Why? The interfaces seem to descend from IDispatch, so
# I'd have thought, for example, calling callback.UpdateNotify() (on the
# IRTDUpdateEvent callback excel gives us) would work without molestation.
# But the callback needs to be cast to a "real" IRTDUpdateEvent type. Hmm...
# This is where my small knowledge of the pywin framework / COM gets hazy.
# --
# Again, we feed in the Excel typelib as the source of these interfaces.
universal.RegisterInterfaces(EXCEL_TLB_GUID,
                             EXCEL_TLB_LCID, EXCEL_TLB_MAJOR, EXCEL_TLB_MINOR,
                             ['IRtdServer','IRTDUpdateEvent'])

class RTDServer(object):
  """Base RTDServer class.

  Provides most of the features needed to implement the IRtdServer interface.
  Manages topic adding, removal, and packing up the values for excel.

  Shouldn't be instanciated directly.

  Instead, descendant classes should override the CreateTopic() method.
  Topic objects only need to provide a GetValue() function to play nice here.
  The values given need to be atomic (eg. string, int, float... etc).

  Also note: nothing has been done within this class to ensure that we get
  time to check our topics for updates. I've left that up to the subclass
  since the ways, and needs, of refreshing your topics will vary greatly. For
  example, the sample implementation uses a timer thread to wake itself up.
  Whichever way you choose to do it, your class needs to be able to wake up
  occasionally, since excel will never call your class without being asked to
  first.

  Excel will communicate with our object in this order:
    1. Excel instantiates our object and calls ServerStart, providing us with
       an IRTDUpdateEvent callback object.
    2. Excel calls ConnectData when it wants to subscribe to a new "topic".
    3. When we have new data to provide, we call the UpdateNotify method of the
       callback object we were given.
    4. Excel calls our RefreshData method, and receives a 2d SafeArray (row-major)
       containing the Topic ids in the 1st dim, and the topic values in the
       2nd dim.
    5. When not needed anymore, Excel will call our DisconnectData to
       unsubscribe from a topic.
    6. When there are no more topics left, Excel will call our ServerTerminate
       method to kill us.

  Throughout, at undetermined periods, Excel will call our Heartbeat
  method to see if we're still alive. It must return a non-zero value, or
  we'll be killed.

  NOTE: By default, excel will at most call RefreshData once every 2 seconds.
        This is a setting that needs to be changed excel-side. To change this,
        you can set the throttle interval like this in the excel VBA object model:
          Application.RTD.ThrottleInterval = 1000 ' milliseconds
  """
  _com_interfaces_ = ['IRtdServer']
  _public_methods_ = ['ConnectData','DisconnectData','Heartbeat',
                      'RefreshData','ServerStart','ServerTerminate']
  _reg_clsctx_ = pythoncom.CLSCTX_INPROC_SERVER
  #_reg_clsid_ = "# subclass must provide this class attribute"
  #_reg_desc_ = "# subclass should provide this description"
  #_reg_progid_ = "# subclass must provide this class attribute"

  ALIVE = 1
  NOT_ALIVE = 0

  def __init__(self):
    """Constructor"""
    super(RTDServer, self).__init__()
    self.IsAlive = self.ALIVE
    self.__marshall_callback = None
    self.__callback  = None
    self.__excel_thread_id = None
    self.topics = {}
    self.updatedTopics = {}

  def SignalExcel(self):
    """Use the callback we were given to tell excel new data is available."""
    if self.__callback is None:
        raise COMException(desc="Callback excel provided is Null")
    self.__callback.UpdateNotify()

  def ConnectData(self, TopicID, Strings, GetNewValues):
    """Creates a new topic out of the Strings excel gives us."""
    try:
      self.topics[TopicID] = self.CreateTopic(TopicID, Strings)
    except Exception as e:
      raise COMException(desc=repr(e))

    """
    If this is called during the file load, GetNewValues will be False and doesn't need to send Excel new data. 
    If we change "GetNewValues" to True on file load, we have to send back new data. If we don't send a new value, it will result in a #N/A.
    GetNewValues input value is False if the ConnectData is called while user is opening a file with existing RTD formulas in it.
    In other cases (e.g. user entered a new RTD formula) it will be True.
    """
    if GetNewValues:
      result = self.topics[TopicID]
      if result is None:
        result = "# %s: Waiting for update" % self.__class__.__name__
      else:
        result = result.GetValue()
    else:
      # Tell Excel use a cached value
      # Note: Excel can use the cached value only for the binary (.xlb and .xls) format files.
      result = None

    # fire out internal event...
    self.OnConnectData(TopicID)

    # GetNewValues as per interface is ByRef, so we need to pass it back too.
    return result, GetNewValues

  def DisconnectData(self, TopicID):
    """Deletes the given topic."""
    self.OnDisconnectData(TopicID)

    if TopicID in self.topics:
      self.topics[TopicID] = None
      del self.topics[TopicID]

  def Heartbeat(self):
    """Called by excel to see if we're still here."""
    return self.IsAlive

  def RefreshData(self, TopicCount):
    """Packs up the topic values. Called by excel when it's ready for an update.

    Needs to:
      * Return the current number of topics, via the "ByRef" TopicCount
      * Return a 2d SafeArray of the topic data.
        - 1st dim: topic numbers
        - 2nd dim: topic values
    """

    # Excel expects a 2-dimensional array. The first dim contains the
    # topic numbers, and the second contains the values for the topics.
    # In true VBA style (yuck), we need to pack the array in row-major format,
    # which looks like:
    #   ( (topic_num1, topic_num2, ..., topic_numN), \
    #     (topic_val1, topic_val2, ..., topic_valN) )
    topicIDs = []
    topicValues = []
    updatedTopics = self.updatedTopics
    self.updatedTopics = {}

    self.OnRefreshData()

    for topicID, topicValue in updatedTopics.items():
        topicIDs.append(topicID)
        topicValues.append(topicValue)

    results = [topicIDs, topicValues]
    TopicCount = len(topicIDs)

    # TopicCount is meant to be passed to us ByRef, so return it as well, as per
    # the way pywin32 handles ByRef arguments.
    return tuple(results), TopicCount

  def ServerStart(self, CallbackObject):
    """Excel has just created us... We take its callback for later, and set up shop."""
    self.IsAlive = self.ALIVE

    if CallbackObject is None:
      raise COMException(desc='Excel did not provide a callback')

    try:
      self.__excel_thread_id = threading.current_thread().ident
      
      # Need to "cast" the raw PyIDispatch object to the IRTDUpdateEvent interface
      IRTDUpdateEventKlass = win32com.client.CLSIDToClass.GetClass(IID_RTDUpdateEvent)
      self.__callback = IRTDUpdateEventKlass(CallbackObject)

      # Prepare for marshalling callback object because we're going to likely call SignalExcel in a different thread.
      self.__marshall_callback = pythoncom.CoMarshalInterThreadInterfaceInStream(pythoncom.IID_IDispatch, self.__callback)

      self.OnServerStart()
    except Exception as e:
      logging.error("ServerStart: {}".format(repr(e)))

    logging.info("ServerStart Done")
    return self.IsAlive

  def ServerTerminate(self):
    """Called when excel no longer wants us."""
    self.IsAlive = self.NOT_ALIVE # On next heartbeat, excel will free us
    self.OnServerTerminate()

  def CreateTopic(self, TopicId, TopicStrings=None):
    """Topic factory method. Subclass must override.

    Topic objects need to provide:
      * GetValue() method which returns an atomic value.

    Will raise NotImplemented if not overridden.
    """
    raise NotImplemented('Subclass must implement')

  def SetCallbackThread(self):
    if self.__marshall_callback is None:
      raise COMException(desc='self.__marshall_callback is not initialized')

    self.__callback = win32com.client.Dispatch (
      pythoncom.CoGetInterfaceAndReleaseStream (
          self.__marshall_callback, 
          pythoncom.IID_IDispatch
      )
    )

    logging.info("SetCallbackThread Done")

  # Overridable class events...
  def OnConnectData(self, TopicID):
    """Called when a new topic has been created, at excel's request."""
    pass
  def OnDisconnectData(self, TopicID):
    """Called when a topic is about to be deleted, at excel's request."""
    pass
  def OnRefreshData(self):
    """Called when excel has requested all current topic data."""
    pass
  def OnServerStart(self):
    """Called when excel has instanciated us."""
    pass
  def OnServerTerminate(self):
    """Called when excel is about to destroy us."""
    pass

class RTDTopic(object):
  """Base RTD Topic.
  Only method required by our RTDServer implementation is GetValue().
  The others are more for convenience."""
  def __init__(self, TopicStrings):
    super(RTDTopic, self).__init__()
    self.TopicStrings = TopicStrings
    self.__currentValue = None
    self.__dirty = False

  def Update(self, sender):
    """Called by the RTD Server.
    Gives us a chance to check if our topic data needs to be
    changed (eg. check a file, quiz a database, etc)."""
    raise NotImplemented('subclass must implement')

  def Reset(self):
    """Call when this topic isn't considered "dirty" anymore."""
    self.__dirty = False

  def GetValue(self):
    return self.__currentValue

  def SetValue(self, value):
    self.__dirty = True
    self.__currentValue = value

  def HasChanged(self):
    return self.__dirty
