# Prerequisites

## Excel version

* __IMPORTANT__: Office 365 Version 1909 (16.0.12130.20XXX) or later is required.
* Due to the issue described [here](https://mail.python.org/pipermail/python-win32/2012-April/012207.html), it won't work in older builds.

## Python version 3.x for Windows

* Install it from https://www.python.org/downloads/windows
* Recommend install 64 bit (x64) build

# Setup

## Install pywin32

* go to the [pywin32 release site](https://github.com/mhammond/pywin32/releases)
* You should choose matching installation file with your Python version and bitness.
* For Python 3.8 64bit, use `pywin32-228.win-amd64-py3.8.exe`
* __IMPORTANT__: This folder  `<Your Python Installation Path>\Lib\site-packages\pywin32_system32` need to be added to the `PATH` environment variable.
    * Alternatively (This is what I'm doing), you can copy the two DLL files (`pythoncom38.dll` and `pywintypes38.dll`) under the `<Your Python Installation Path>\Lib\site-packages\pywin32_system32` to `<Your Python Installation Path>` assuming that is already in the `PATH` environment variable.
    * If you didn't change Python installation folder from the Python Installer, default Python installation folder would be `C:\User\<Your UserName>\AppData\Local\Programs\Python\Python38` or `%LOCALAPPDATA%\Programs\Python\Python38`.
    * If your Python application path was not added to the `PATH` environment variable at the Python installation time, you can follow instructions in [this web page](https://datatofish.com/add-python-to-windows-path/) to add Python application path to the `PATH` environment variable.

## Install required python packages

* Run `pip install -r requirements.txt` or `py -m pip install -r requirements.txt`

## How to register a Python RTD

* To install Python Excel RTD COM server:
    * Type `py <python file name> --register`
    * e.g. `py stockrow_rtd.py --register`

## Change `RTDThrottleInterval` to zero (recommended)

* By setting `RTDThrottleInterval` to zero, any update from the RTD COM server will be refreshed to Excel as quickly as possible.

* Use either of following ways to change `RTDThrottleInterval` value to zero
  1. Type this command line `reg add HKCU\SOFTWARE\Microsoft\Office\16.0\Excel\Options /v RTDThrottleInterval /t REG_DWORD /d 0 /f`
  2. Run `DisableThrottling.reg` (You can double click this from the explorer)

# Demo

## Stockrow RTD demo

![](demo/stockrow_rtd_demo1.gif)

## TD Ameritrade RTD Demo

> Look at [TD Ameritrade RTD](https://github.com/chaelim/ExcelRTD/blob/master/python/TDAmeritrade_RTD.md)
