#--Final Copy of TRADE_OC4.py
import json
import os
import os.path
import sys
import copy
import time
import math
import ast
import threading
import glob
try:
    import pandas as pd
except ImportError:
    print("\nError : Microsoft Visual C++ 2019 Redistributable files are required\n"
            "x64: https://aka.ms/vs/16/release/vc_redist.x64.exe\n"
            "x86: https://aka.ms/vs/16/release/vc_redist.x86.exe")
    input()
    sys.exit()
import numpy as np
from ping3 import ping              #--for RTT and SD-RTT
import dateutil.parser
import winsound
from multiprocessing import Process,freeze_support,Manager
from datetime import datetime,timedelta,date
from tkinter import Label,DISABLED,CENTER,NORMAL
import tkinter as tk
from tkinter import simpledialog
from tkinter.messagebox import showinfo
from tkinter import ttk             #--Added for OC
from tkinter import PhotoImage
# from tkcalendar import DateEntry
from PIL import Image, ImageTk
from kiteconnect import KiteTicker #--Required for Websocket

import psutil
# import pymysql

import socket
socket.getaddrinfo('localhost', 8080)

import warnings
warnings.filterwarnings("ignore")
warnings.simplefilter(action='ignore', category=FutureWarning)

import requests
import dateutil
import dateutil.parser

from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph
# import gc
import re
import csv
#------------------------
import logging
class SuppressConnectionErrorFilter(logging.Filter): #--Define a custom filter to suppress specific error messages
    def filter(self, record):
        #--Filter out messages containing specific text
        if ("Connection error: 1006" in record.getMessage() or
             "Connection closed: None" in record.getMessage() or
             "Connection closed: 1006" in record.getMessage()):
            return False
        return True
logger = logging.getLogger("kiteconnect.ticker")    #--Get the specific logger and apply the filter
logger.addFilter(SuppressConnectionErrorFilter())
logger.setLevel(logging.ERROR)                      #--Set the logging level (if necessary)
#------------------------
cwd = os.getcwd() #--Get Current working directory path

def Trade_Log_TimeStamp():
    now = datetime.now()
    year = now.strftime("%Y")
    month = now.strftime("%m")
    day = now.strftime("%d")
    hour = now.strftime("%H")
    minutes = now.strftime("%M")
    seconds = now.strftime("%S")
    formatted_datetime = f"{year}{month}{day}_{hour}{minutes}{seconds}"
    return formatted_datetime

#===== Create Log Folder in Root ==============
path = os.path.join(os.getcwd(), 'Log')
if not os.path.exists(path):os.makedirs(path)
TxtLogFilePath = os.path.join(cwd, 'Log', 'Log_OTM.txt')
PdfLogFilePath = os.path.join(cwd, 'Log', f"Log_OTM_{Trade_Log_TimeStamp()}.pdf")
TxtSettingPath = os.path.join(cwd, 'Log', 'OTM_Settings.txt')
Cookiesfile_path = os.path.join(os.getenv("USERPROFILE"),"Downloads","cookies.txt")
Initial_Strikes_Path = os.path.join(cwd, 'Initial_Strikes.csv')
#===========================================
keep_scheduled = None
#===========================================
class Logger(object):
    def __init__(self):
        self.terminal = sys.stdout
        if (os.path.exists(TxtLogFilePath)):os.remove(TxtLogFilePath)
        self.log = open(TxtLogFilePath, "w")

    def write(self, message):
        self.terminal.write(message)
        self.log.write(message)

    def close(self):
        self.log.close()

    def flush(self):
        #--this flush method is needed for python 3 compatibility.
        #--this handles the flush command by doing nothing.
        #--you might want to specify some extra behavior here.
        pass
class KiteApp:
    #--Products
    PRODUCT_MIS = "MIS"
    PRODUCT_CNC = "CNC"
    PRODUCT_NRML = "NRML"
    PRODUCT_CO = "CO"

    #--Order types
    ORDER_TYPE_MARKET = "MARKET"
    ORDER_TYPE_LIMIT = "LIMIT"
    ORDER_TYPE_SLM = "SL-M"
    ORDER_TYPE_SL = "SL"

    #--Varities
    VARIETY_REGULAR = "regular"
    VARIETY_CO = "co"
    VARIETY_AMO = "amo"

    #--Transaction type
    TRANSACTION_TYPE_BUY = "BUY"
    TRANSACTION_TYPE_SELL = "SELL"

    #--Validity
    VALIDITY_DAY = "DAY"
    VALIDITY_IOC = "IOC"

    #--Exchanges
    EXCHANGE_NSE = "NSE"
    EXCHANGE_BSE = "BSE"
    EXCHANGE_NFO = "NFO"
    EXCHANGE_CDS = "CDS"
    EXCHANGE_BFO = "BFO"
    EXCHANGE_MCX = "MCX"

    def __init__(self,enctoken,userid="      "):
        self.user_id = userid
        self.enctoken = enctoken
        self.headers = {"Authorization": f"enctoken {enctoken}"}
        self.session = requests.session()
        #self.root_url = "https://api.kite.trade"
        self.root_url = "https://kite.zerodha.com/oms"
        self.session.get(self.root_url, headers=self.headers)

    def kws(self):
        #return KiteTicker(api_key='kitefront', access_token=self.enctoken+"&user_id="+self.user_id, root='wss://ws.kite.trade')

        return KiteTicker(api_key='kitefront', access_token=self.enctoken+"&user_id="+self.user_id, root='wss://ws.kite.trade',
                          reconnect_max_tries=300, connect_timeout=60)

    def instruments(self, exchange=None):
        #data = self.session.get(f"{self.root_url}/instruments",headers=self.headers).text.split("\n")
        data = self.session.get("https://api.kite.trade/instruments").text.split("\n")
        Exchange = []
        for i in data[1:-1]:
            row = i.split(",")
            if exchange is None or exchange == row[11]:
                Exchange.append({'instrument_token': int(row[0]), 'exchange_token': row[1], 'tradingsymbol': row[2],
                                 'name': row[3][1:-1], 'last_price': float(row[4]),
                                 'expiry': dateutil.parser.parse(row[5]).date() if row[5] != "" else None,
                                 'strike': float(row[6]), 'tick_size': float(row[7]), 'lot_size': int(row[8]),
                                 'instrument_type': row[9], 'segment': row[10],
                                 'exchange': row[11]})
        return Exchange

    def quote(self, instruments):
        data = self.session.get(f"{self.root_url}/quote", params={"i": instruments}, headers=self.headers).json()["data"]
        return data

    def ltp(self, instruments):
        data = self.session.get(f"{self.root_url}/quote/ltp", params={"i": instruments}, headers=self.headers).json()["data"]
        return data

    def historical_data(self, instrument_token, from_date, to_date, interval, continuous=False, oi=False):
        params = {"from": from_date,
                  "to": to_date,
                  "interval": interval,
                  "continuous": 1 if continuous else 0,
                  "oi": 1 if oi else 0}
        lst = self.session.get(
            f"{self.root_url}/instruments/historical/{instrument_token}/{interval}", params=params,
            headers=self.headers).json()["data"]["candles"]
        records = []
        for i in lst:
            record = {"date": dateutil.parser.parse(i[0]), "open": i[1], "high": i[2], "low": i[3],
                      "close": i[4], "volume": i[5],}
            if len(i) == 7:
                record["oi"] = i[6]
            records.append(record)
        return records

    def profile(self):
        profile = self.session.get(f"{self.root_url}/user/profile", headers=self.headers).json()["data"]
        return profile

    def margins(self):
        margins = self.session.get(f"{self.root_url}/user/margins", headers=self.headers).json()["data"]
        return margins

    def order_margin(self,exchange,tradingsymbol,transaction_type,variety,product,order_type,quantity,price=None,trigger_price=None):
        params = locals()
        del params["self"]
        for k in list(params.keys()):
            if params[k] is None:
                del params[k]
        params = [params]  #--Wrap params in a list
        #print("---- Method 1 Parameters ------")
        #print(params)
        margin123 = self.session.post(f"{self.root_url}/margins/basket?consider_positions=true&mode=compact",
                                   headers=self.headers,json=params).json()["data"]
        return margin123

    def orders(self):
        orders = self.session.get(f"{self.root_url}/orders", headers=self.headers).json()["data"]
        return orders

    def positions(self):
        positions = self.session.get(f"{self.root_url}/portfolio/positions", headers=self.headers).json()["data"]
        return positions

    def place_order(self, variety, exchange, tradingsymbol, transaction_type, quantity, product, order_type, price=None,
                    validity=None, disclosed_quantity=None, trigger_price=None, squareoff=None, stoploss=None,
                    trailing_stoploss=None, tag=None):
        params = locals()
        del params["self"]
        for k in list(params.keys()):
            if params[k] is None:
                del params[k]
        order_id = self.session.post(f"{self.root_url}/orders/{variety}",
                                     data=params, headers=self.headers).json()["data"]["order_id"]
        return order_id

    def modify_order(self, variety, order_id, parent_order_id=None, quantity=None, price=None, order_type=None,
                     trigger_price=None, validity=None, disclosed_quantity=None):
        params = locals()
        del params["self"]
        for k in list(params.keys()):
            if params[k] is None:
                del params[k]

        order_id = self.session.put(f"{self.root_url}/orders/{variety}/{order_id}",
                                    data=params, headers=self.headers).json()["data"][
            "order_id"]
        return order_id

    def cancel_order(self, variety, order_id, parent_order_id=None):
        order_id = self.session.delete(f"{self.root_url}/orders/{variety}/{order_id}",
                                       data={"parent_order_id": parent_order_id} if parent_order_id else {},
                                       headers=self.headers).json()["data"]["order_id"]
        return order_id
class APH_GUI:
    Color1 = "#B9EEE9"            #--"#FFFFCC"
    R0C1flag = 'NRML'               #--MIS/NRML
    R1C1flag = 'LIMIT'              #--LIMIT / MARKET
    R0C5flag = "DEMO"              #--DEMO/REAL
    R2C4flag = "RUN"                #--Clear Trade (RUN/RESET) Flag
    R1C3_Value = 0.00              #--OpeningBalance
    R1C4_Value = 0.00              #--LiveBalance
    LPT_Value = 250
    PPT_Value = 750
    R3C2_Value = R3C4_Value = R3C6_Value = R3C8_Value = 1
    Trade_Flag = 'DEMO_ONLY'
    NBF = ""
    R2C11flag = "Stop Loss (SL)-20"
    R2C11_Value = 20
class APH_ARM:
    #---------------------------------
    def __init__(self, ns):
        self.ns = ns                    #--Store namespace inside instance

        self.std_str_file = None        # CSV Straddle Strangle backup setup
        self.std_str_csv_writer = None
        self.init_std_str_backup()

    def init_std_str_backup(self):
        """Initialize CSV file with headers."""
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"std_str_prices_backup_{timestamp}.csv"
        self.std_str_file = open(filename, 'w', newline='', encoding='utf-8')
        self.std_str_csv_writer = csv.writer(self.std_str_file)
        self.std_str_csv_writer.writerow([
            'Time', 'Std CE', 'Std PE', 'Std UDS', 'Str CE', 'Str PE','Str UDS', 'AVG_UDS', 'BUY_CE', 'BUY_PE'
        ])
        self.std_str_file.flush()  # Ensure write to disk without closing
        print(f"CSV backup started: {filename}")

    Initial_Strikes_df = pd.DataFrame(columns=["STRIKE", "CE_LTP", "PE_LTP"])
    Use_Initial_Strikes = False

    STRADDLE_DIALOG = None     # Global variable to track dialog instance

    # Straddle_Start_Time = datetime.strptime('09:19:59', '%H:%M:%S').time()
    Straddle_Start_Time = datetime.strptime('09:19:59', '%H:%M:%S').time()
    STD_STR_Exp = ""
    STD_STR_Subscribe = pd.DataFrame()

    STD_Dis_STD = 0
    STD_Dis_Time = ""
    STD_Dis_PBEP = 0
    STD_Dis_MBEP = 0

    STD_Strike = 0
    STD_CE_Symbol = ""
    STD_PE_Symbol = ""
    STD_CE_Price_Org = 0
    STD_PE_Price_Org = 0
    STD_CE_Price_Cur = 0
    STD_PE_Price_Cur = 0
    STD_CombPrice_Org = 0
    STD_CombinePrice_Cur = 0
    STD_CE_Per_Change = 0
    STD_PE_Per_Change = 0
    STD_Combine_Per_Change = 0

    STD_UDS_Cur = 0
    STR_UDS_Cur = 0
    AVG_UDS_Cur = 0
    AVG_UDS_Prev = 0

    # STD_UDS_df = pd.DataFrame(columns=['UDS'])  #--Not Implemented, so commented
    # STR_UDS_df = pd.DataFrame(columns=['UDS'])  #--Not Implemented, so commented
    AVG_UDS_df = pd.DataFrame(columns=['UDS'])

    # RISING_EDGE_df = pd.DataFrame(columns=['BUY_CE'])   #--Not Implemented, so commented
    # FALLING_EDGE_df = pd.DataFrame(columns=['BUY_PE'])  #--Not Implemented, so commented
    UDS_DEADBAND = 5
    STD_STR_MIN_diff = 0.4
    CONSECUTIVE_BAR_value = 3
    CONFIRM_BUYCE_Cur = False
    CONFIRM_BUYPE_Cur = False
    R_EDGE_Counter = 0
    F_EDGE_Counter = 0

    STR_CE_Strike = 0
    STR_PE_Strike = 0
    STR_CE_Symbol = ""
    STR_PE_Symbol = ""
    STR_CE_Price_Org = 0
    STR_PE_Price_Org = 0
    STR_CE_Price_Cur = 0
    STR_PE_Price_Cur = 0
    STR_CombPrice_Org = 0
    STR_CombinePrice_Cur = 0
    STR_CE_Per_Change = 0
    STR_PE_Per_Change = 0
    STR_Combine_Per_Change = 0

    Directional_Trend_UDS = True
    Directional_Trend_STD_STR = True
    Direction_Signal_On = 'COM' #--'IND', 'COM', 'HYB' Default
    Combine_Per = 5.0
    Individual_Per = 25.0

    STD_Strike_Change = False
    STD_Price_Change = False
    #---------------------------------
    ErrorExcptionInfo_scheduled = None
    BuySellSpread = 2
    myopenorder_status = False
    Zerodha_Orders_Flag = False
    ZDOrder_Info = False
    price_try = 0.0
    Order_WaitTime = 10
    My_DayOrders = []

    OC_Strikes = 10
    Display_Row = 5
    Script_YOE = 0                     #--Stores Days to Expiry
    KWS_SPT_Data = {}
    KWS_OPT_Data = {}
    KWS_OPT_Tokens = {}
    KWS_Subscribe_Tokens = {}
    New_Token_Ready = False
    New_SYM_Tokens = {}
    New_FUT_Tokens = {}

    Wait_KWS_Process = False
    kws_option_flag = False
    Kws_ReconnAttempt = Kws_DelayCounter = 0
    Kws_MaxReconAttempt = 5 #--Try 3 or 5
    KWS_Closed = False      #--Set when we close/terminate kws after internet restored
    KWS_Resterted = False   #--Set when we try to reconnect Zerodha

    Internet_NotConnected = False
    InternetLost_Once = True
    KWS_RestartCheck = KWS_URestart = False
    Delay_KWS_Restart = False
    KWS_Working = False
    KWS_Functional = True
    KWS_Check = 0
    ZLatency = 15

    Symbol_Changed = False
    Opt_Exp_Changed = False
    OptFut_Exp_Changed = False

    SPT_Token = FUT_Token = OPT_Token = ""
    df_opt = pd.DataFrame()

    Stop_Trade_Update = False
    Desplay_NetPosition_flag = False
    Target_SL_Hit_Flag = False
    Edit_SL_Target = False
    Spread_already_taken = False

    Buy_Symbol = ""
    Buy_CE_PE = ""
    Buy_AskPrice = 0
    Buy_SqOffPrice = 0
    Buy_PL_P = 0
    Buy_PL_Rs = 0
    Buy_Strike = 0
    Buy_MIS_NRML = "MIS"
    Buy_MARKET_LIMIT = "MARKET"
    Buy_Qty = 50

    Sell_Symbol = ""
    Sell_CE_PE = ""
    Sell_BidPrice = 0
    Sell_SqOffPrice = 0
    Sell_PL_P = 0
    Sell_PL_Rs = 0
    Sell_Strike = 0
    Sell_MIS_NRML = "MIS"
    Sell_MARKET_LIMIT = "MARKET"
    Sell_Qty = 50

    oc_symbol_SqOff = ""

    Buy_Symbol_SqOff = ""
    Buy_Strike_SqOff = 0
    Buy_CE_PE_SqOff = ""
    Buy_Exp_SqOff = ""
    Buy_AskPrice_SqOff = 0
    Buy_Qty_SqOff = 0
    Buy_MIS_NRML_SqOff = ""
    Buy_MARKET_LIMIT_SqOff = ""

    Sell_Symbol_SqOff = ""
    Sell_Strike_SqOff = 0
    Sell_CE_PE_SqOff = ""
    Sell_Exp_SqOff = ""
    Sell_BidPrice_SqOff = 0
    Sell_Qty_SqOff = 0
    Sell_MIS_NRML_SqOff = ""
    Sell_MARKET_LIMIT_SqOff = ""

    TotalBuyOrders = TotalSellOrders = 0
    BoughtPrice = 0
    SoldPrice = 0
    AvgBoughtPrice = 0 #--Not Used
    Opt_Bought1_df = pd.DataFrame(columns=['OrderNo','Time_In','Action','Price','Qty','SqOffPrice','Symbol','P/L(P)','P/L(Rs)','SL(P)','Target(P)','SQOff'])
    Opt_Bought2_df = pd.DataFrame(columns=['OrderNo','Time_In','Action','Price','Qty','SqOffPrice','Symbol','P/L(P)','P/L(Rs)','SL(P)','Target(P)','SQOff'])
    Opt_Sold1_df = pd.DataFrame(columns=['OrderNo','Time_In','Action','Price','Qty','SqOffPrice','Symbol','P/L(P)','P/L(Rs)','SL(P)','Target(P)','SQOff'])
    Opt_Sold2_df = pd.DataFrame(columns=['OrderNo','Time_In','Action','Price','Qty','SqOffPrice','Symbol','P/L(P)','P/L(Rs)','SL(P)','Target(P)','SQOff'])
    Opt_Bought3_df = pd.DataFrame(columns=['Time_In','Action', 'Avg_Price','Total_Qty','SqOffPrice','Symbol','P/L(P)','P/L(Rs)','SL(P)','Target(P)','SQOff'])
    Opt_Sold3_df = pd.DataFrame(columns=['Time_In','Action','Avg_Price','Total_Qty','SqOffPrice','Symbol','P/L(P)','P/L(Rs)','SL(P)','Target(P)','SQOff'])
    Display_NetBuySell_df = pd.DataFrame(columns=['Time_In','Action','Avg_Price','Total_Qty','SqOffPrice','Symbol','P/L(P)','P/L(Rs)','SL(P)','Target(P)','SQOff'])
    My_AllTrades = pd.DataFrame(columns=['Time_In','Time_Out','Action', 'Avg_Price','Total_Qty','SqOffPrice','Symbol','P/L(P)','P/L(Rs)'])
    NetBuySell_len = 0
    pev_num_rows = prev_num_columns = 0

    Option_Bought = {}
    Option_Sold = {}
    Option_NetBuySell = {}
    symbols_in_net_positions = {}

    CIO_Updated = False
    FNO_Symbol = []
    ARM_Version = "v415.0"
    ARM_Version_Date = "15Oct2024_v415.0" #--Not Used
    todays_server_date = None
    Trade_Start_Time = datetime.strptime('09:14:59', '%H:%M:%S').time()
    Trade_Stop_Time = datetime.strptime('15:30:00', '%H:%M:%S').time()
    ARM_Risk = 2.00             #--ARM Default Total Risk per Trade in %
    My_try = 0
    Etoken_File_Exists = False
    enctoken = User_Prv_Etoken = "Arvind"

    Opening_Balance = Live_Balance = 0.00
    Want_ChangeBal = OP_Bal_Changed = False #--Not Used

    APH_SharedPVar = {}

    oc_symbol = ""
    User_Type = ""
    MY_TRADE = ""
    your_user_id = ""

    oc_opt_expiry = ""
    pre_oc_symbol = pre_oc_opt_expiry = ""

    License_Type = "" #--Check this
    Login_selected_flag = True

    Program_Running = False
    NO_Response = 2

    mydayorders = []            #--Not Used
    opt_expiries_list = []
    instrument_dict_opt = {}

    prev_day_oi_opt = {} #--Added for OC
    prev_day_oi_opt_df = pd.DataFrame() #--Added for OC

    stop_thread_opt = stop_thread_fut = False #--Added for OC
    found_lotsize_opt = 0 #--Added for OC

    columns =["CE_COI","CE_OI","CE_VOL","CE_CLTP","CE_LTP","CE_Bid","CE_ATP","CE_Ask",
                "STRIKE",
                "PE_Ask","PE_ATP","PE_Bid","PE_LTP","PE_CLTP","PE_VOL","PE_OI","PE_COI"]
    df_opt_final = pd.DataFrame(columns=columns)

    Future_Ananysis = ""
    Ready_Once = True

    order_id_buy = None     #--Check for this variable
    order_id_sell = None    #--Check for this variable

    instrument_nse = pd.DataFrame()
    instrument_nfo = pd.DataFrame()
    fut_exchange_nfo = opt_exchange_nfo = None
    Exchange_Data_Process = None

    Spot_Value = 0.00
    NF_OPT_Trade_Symbol = ""

    NIFTY_LEG = False

#--------------------------------
import time
import tkinter as tk
from tkinter import DISABLED, NORMAL, CENTER

class StraddleDialog(tk.Toplevel):
    def __init__(self):
        super().__init__()

        self.title("Direction Settings")
        self.geometry("660x440+200+200")
        self.configure(bg="#D6D6D2")
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self.alive = True           # alive flag for .after loop
        self._after_id = None

        # copy values from ARM
        self.STD_Dis_STD_Value = ARM.STD_Dis_STD
        self.STD_Dis_Time_Value = ARM.STD_Dis_Time
        self.STD_Dis_PBEP = ARM.STD_Dis_PBEP
        self.STD_Dis_MBEP = ARM.STD_Dis_MBEP

        self.STD_Strike_Value = ARM.STD_Strike
        self.STD_CE_Org_Value = ARM.STD_CE_Price_Org
        self.STD_PE_Org_Value = ARM.STD_PE_Price_Org
        self.STD_CombPrice_Value = ARM.STD_CombPrice_Org

        # self.STD_UDS_df = ARM.STD_UDS_df  #--Not Implemented, so commented
        # self.STR_UDS_df = ARM.STR_UDS_df  #--Not Implemented, so commented
        self.AVG_UDS_df = ARM.AVG_UDS_df

        # self.RISING_EDGE_df = ARM.RISING_EDGE_df    #--Not Implemented, so commented
        # self.FALLING_EDGE_df = ARM.FALLING_EDGE_df  #--Not Implemented, so commented

        self.STR_CE_Strike_Value = ARM.STR_CE_Strike
        self.STR_PE_Strike_Value = ARM.STR_PE_Strike
        self.STR_CE_Org_Value = ARM.STR_CE_Price_Org
        self.STR_PE_Org_Value = ARM.STR_PE_Price_Org
        self.STR_CombPrice_Value = ARM.STR_CombPrice_Org

        self.Combine_Per_Value = ARM.Combine_Per
        self.Individual_Per_Value = ARM.Individual_Per
        self.Direction_Value = ARM.Direction_Signal_On

        self.STD_Strike_Change = False

        # Build UI (converted from init_ui)
        self.build_ui()

        # Bind click on container (safe outside-click handler)
        # This emulates PyQt mousePressEvent but does not steal entry events.
        self._container.bind("<Button-1>", self.mousePressEvent, add="+")

        # Start periodic update (every 1000 ms)
        self.update_current_prices()

        try:
            self.focus_force()      # set focus to dialog
        except Exception:
            pass

    # -------------------------
    # Price update loop (every 1 second)
    # -------------------------
    def update_current_prices(self):
        """Update labels from ARM every second (similar to QTimer)."""
        if not getattr(self, "alive", True):
            return

        try:
            if not ARM.STD_Strike_Change:
                # STD current values
                self.display_label3_std.config(text=f"{ARM.STD_CombinePrice_Cur:.2f}")
                self.display_label31_std.config(text=f"{ARM.STD_Combine_Per_Change:.2f} %")

                self.display_label4_std.config(text=f"{ARM.STD_CE_Price_Cur:.2f}")
                self.display_label41_std.config(text=f"{ARM.STD_CE_Per_Change:.2f} %")

                self.display_label5_std.config(text=f"{ARM.STD_PE_Price_Cur:.2f}")
                self.display_label51_std.config(text=f"{ARM.STD_PE_Per_Change:.2f} %")

                # STR current values
                self.display_label3_str.config(text=f"{ARM.STR_CombinePrice_Cur:.2f}")
                self.display_label31_str.config(text=f"{ARM.STR_Combine_Per_Change:.2f} %")

                self.display_label4_str.config(text=f"{ARM.STR_CE_Price_Cur:.2f}")
                self.display_label41_str.config(text=f"{ARM.STR_CE_Per_Change:.2f} %")

                self.display_label5_str.config(text=f"{ARM.STR_PE_Price_Cur:.2f}")
                self.display_label51_str.config(text=f"{ARM.STR_PE_Per_Change:.2f} %")

            # If strike change flag is set, refresh stored values into widgets
            if getattr(self, "STD_Strike_Change", True):
                self.STD_Strike_Change = False

                # copy updated ARM originals
                self.STD_CE_Org_Value = ARM.STD_CE_Price_Org
                self.STD_PE_Org_Value = ARM.STD_PE_Price_Org

                self.STR_CE_Strike_Value = ARM.STR_CE_Strike
                self.STR_PE_Strike_Value = ARM.STR_PE_Strike
                self.STR_CE_Org_Value = ARM.STR_CE_Price_Org
                self.STR_PE_Org_Value = ARM.STR_PE_Price_Org
                self.STR_CombPrice_Value = ARM.STR_CombPrice_Org

                # update disabled entries (temporarily enable to change)
                for entry, val in (
                    (self.std_ce_entry, str(self.STD_CE_Org_Value)),
                    (self.std_pe_entry, str(self.STD_PE_Org_Value)),
                    (self.str_ce_strike_entry, str(self.STR_CE_Strike_Value)),
                    (self.str_pe_strike_entry, str(self.STR_PE_Strike_Value)),
                    (self.str_ce_entry, str(self.STR_CE_Org_Value)),
                    (self.str_pe_entry, str(self.STR_PE_Org_Value)),
                ):
                    try:
                        entry.configure(state=NORMAL)
                        entry.delete(0, "end")
                        entry.insert(0, val)
                        entry.configure(state=DISABLED)
                    except Exception:
                        pass

                # update displays
                try:
                    self.display1_label1.config(text=f"{ARM.STD_Dis_MBEP:.2f}, {ARM.STD_Dis_STD:.2f}, {ARM.STD_Dis_PBEP:.2f}, {ARM.STD_Dis_Time}")
                    self.display_label2_std.config(text=f"{ARM.STD_CombPrice_Org:.2f}")
                    self.display_label2_str.config(text=f"{ARM.STR_CombPrice_Org:.2f}")
                except Exception:
                    pass
        except Exception:
            # ignore display errors if widgets not ready
            pass

        # -----------------------------------
        # STD UDS column update
        # -----------------------------------
        try:
            df = self.AVG_UDS_df

            if df is not None and not df.empty:
                # Take last 13 rows
                last_rows = df.tail(13)

                # Extract last column values
                uds_values = last_rows.iloc[:, -1].tolist()
            else:
                uds_values = []

            # Pad with zeros at the top if less than 13 values
            if len(uds_values) < 13:
                uds_values = [0] * (13 - len(uds_values)) + uds_values

            # Update labels
            for lbl, val in zip(self.std_uds_labels, uds_values):
                if val > ARM.UDS_DEADBAND:
                    lbl.config(text=f"{float(val):.2f}",bg="green",fg="white")
                elif val < -ARM.UDS_DEADBAND:
                    lbl.config(text=f"{float(val):.2f}",bg="red",fg="white")
                else:
                    lbl.config(text=f"{float(val):.2f}",bg="yellow",fg="black")
        except Exception:
            pass
        # schedule next call
        if getattr(self, "alive", True):
            try:
                self._after_id = self.after(1000, self.update_current_prices)
            except Exception:
                pass

    # -------------------------
    # Close handler
    # -------------------------
    def on_close(self):
        """Stop the update loop, clear ARM dialog reference, destroy window."""
        self.alive = False
        # cancel pending after
        try:
            if self._after_id is not None:
                self.after_cancel(self._after_id)
                self._after_id = None
        except Exception:
            pass

        try:
            ARM.STRADDLE_DIALOG = None
        except Exception:
            pass

        try:
            self.destroy()
        except Exception:
            pass

    # -------------------------
    # UI Builder (converted)
    # -------------------------
    def build_ui(self):
        container = tk.Frame(self, bg="#D6D6D2")
        container.grid(sticky="nsew", padx=2, pady=2)
        # keep reference so mousePressEvent binds to container
        self._container = container

        # Make columns similar to your previous layout
        for c in range(6):
            container.grid_columnconfigure(c, weight=0, minsize=95)

        # Row 0 headers ------------------------------------
        strike_label = tk.Label(container, text="Strategy Strike",
                                bg="#FF99FF", fg="black", font=("Arial", 10))
        strike_label.grid(row=0, column=0, columnspan=2, sticky="nsew", padx=2, pady=2)

        CE_label = tk.Label(container, text="CE LTP",
                            bg="#FF99FF", fg="black", font=("Arial", 10))
        CE_label.grid(row=0, column=2, columnspan=2, sticky="nsew", padx=2, pady=2)

        PE_label = tk.Label(container, text="PE LTP",
                            bg="#FF99FF", fg="black", font=("Arial", 10))
        PE_label.grid(row=0, column=4, columnspan=2, sticky="nsew", padx=2, pady=2)

        UDS_label = tk.Label(container, text="UDS",
                            bg="#FF99FF", fg="black", font=("Arial", 10))
        UDS_label.grid(row=0, column=6,sticky="nsew", padx=2, pady=2)

        # Row 1: straddle entries ------------------------------------
        # STD strike (editable) -- use FocusIn/Return/FocusOut bindings
        self.std_strike_entry = tk.Entry(container, fg="blue", justify=CENTER, font=("Arial", 10), width=8)
        self.std_strike_entry.insert(0, str(self.STD_Strike_Value))
        self.std_strike_entry.bind("<FocusIn>", lambda ev: self.on_entry_click(self.std_strike_entry, str(self.STD_Strike_Value)))
        self.std_strike_entry.bind("<FocusOut>", lambda ev: self.on_focus_out(self.std_strike_entry, str(self.STD_Strike_Value)))
        self.std_strike_entry.bind("<Return>", lambda ev: self.on_focus_out(self.std_strike_entry, str(self.STD_Strike_Value)))
        self.std_strike_entry.grid(row=1, column=0, columnspan=2, pady=2)

        # STD CE (disabled)
        self.std_ce_entry = tk.Entry(container, justify=CENTER, font=("Arial", 10), width=6)
        self.std_ce_entry.insert(0, str(self.STD_CE_Org_Value))
        self.std_ce_entry.grid(row=1, column=2, columnspan=2, pady=2)

        # STD PE (disabled)
        self.std_pe_entry = tk.Entry(container, justify=CENTER, font=("Arial", 10), width=6)
        self.std_pe_entry.insert(0, str(self.STD_PE_Org_Value))
        self.std_pe_entry.grid(row=1, column=4, columnspan=2, pady=2)

        self.std_ce_entry.configure(state=DISABLED)
        self.std_pe_entry.configure(state=DISABLED)
        # Row 2: strangle entries ------------------------------------
        self.str_ce_strike_entry = tk.Entry(container, justify=CENTER, font=("Arial", 10), width=8)
        self.str_ce_strike_entry.insert(0, str(self.STR_CE_Strike_Value))
        self.str_ce_strike_entry.grid(row=2, column=0, pady=2)

        self.str_pe_strike_entry = tk.Entry(container, justify=CENTER, font=("Arial", 10), width=8)
        self.str_pe_strike_entry.insert(0, str(self.STR_PE_Strike_Value))
        self.str_pe_strike_entry.grid(row=2, column=1, pady=2)

        self.str_ce_entry = tk.Entry(container, justify=CENTER, font=("Arial", 10), width=6)
        self.str_ce_entry.insert(0, str(self.STR_CE_Org_Value))
        self.str_ce_entry.grid(row=2, column=2, columnspan=2, pady=2)

        self.str_pe_entry = tk.Entry(container, justify=CENTER, font=("Arial", 10), width=6)
        self.str_pe_entry.insert(0, str(self.STR_PE_Org_Value))
        self.str_pe_entry.grid(row=2, column=4, columnspan=2, pady=2)

        self.str_ce_strike_entry.configure(state=DISABLED)
        self.str_pe_strike_entry.configure(state=DISABLED)
        self.str_ce_entry.configure(state=DISABLED)
        self.str_pe_entry.configure(state=DISABLED)
        # Row 3: combined / individual labels ------------------------------------
        com_label = tk.Label(container, text="Combined LTP BO (1 % to 25 %)",
                             bg="#FF99FF", fg="black", font=("Arial", 10))
        com_label.grid(row=3, column=0, columnspan=3, sticky="nsew", padx=2, pady=2)

        ind_label = tk.Label(container, text="Individual LTP BO (1 % to 50 %)",
                             bg="#FF99FF", fg="black", font=("Arial", 10))
        ind_label.grid(row=3, column=3, columnspan=3, sticky="nsew", padx=2, pady=2)

        # Row 4: entries (editable) ------------------------------------
        self.comb_entry = tk.Entry(container, fg="blue", justify=CENTER, font=("Arial", 10), width=6)
        self.comb_entry.insert(0, str(self.Combine_Per_Value))
        self.comb_entry.bind("<FocusIn>", lambda ev: self.on_entry_click(self.comb_entry, str(self.Combine_Per_Value)))
        self.comb_entry.bind("<FocusOut>", lambda ev: self.on_focus_out(self.comb_entry, str(self.Combine_Per_Value)))
        self.comb_entry.bind("<Return>", lambda ev: self.on_focus_out(self.comb_entry, str(self.Combine_Per_Value)))
        self.comb_entry.grid(row=4, column=0, columnspan=3, padx=2, pady=2)

        self.indi_entry = tk.Entry(container, fg="blue", justify=CENTER, font=("Arial", 10), width=6)
        self.indi_entry.insert(0, str(self.Individual_Per_Value))
        self.indi_entry.bind("<FocusIn>", lambda ev: self.on_entry_click(self.indi_entry, str(self.Individual_Per_Value)))
        self.indi_entry.bind("<FocusOut>", lambda ev: self.on_focus_out(self.indi_entry, str(self.Individual_Per_Value)))
        self.indi_entry.bind("<Return>", lambda ev: self.on_focus_out(self.indi_entry, str(self.Individual_Per_Value)))
        self.indi_entry.grid(row=4, column=3, columnspan=3, padx=2, pady=2)

        # Row 5: order label ------------------------------------
        order_label = tk.Label(container, text="Buy Signal on Break Out (BO)",
                               bg="#FF99FF", fg="black", font=("Arial", 10))
        order_label.grid(row=5, column=0, columnspan=6, sticky="nsew", padx=2, pady=2)

        # Row 6: radio buttons (use StringVar)
        self._direction_var = tk.StringVar(value=self.Direction_Value if getattr(self, "Direction_Value", None) else "IND")

        self.ind_radio = tk.Radiobutton(container, text="Individual", variable=self._direction_var, value="IND",
                                        command=lambda: self.on_radio_clicked("IND"), anchor="w")
        self.com_radio = tk.Radiobutton(container, text="Combined", variable=self._direction_var, value="COM",
                                        command=lambda: self.on_radio_clicked("COM"), anchor="w")
        self.hyb_radio = tk.Radiobutton(container, text="Hybrid", variable=self._direction_var, value="HYB",
                                        command=lambda: self.on_radio_clicked("HYB"), anchor="w")

        for rb in (self.ind_radio, self.com_radio, self.hyb_radio):
            rb.configure(fg="blue", bg="#D6D6D2", font=("Arial", 10))

        self.ind_radio.grid(row=6, column=0, columnspan=2, padx=2, pady=2)
        self.com_radio.grid(row=6, column=2, columnspan=2, padx=2, pady=2)
        self.hyb_radio.grid(row=6, column=4, columnspan=2, padx=2, pady=2)

        if getattr(self, "Direction_Value", None) == "HYB":
            self._direction_var.set("HYB")
        elif getattr(self, "Direction_Value", None) == "COM":
            self._direction_var.set("COM")
        else:
            self._direction_var.set("IND")

        self.ind_radio.configure(state=DISABLED)
        self.com_radio.configure(state=DISABLED)
        self.hyb_radio.configure(state=DISABLED)
        # Row 7..13: display labels and dynamic values ------------------------------------
        display_label = tk.Label(container, text="Current Settings Selected",
                                 bg="#FF99FF", fg="black", font=("Arial", 10))
        display_label.grid(row=7, column=0, columnspan=6, sticky="nsew", padx=2, pady=2)

        display_label1 = tk.Label(container, text="-STD_BEP, CMP, +STD_BEP, Time",
                                  bg="#DCDCDC", fg="black", font=("Arial", 10))
        display_label1.grid(row=8, column=0, columnspan=3, sticky="nsew", padx=2, pady=2)

        Straddle_label = tk.Label(container, text="Straddle", bg="#FF99FF", fg="black", font=("Arial", 10))
        Straddle_label.grid(row=9, column=2, columnspan=2, sticky="nsew", padx=2, pady=2)

        Strangle_label = tk.Label(container, text="Strangle", bg="#FF99FF", fg="black", font=("Arial", 10))
        Strangle_label.grid(row=9, column=4, columnspan=2, sticky="nsew", padx=2, pady=2)

        # Row 10..13 labels and dynamic widgets ------------------------------------
        display_label2 = tk.Label(container, text="Comb LTP (Org)", bg="#C0C0C0", fg="black", font=("Arial", 10))
        display_label2.grid(row=10, column=0, sticky="nsew", padx=2, pady=2)
        self.display_label2_std = tk.Label(container, text=f"{self.STD_CombPrice_Value:.2f}",
                                          bg="#C0C0C0", fg="black", font=("Arial", 10))
        self.display_label2_std.grid(row=10, column=2, columnspan=2, sticky="nsew", padx=2, pady=2)

        display_label3 = tk.Label(container, text="Comb LTP (Cur)", bg="#C0C0C0", fg="black", font=("Arial", 10))
        display_label3.grid(row=11, column=0, sticky="nsew", padx=2, pady=2)
        display_label31 = tk.Label(container, text="% Change", bg="#C0C0C0", fg="black", font=("Arial", 10))
        display_label31.grid(row=11, column=1, sticky="nsew", padx=2, pady=2)

        self.display_label3_std = tk.Label(container, text=f"{ARM.STD_CombinePrice_Cur:.2f}",
                                          bg="#C0C0C0", fg="black", font=("Arial", 10))
        self.display_label3_std.grid(row=11, column=2, sticky="nsew", padx=2, pady=2)
        self.display_label31_std = tk.Label(container, text=f"{ARM.STD_Combine_Per_Change:.2f} %",
                                           bg="#C0C0C0", fg="black", font=("Arial", 10))
        self.display_label31_std.grid(row=11, column=3, sticky="nsew", padx=2, pady=2)

        display_label4 = tk.Label(container, text="CE LTP (Cur)", bg="#DCDCDC", fg="black", font=("Arial", 10))
        display_label4.grid(row=12, column=0, sticky="nsew", padx=2, pady=2)
        display_label41 = tk.Label(container, text="% Change", bg="#DCDCDC", fg="black", font=("Arial", 10))
        display_label41.grid(row=12, column=1, sticky="nsew", padx=2, pady=2)
        self.display_label4_std = tk.Label(container, text=f"{ARM.STD_CE_Price_Cur:.2f}",
                                          bg="#DCDCDC", fg="black", font=("Arial", 10))
        self.display_label4_std.grid(row=12, column=2, sticky="nsew", padx=2, pady=2)
        self.display_label41_std = tk.Label(container, text=f"{ARM.STD_CE_Per_Change:.2f} %",
                                           bg="#DCDCDC", fg="black", font=("Arial", 10))
        self.display_label41_std.grid(row=12, column=3, sticky="nsew", padx=2, pady=2)

        display_label5 = tk.Label(container, text="PE LTP (Cur)", bg="#DCDCDC", fg="black", font=("Arial", 10))
        display_label5.grid(row=13, column=0, sticky="nsew", padx=2, pady=2)
        display_label51 = tk.Label(container, text="% Change", bg="#DCDCDC", fg="black", font=("Arial", 10))
        display_label51.grid(row=13, column=1, sticky="nsew", padx=2, pady=2)
        self.display_label5_std = tk.Label(container, text=f"{ARM.STD_PE_Price_Cur:.2f}",
                                          bg="#DCDCDC", fg="black", font=("Arial", 10))
        self.display_label5_std.grid(row=13, column=2, sticky="nsew", padx=2, pady=2)
        self.display_label51_std = tk.Label(container, text=f"{ARM.STD_PE_Per_Change:.2f} %",
                                           bg="#DCDCDC", fg="black", font=("Arial", 10))
        self.display_label51_std.grid(row=13, column=3, sticky="nsew", padx=2, pady=2)

        # aggregate display (row 8, columns 3..5)
        self.display1_label1 = tk.Label(container,
                                       text=f"{self.STD_Dis_MBEP:.2f}, {self.STD_Dis_STD_Value:.2f}, {self.STD_Dis_PBEP:.2f}, {self.STD_Dis_Time_Value}",
                                       bg="#DCDCDC", fg="black", font=("Arial", 10))
        self.display1_label1.grid(row=8, column=3, columnspan=3, sticky="nsew", padx=2)

        # Secondary STR block (right)
        self.display_label2_str = tk.Label(container, text=f"{self.STR_CombPrice_Value:.2f}",
                                          bg="#C0C0C0", fg="black", font=("Arial", 10))
        self.display_label2_str.grid(row=10, column=4, columnspan=2, sticky="nsew", padx=2, pady=2)

        self.display_label3_str = tk.Label(container, text=f"{ARM.STR_CombinePrice_Cur:.2f}",
                                          bg="#C0C0C0", fg="black", font=("Arial", 10))
        self.display_label3_str.grid(row=11, column=4, sticky="nsew", padx=2, pady=2)
        self.display_label31_str = tk.Label(container, text=f"{ARM.STR_Combine_Per_Change:.2f} %",
                                           bg="#C0C0C0", fg="black", font=("Arial", 10))
        self.display_label31_str.grid(row=11, column=5, sticky="nsew", padx=2, pady=2)

        self.display_label4_str = tk.Label(container, text=f"{ARM.STR_CE_Price_Cur:.2f}",
                                          bg="#DCDCDC", fg="black", font=("Arial", 10))
        self.display_label4_str.grid(row=12, column=4, sticky="nsew", padx=2, pady=2)
        self.display_label41_str = tk.Label(container, text=f"{ARM.STR_CE_Per_Change:.2f} %",
                                           bg="#DCDCDC", fg="black", font=("Arial", 10))
        self.display_label41_str.grid(row=12, column=5, sticky="nsew", padx=2, pady=2)

        self.display_label5_str = tk.Label(container, text=f"{ARM.STR_PE_Price_Cur:.2f}",
                                          bg="#DCDCDC", fg="black", font=("Arial", 10))
        self.display_label5_str.grid(row=13, column=4, sticky="nsew", padx=2, pady=2)
        self.display_label51_str = tk.Label(container, text=f"{ARM.STR_PE_Per_Change:.2f} %",
                                           bg="#DCDCDC", fg="black", font=("Arial", 10))
        self.display_label51_str.grid(row=13, column=5, sticky="nsew", padx=2, pady=2)

        # Separator row 14 ------------------------------------
        sep = tk.Frame(container, bg="#FF99FF", height=10)
        sep.grid(row=14, column=0, columnspan=7, sticky="ew", pady=(6, 6))

        # OK button row 15 ------------------------------------
        self.ok_button = tk.Button(container, text="Ok", bg="#FF99FF", fg="black", width=6,
                                   font=("Arial", 10, "bold"), command=self.on_close, takefocus=False)
        self.ok_button.grid(row=15, column=2, columnspan=2, sticky="", pady=6)

        # Ensure window fits content
        self.update_idletasks()
        self.minsize(self.winfo_width(), self.winfo_height())

        #--------------------------------------------
        self.std_uds_labels = []

        for i in range(13):
            lbl = tk.Label(
                container,
                text="0",
                bg="#D6D6D2",
                fg="black",
                font=("Arial", 10),
                # relief="ridge",
                width=8
            )
            lbl.grid(row=i + 1, column=6, sticky="nsew", padx=2, pady=1)
            self.std_uds_labels.append(lbl)
    # -------------------------
    # Radio handler
    # -------------------------
    def on_radio_clicked(self, button_value):
        """Accepts 'IND','COM','HYB' strings (wired from Radiobuttons)."""
        val = button_value
        if val == "IND" or val == "Individual":
            ARM.Direction_Signal_On = "IND"
        elif val == "COM" or val == "Combined":
            ARM.Direction_Signal_On = "COM"
        elif val == "HYB" or val == "Hybrid":
            ARM.Direction_Signal_On = "HYB"
        else:
            ARM.Direction_Signal_On = val

    # -------------------------
    # Validation methods (ported)
    # -------------------------
    def validate_std_strike(self):
        # keep a debug print you used earlier
        try:
            prev_value = self.STD_Strike_Value
            txt = self.std_strike_entry.get().strip()
            self.STD_Strike_Value = int(txt)
            if 10000 <= self.STD_Strike_Value <= 50000 and self.STD_Strike_Value % 50 == 0:
                try:
                    self.focus_set()
                except Exception:
                    pass

                if prev_value != self.STD_Strike_Value:
                    try:
                        ARM.STD_STR_Subscribe.drop(ARM.STD_STR_Subscribe.index, inplace=True)
                    except Exception:
                        pass

                    ARM.STD_Dis_STD = 0
                    ARM.STD_Dis_Time = ""
                    ARM.STD_Dis_PBEP = 0
                    ARM.STD_Dis_MBEP = 0

                    ARM.STD_Strike = self.STD_Strike_Value
                    ARM.STD_CE_Symbol = ""
                    ARM.STD_PE_Symbol = ""
                    ARM.STD_CE_Price_Org = 0
                    ARM.STD_PE_Price_Org = 0
                    ARM.STD_CE_Price_Cur = 0
                    ARM.STD_PE_Price_Cur = 0
                    ARM.STD_CombPrice_Org = 0
                    ARM.STD_CombinePrice_Cur = 0
                    ARM.STD_CE_Per_Change = 0
                    ARM.STD_PE_Per_Change = 0
                    ARM.STD_Combine_Per_Change = 0

                    ARM.STR_CE_Strike = 0
                    ARM.STR_PE_Strike = 0
                    ARM.STR_CE_Symbol = ""
                    ARM.STR_PE_Symbol = ""
                    ARM.STR_CE_Price_Org = 0
                    ARM.STR_PE_Price_Org = 0
                    ARM.STR_CE_Price_Cur = 0
                    ARM.STR_PE_Price_Cur = 0
                    ARM.STR_CombPrice_Org = 0
                    ARM.STR_CombinePrice_Cur = 0
                    ARM.STR_CE_Per_Change = 0
                    ARM.STR_PE_Per_Change = 0
                    ARM.STR_Combine_Per_Change = 0

                    ARM.STD_Strike_Change = True
                    self.STD_Strike_Change = True

                    # refresh local copies
                    self.STD_Dis_STD_Value = ARM.STD_Dis_STD
                    self.STD_Dis_Time_Value = ARM.STD_Dis_Time
                    self.STD_Dis_PBEP = ARM.STD_Dis_PBEP
                    self.STD_Dis_MBEP = ARM.STD_Dis_MBEP

                    self.STD_CE_Org_Value = ARM.STD_CE_Price_Org
                    self.STD_PE_Org_Value = ARM.STD_PE_Price_Org
                    self.STD_CombPrice_Value = ARM.STD_CombPrice_Org

                    self.STR_CE_Strike_Value = ARM.STR_CE_Strike
                    self.STR_PE_Strike_Value = ARM.STR_PE_Strike
                    self.STR_CE_Org_Value = ARM.STR_CE_Price_Org
                    self.STR_PE_Org_Value = ARM.STR_PE_Price_Org
                    self.STR_CombPrice_Value = ARM.STR_CombPrice_Org

                    # update UI entries (temporarily enable disabled ones)
                    try:
                        self.std_ce_entry.configure(state=NORMAL)
                        self.std_ce_entry.delete(0, "end")
                        self.std_ce_entry.insert(0, str(self.STD_CE_Org_Value))
                        self.std_ce_entry.configure(state=DISABLED)
                    except Exception:
                        pass

                    try:
                        self.std_pe_entry.configure(state=NORMAL)
                        self.std_pe_entry.delete(0, "end")
                        self.std_pe_entry.insert(0, str(self.STD_PE_Org_Value))
                        self.std_pe_entry.configure(state=DISABLED)
                    except Exception:
                        pass

                    try:
                        self.str_ce_strike_entry.configure(state=NORMAL)
                        self.str_ce_strike_entry.delete(0, "end")
                        self.str_ce_strike_entry.insert(0, str(self.STR_CE_Strike_Value))
                        self.str_ce_strike_entry.configure(state=DISABLED)
                    except Exception:
                        pass

                    try:
                        self.str_pe_strike_entry.configure(state=NORMAL)
                        self.str_pe_strike_entry.delete(0, "end")
                        self.str_pe_strike_entry.insert(0, str(self.STR_PE_Strike_Value))
                        self.str_pe_strike_entry.configure(state=DISABLED)
                    except Exception:
                        pass

                    try:
                        self.str_ce_entry.configure(state=NORMAL)
                        self.str_ce_entry.delete(0, "end")
                        self.str_ce_entry.insert(0, str(self.STR_CE_Org_Value))
                        self.str_ce_entry.configure(state=DISABLED)
                    except Exception:
                        pass

                    try:
                        self.str_pe_entry.configure(state=NORMAL)
                        self.str_pe_entry.delete(0, "end")
                        self.str_pe_entry.insert(0, str(self.STR_PE_Org_Value))
                        self.str_pe_entry.configure(state=DISABLED)
                    except Exception:
                        pass

                    # update labels
                    try:
                        self.display1_label1.config(text=f"{ARM.STD_Dis_MBEP:.2f}, {ARM.STD_Dis_STD:.2f}, {ARM.STD_Dis_PBEP:.2f}, {ARM.STD_Dis_Time}")
                        self.display_label2_std.config(text=f"{ARM.STD_CombPrice_Org:.2f}")
                        self.display_label2_str.config(text=f"{ARM.STR_CombPrice_Org:.2f}")

                        self.display_label3_std.config(text=f"{ARM.STD_CombinePrice_Cur:.2f}")
                        self.display_label31_std.config(text=f"{ARM.STD_Combine_Per_Change:.2f} %")

                        self.display_label4_std.config(text=f"{ARM.STD_CE_Price_Cur:.2f}")
                        self.display_label41_std.config(text=f"{ARM.STD_CE_Per_Change:.2f} %")

                        self.display_label5_std.config(text=f"{ARM.STD_PE_Price_Cur:.2f}")
                        self.display_label51_std.config(text=f"{ARM.STD_PE_Per_Change:.2f} %")

                        self.display_label3_str.config(text=f"{ARM.STR_CombinePrice_Cur:.2f}")
                        self.display_label31_str.config(text=f"{ARM.STR_Combine_Per_Change:.2f} %")

                        self.display_label4_str.config(text=f"{ARM.STR_CE_Price_Cur:.2f}")
                        self.display_label41_str.config(text=f"{ARM.STR_CE_Per_Change:.2f} %")

                        self.display_label5_str.config(text=f"{ARM.STR_PE_Price_Cur:.2f}")
                        self.display_label51_str.config(text=f"{ARM.STR_PE_Per_Change:.2f} %")
                    except Exception:
                        pass
                else:
                    ARM.STD_CE_Price_Org = self.STD_CE_Org_Value
                    ARM.STD_PE_Price_Org = self.STD_PE_Org_Value
                    ARM.Combine_Per = self.Combine_Per_Value

                # show strike entry with blue style (approx)
                try:
                    self.std_strike_entry.delete(0, "end")
                    self.std_strike_entry.insert(0, str(self.STD_Strike_Value))
                    self.std_strike_entry.configure(fg="blue")
                except Exception:
                    pass
            else:
                # invalid input: revert
                try:
                    self.std_strike_entry.delete(0, "end")
                    self.std_strike_entry.insert(0, str(prev_value))
                    self.std_strike_entry.configure(fg="blue")
                    self.STD_Strike_Value = prev_value
                except Exception:
                    pass
        except ValueError:
            # parse error -> revert
            try:
                self.reset_entry(self.std_strike_entry, self.STD_Strike_Value)
            except Exception:
                pass

    def validate_std_ce(self):
        try:
            prev_value = self.STD_CE_Org_Value
            txt = self.std_ce_entry.get().strip()
            self.STD_CE_Org_Value = round(float(txt), 2)
            if self.STD_CE_Org_Value >= 0.05 and self.STD_CE_Org_Value != ARM.STD_CE_Price_Org:
                # update display and ARM
                self.std_ce_entry.configure(state=NORMAL)
                self.std_ce_entry.delete(0, "end")
                self.std_ce_entry.insert(0, str(self.STD_CE_Org_Value))
                self.std_ce_entry.configure(state=DISABLED)

                ARM.STD_CE_Price_Org = self.STD_CE_Org_Value
                ARM.STD_CombPrice_Org = round((ARM.STD_CE_Price_Org + ARM.STD_PE_Price_Org), 2)

                ARM.STD_Dis_MBEP = ARM.STD_Strike - ARM.STD_CombPrice_Org
                ARM.STD_Dis_STD = ARM.Spot_Value
                ARM.STD_Dis_PBEP = ARM.STD_Strike + ARM.STD_CombPrice_Org
                ARM.STD_Dis_Time = time.strftime("%H:%M:%S")

                self.display1_label1.config(text=f"{ARM.STD_Dis_MBEP:.2f}, {ARM.STD_Dis_STD:.2f}, {ARM.STD_Dis_PBEP:.2f}, {ARM.STD_Dis_Time}")
                self.display_label2_std.config(text=f"{ARM.STD_CombPrice_Org:.2f}")
                try:
                    self.focus_set()
                except Exception:
                    pass
            else:
                self.STD_CE_Org_Value = prev_value
                self.reset_entry(self.std_ce_entry, prev_value)
        except Exception:
            self.reset_entry(self.std_ce_entry, getattr(self, "STD_CE_Org_Value", 0.0))

    def validate_std_pe(self):
        try:
            prev_value = self.STD_PE_Org_Value
            txt = self.std_pe_entry.get().strip()
            self.STD_PE_Org_Value = round(float(txt), 2)
            if self.STD_PE_Org_Value >= 0.05 and self.STD_PE_Org_Value != ARM.STD_PE_Price_Org:
                self.std_pe_entry.configure(state=NORMAL)
                self.std_pe_entry.delete(0, "end")
                self.std_pe_entry.insert(0, str(self.STD_PE_Org_Value))
                self.std_pe_entry.configure(state=DISABLED)

                ARM.STD_PE_Price_Org = self.STD_PE_Org_Value
                ARM.STD_CombPrice_Org = round((ARM.STD_CE_Price_Org + ARM.STD_PE_Price_Org), 2)

                ARM.STD_Dis_MBEP = ARM.STD_Strike - ARM.STD_CombPrice_Org
                ARM.STD_Dis_STD = ARM.Spot_Value
                ARM.STD_Dis_PBEP = ARM.STD_Strike + ARM.STD_CombPrice_Org
                ARM.STD_Dis_Time = time.strftime("%H:%M:%S")

                self.display1_label1.config(text=f"{ARM.STD_Dis_MBEP:.2f}, {ARM.STD_Dis_STD:.2f}, {ARM.STD_Dis_PBEP:.2f}, {ARM.STD_Dis_Time}")
                self.display_label2_std.config(text=f"{ARM.STD_CombPrice_Org:.2f}")
                try:
                    self.focus_set()
                except Exception:
                    pass
            else:
                self.STD_PE_Org_Value = prev_value
                self.reset_entry(self.std_pe_entry, prev_value)
        except Exception:
            self.reset_entry(self.std_pe_entry, getattr(self, "STD_PE_Org_Value", 0.0))

    def validate_combine_per(self):
        try:
            prev_value = self.Combine_Per_Value
            txt = self.comb_entry.get().strip()
            self.Combine_Per_Value = round(float(txt), 1)
            if 0 <= self.Combine_Per_Value <= 25.0 and self.Combine_Per_Value != ARM.Combine_Per:
                self.comb_entry.delete(0, "end")
                self.comb_entry.insert(0, str(self.Combine_Per_Value))
                try:
                    self.comb_entry.configure(fg="blue")
                except Exception:
                    pass
                ARM.Combine_Per = self.Combine_Per_Value
                ARM.STD_CombPrice_Org = round((ARM.STD_CE_Price_Org + ARM.STD_PE_Price_Org), 2)
                try:
                    self.focus_set()
                except Exception:
                    pass
            else:
                self.Combine_Per_Value = prev_value
                self.reset_entry(self.comb_entry, prev_value)
        except Exception:
            self.reset_entry(self.comb_entry, getattr(self, "Combine_Per_Value", 0.0))

    def validate_individual_per(self):
        try:
            prev_value = self.Individual_Per_Value
            txt = self.indi_entry.get().strip()
            self.Individual_Per_Value = round(float(txt), 1)
            if 1.0 <= self.Individual_Per_Value <= 50.0 and self.Individual_Per_Value != ARM.Individual_Per:
                self.indi_entry.delete(0, "end")
                self.indi_entry.insert(0, str(self.Individual_Per_Value))
                try:
                    self.indi_entry.configure(fg="blue")
                except Exception:
                    pass
                ARM.Individual_Per = self.Individual_Per_Value
                try:
                    self.focus_set()
                except Exception:
                    pass
            else:
                self.Individual_Per_Value = prev_value
                self.reset_entry(self.indi_entry, prev_value)
        except Exception:
            self.reset_entry(self.indi_entry, getattr(self, "Individual_Per_Value", 0.0))

    # -------------------------
    # Entry styling helpers
    # -------------------------
    def reset_entry(self, entry, value):
        try:
            entry.delete(0, "end")
            entry.insert(0, str(value))
            try:
                entry.configure(fg="blue")
            except Exception:
                pass
        except Exception:
            pass

    def on_entry_click(self, entry, default_value):
        try:
            entry.focus_set()  # MUST be first so cursor becomes visible
        except Exception:
            pass

        try:
            # if entry.get().strip() == default_value:
                # entry.delete(0, "end")
            entry.delete(0, "end")
            entry.insert(0, str(default_value))
        except Exception:
            pass

        try:
            entry.configure(fg="gray")
        except Exception:
            pass

    def on_focus_out(self, entry, default_value):
        try:
            value = entry.get().strip()

            # If empty → restore default, no validation
            if not value:
                entry.delete(0, "end")
                entry.insert(0, str(default_value))
                entry.configure(fg="blue")
                return

            # If unchanged → no validation
            if str(value) == str(default_value):
                entry.configure(fg="blue")
                return

            # Changed value → validate based on entry
            if entry == self.std_strike_entry:
                self.validate_std_strike()
            elif entry == self.comb_entry:
                self.validate_combine_per()
            elif entry == self.indi_entry:
                self.validate_individual_per()

            entry.configure(fg="blue")

        except Exception:
            pass

    def eventFilter(self, obj, event):      # Event filter compatibility stub (not used in Tkinter)
        # Tkinter binds events on widgets directly; keep stub for parity.
        return False

    # -------------------------
    # Mouse press handler (container click)
    # -------------------------
    def mousePressEvent(self, event):
        """
        Bound to <Button-1> on the container frame. If user clicks outside editable
        entries or clicks the OK button, reset entries and clear focus — mimics PyQt.
        """
        try:
            # event.widget is the exact widget that received the click
            clicked = event.widget

            # If click was on OK button, ignore (we want its command)
            if clicked is self.ok_button:
                return

            # If click was inside any editable entry, do nothing (entry already handles focus/validation)
            editable_widgets = (self.std_strike_entry, self.comb_entry, self.indi_entry)
            if clicked in editable_widgets:
                return

            # If clicked inside any disabled entry or label etc, we still want to reset editable entries
            try:
                self.reset_entry(self.std_strike_entry, getattr(self, "STD_Strike_Value", 0))
                self.reset_entry(self.comb_entry, getattr(self, "Combine_Per_Value", 0.0))
                self.reset_entry(self.indi_entry, getattr(self, "Individual_Per_Value", 0.0))
            except Exception:
                pass

            # clear focus (return focus to Toplevel)
            try:
                self.focus_set()
            except Exception:
                pass

        except Exception:
            pass

#--------------------------------
def convert_txt_to_pdf(input_file, output_file):
    #--Read the text content from the input file
    with open(input_file, 'r') as f:
        text_content = f.readlines()

    #--Create a PDF document
    doc = SimpleDocTemplate(output_file, pagesize=letter)
    styles = getSampleStyleSheet()

    #--Add text content to the PDF
    flowables = []
    for line in text_content:
        para = Paragraph(line.strip(), style=styles["Normal"])
        flowables.append(para)

    #--Build the PDF document
    doc.build(flowables)

#================= Function Those Must Respond=======================================================
def APH_Delay(x):
    x=x-1
    if (x<0):
        return
    else:
        root.after(1000,APH_Delay(x))

def Get_OB_LB():
    try:
        my_balance = kite.margins()
        GUI.R1C3_Value = round((my_balance["equity"]["available"]["opening_balance"]+my_balance["equity"]["available"]["intraday_payin"]), ndigits = 2)
        GUI.R1C4_Value = round((my_balance["equity"]["available"]["live_balance"]), ndigits = 2)
        R1C3.set("{:.2f}".format(GUI.R1C3_Value))
        R1C4.set("{:.2f}".format(GUI.R1C4_Value))
    except Exception as e:
        ARM.NO_Response -=1
        if (ARM.NO_Response < 0):
            print("Zerodha/Internet Not Responding (103)")
            ARM.NO_Response = 2
        else:pass
        R2C8.set("OP/LIV Balance not Updating")
        R2C8entry.config(fg="white",bg="red")
        if ARM.ErrorExcptionInfo_scheduled is None: ARM.ErrorExcptionInfo_scheduled = root.after(2000,Clear_ErrorExcptionInfo)
        APH_Delay(5)
        Get_OB_LB()

def get_hist_oi_opt(symbol,v):
    try:
        ARM.prev_day_oi_opt[symbol]
    except:
        try:
            pre_day_data_opt = kite.historical_data(v["token"], (datetime.now() - timedelta(days=5)).date(),
                                    (datetime.now() - timedelta(days=1)).date(), "day", oi=True)
            try:
                ARM.prev_day_oi_opt[symbol] = pre_day_data_opt[-1]["oi"]
            except:
                ARM.prev_day_oi_opt[symbol] = 0
        except Exception as e:
            root.after(500,get_hist_oi_opt(symbol,v))
    return
def get_oi_opt(data):
    ARM.CIO_Updated = False
    for symbol, v in data.items():
        if ARM.stop_thread_opt:
            break
        else:
            get_hist_oi_opt(symbol,v)
    ARM.CIO_Updated = True
    #print(ARM.prev_day_oi_opt)
    if ARM.prev_day_oi_opt:
        ARM.prev_day_oi_opt_df = pd.DataFrame(list(ARM.prev_day_oi_opt.items()), columns=['TradeSymbol', 'Prv_OI'])
        ARM.prev_day_oi_opt_df['TradeSymbol'] = ARM.prev_day_oi_opt_df['TradeSymbol'].str.replace('NFO:', '', regex=False)
        #ARM.prev_day_oi_opt_df.to_csv("prev_day_oi_opt_df.csv", index=False)
        #print("Updated previous day OPT-OI Data")
    else:pass
    return
def Disable_Root_Window():
    root.attributes("-disabled", True)
    root.grab_set()
def Enable_Root_Window():
    root.attributes("-disabled", False)
    root.grab_release()

def Get_Exchange_Data(kite,ns):
    Got_Exchange_Data = False
    while not Got_Exchange_Data:
        try:
            #-----PART-1 Equity DATA ------------------
            instrument_nse = pd.DataFrame(kite.instruments("NSE"))
            instrument_nse = instrument_nse[~instrument_nse['tradingsymbol'].str.match(r'^\d')]
            instrument_nse = instrument_nse[~instrument_nse['name'].str.match(r'^\d')]
            instrument_nse.reset_index(drop=True, inplace=True)

            #-----PART-2 FUTURE & OPTION DATA ------------------
            instrument_nfo = pd.DataFrame(kite.instruments("NFO"))

            ns.df1 = instrument_nse
            ns.df2 = instrument_nfo

            Got_Exchange_Data = True
            #print("Downloaded Exchange Data...")
        except:
            print("Downloading Exchange Data, Please wait . . . .")
            #print("Exchange Download Error (E)...")
            time.sleep(10)

def POPUP_YESNO(title, message):
    yesno_window = tk.Toplevel(root)
    yesno_window.title(title)

    x = root.winfo_x()
    y = root.winfo_y()
    yesno_window.geometry(f"350x193+{x+350}+{y+250}")    #--(W x H + W + H)  ("230x225+325+295")
    yesno_window.wm_attributes("-topmost", 1)           #--Set the window to stay on top
    yesno_window.resizable(False,False)
    #yesno_window.iconbitmap(cwd + '\\Images\\Question.ico')
    yesno_window.iconbitmap(os.path.join(cwd, 'Images', 'Question.ico'))

    yesno_window.grid_columnconfigure(0, weight=1) #--Columns are seperated by equal distance
    Label(yesno_window, text="",font=('Arial', 4),fg="black",bg='#ffe5ff').grid(row=0,column=0,sticky='nsew')
    Label(yesno_window, text="",font=('Arial', 8),fg="black",bg='#ffe5ff').grid(row=1,column=0,sticky='nsew')
    Label(yesno_window, text=message,fg="black",bg='#ffe5ff',font='Helvetica 10 bold').grid(row=2,column=0,sticky='nsew')
    Label(yesno_window, text="",font=('Arial', 8),fg="black",bg='#ffe5ff').grid(row=3,column=0,sticky='nsew')
    Label(yesno_window, text="",font=('Arial', 3),fg="black",bg='#e5e5ff').grid(row=4,column=0,sticky='nsew')
    Label(yesno_window, text="",font=('Arial', 2),fg="black",bg='#e5e5ff').grid(row=5,column=0,sticky='nsew')
    Disable_Root_Window()

    user_choice = None

    def OK_selection():
        nonlocal user_choice
        user_choice = True
        yesno_window.destroy()
        Enable_Root_Window()
        root.deiconify()  #--Ensure that the main window is not minimized
        root.focus_set()

    def No_selection():
        nonlocal user_choice
        user_choice = False
        yesno_window.destroy()
        Enable_Root_Window()
        root.deiconify()  #--Ensure that the main window is not minimized
        root.focus_set()

    def close_window():
        pass
    Label(yesno_window, text="",bg='#e5e5ff').grid(row=6,column=0,sticky='nsew')
    button_frame = tk.Frame(yesno_window,bg='#e5e5ff')
    button_frame.grid(row=6, column=0)

    Yes_button = tk.Button(button_frame, image=yes_button_image, bg='#e5e5ff', command=OK_selection, borderwidth=0)
    Yes_button.pack(side=tk.LEFT)
    Yes_button.bind('<Return>', lambda event: OK_selection())       #--to Bind it to the function that converts it to spacebar to enter

    No_button = tk.Button(button_frame, image=no_button_image, bg='#e5e5ff', command=No_selection, borderwidth=0)
    No_button.pack(side=tk.RIGHT)
    No_button.bind('<Return>', lambda event: No_selection())         #--to Bind it to the function that converts it to spacebar to enter

    Label(yesno_window, text="",font=('Arial', 2),bg='#e5e5ff').grid(row=7,column=0,sticky='nsew')
    Label(yesno_window, text="",font=('Arial', 3),bg='#e5e5ff').grid(row=8,column=0,sticky='nsew')

    Yes_button.focus_set()

    yesno_window.protocol("WM_DELETE_WINDOW", close_window)

    yesno_window.wait_window()  #--Wait for the yesno window to be closed
    return user_choice
def POPUP_WARNING(message):
    warning_window = tk.Toplevel(root)
    warning_window.title("Warning")

    x = root.winfo_x()
    y = root.winfo_y()
    x=x+735
    y=y+280
    warning_window.geometry(f"350x240+{x}+{y}")     #--(W x H + W + H)  ("230x225+325+295")
    warning_window.wm_attributes("-topmost", 1)     #--Set the window to stay on top
    warning_window.resizable(False,False)
    #warning_window.iconbitmap(cwd + '\\Images\\Warning.ico')
    warning_window.iconbitmap(os.path.join(cwd, 'Images', 'Warning.ico'))

    warning_window.grid_columnconfigure(0, weight=1) #--Columns are seperated by equal distance
    Label(warning_window, text="",font=('Arial', 3),fg="black",bg='#ffe5ff').grid(row=0,column=0,sticky='nsew')
    Label(warning_window, text="",font=('Arial', 8),fg="black",bg='#ffe5ff').grid(row=1,column=0,sticky='nsew')
    Label(warning_window, text=message,fg="black",bg='#ffe5ff',font='Helvetica 10 bold').grid(row=2,column=0,sticky='nsew')
    Label(warning_window, text="",font=('Arial', 8),fg="black",bg='#ffe5ff').grid(row=3,column=0,sticky='nsew')
    Label(warning_window, text="",font=('Arial', 3),fg="black",bg='#e5e5ff').grid(row=4,column=0,sticky='nsew')
    Label(warning_window, text="",font=('Arial', 2),fg="black",bg='#e5e5ff').grid(row=5,column=0,sticky='nsew')
    Disable_Root_Window()

    user_choice = None

    def OK_selection():
        nonlocal user_choice
        user_choice = True
        warning_window.destroy()
        Enable_Root_Window()
        root.deiconify()  #--Ensure that the main window is not minimized
        root.focus_set()

    Label(warning_window, text="",bg='#e5e5ff').grid(row=6,column=0,sticky='nsew')
    button_frame = tk.Frame(warning_window,bg='#e5e5ff')
    button_frame.grid(row=6, column=0)

    Ok_button = tk.Button(button_frame, image=ok_button_image, bg='#e5e5ff', command=OK_selection, borderwidth=0)
    Ok_button.pack(side=tk.LEFT)
    Ok_button.bind('<Return>', lambda event: OK_selection())       #--to Bind it to the function that converts it to spacebar to enter

    Label(warning_window, text="",font=('Arial', 2),bg='#e5e5ff').grid(row=7,column=0,sticky='nsew')
    Label(warning_window, text="",font=('Arial', 3),bg='#e5e5ff').grid(row=8,column=0,sticky='nsew')

    Ok_button.focus_set()

    warning_window.protocol("WM_DELETE_WINDOW", OK_selection)

    warning_window.wait_window()  #--Wait for the warning window to be closed
def POPUP_ETOKEN(userid):
    etoken_window = tk.Toplevel(root)
    etoken_window.title("Provide Etoken for : {}".format(userid))

    x = root.winfo_x()
    y = root.winfo_y()
    etoken_window.geometry(f"850x162+{x}+{y}")  #--(W x H + W + H)  ("230x225+325+295")
    etoken_window.wm_attributes("-topmost", 1)  #--Set the window to stay on top
    etoken_window.resizable(False,False)
    #etoken_window.iconbitmap(cwd + '\\Images\\APH_Icon.ico')
    etoken_window.iconbitmap(os.path.join(cwd, 'Images', 'APH_Icon.ico'))

    etoken_window.grid_columnconfigure(0, weight=1) #--Columns are seperated by equal distance

    def Disable_ETOKEN_Close():
        pass

    Label(etoken_window, text="",font=('Arial', 3),fg="black",bg='#ffe5ff').grid(row=0,column=0,sticky='nsew')
    Label(etoken_window, text="",font=('Arial', 8),fg="black",bg='#ffe5ff').grid(row=1,column=0,sticky='nsew')
    Label(etoken_window, text="Provide Etoken",fg="black",bg='#ffe5ff',font='Helvetica 10 bold').grid(row=2,column=0,sticky='nsew')

    etoken_input = tk.StringVar()

    tk.Entry(etoken_window, text=etoken_input,width=135).grid(row=3,column=0)

    Label(etoken_window, text="",font=('Arial', 8),fg="black",bg='#ffe5ff').grid(row=4,column=0,sticky='nsew')
    Label(etoken_window, text="",font=('Arial', 3),fg="black",bg='#e5e5ff').grid(row=5,column=0,sticky='nsew')
    Label(etoken_window, text="",font=('Arial', 2),fg="black",bg='#e5e5ff').grid(row=6,column=0,sticky='nsew')

    Disable_Root_Window()

    def OK_selection():
        etoken_window.destroy()
        Enable_Root_Window()
        root.deiconify()  #--Ensure that the main window is not minimized
        root.focus_set()

    Label(etoken_window, text="",bg='#e5e5ff').grid(row=7,column=0,sticky='nsew')
    button_frame = tk.Frame(etoken_window,bg='#e5e5ff')
    button_frame.grid(row=7, column=0)

    ok_button = tk.Button(button_frame, image=ok_button_image, bg='#e5e5ff', command=OK_selection, borderwidth=0)
    ok_button.pack(side=tk.LEFT)
    ok_button.bind('<Return>', lambda event: OK_selection())       #--to Bind it to the function that converts it to spacebar to enter

    Label(etoken_window, text="",font=('Arial', 2),bg='#e5e5ff').grid(row=8,column=0,sticky='nsew')
    Label(etoken_window, text="",font=('Arial', 3),bg='#e5e5ff').grid(row=9,column=0,sticky='nsew')
    ok_button.focus_set()
    etoken_window.protocol("WM_DELETE_WINDOW", Disable_ETOKEN_Close)
    etoken_window.wait_window()  #--Wait for the etoken_window to be closed
    return etoken_input.get()
def POPUP_LOGOUT():
    Logout_Window = tk.Toplevel(root)
    Logout_Window.title("Logout ?")
    x = root.winfo_x()
    y = root.winfo_y()
    Logout_Window.geometry(f"230x183+{x+1245}+{y}")     #--(W x H + W + H)  ("230x225+325+295")
    Logout_Window.wm_attributes("-topmost", 1)          #--Set the window to stay on top
    Logout_Window.resizable(False,False)
    #Logout_Window.iconbitmap(cwd + '\\Images\\Logout.ico')
    Logout_Window.iconbitmap(os.path.join(cwd, 'Images', 'Logout.ico'))

    Logout_Window.grid_columnconfigure(0, weight=1)     #--Columns are seperated by equal distance

    Disable_Root_Window()

    Label(Logout_Window, text="",font=('Arial', 4),fg="black",bg='#ffe5ff').grid(row=0,column=0,sticky='nsew')
    Label(Logout_Window, text="",font=('Arial', 8),fg="black",bg='#ffe5ff').grid(row=1,column=0,sticky='nsew')

    if (ARM.Kws_ReconnAttempt < ARM.Kws_MaxReconAttempt):
        Label(Logout_Window, text="Please, verify",fg="black",bg='#ffe5ff',font='Helvetica 10 bold').grid(row=2,column=0,sticky='nsew')
        Label(Logout_Window, text="All positions are cleared,",fg="black",bg='#ffe5ff',font='Helvetica 10 bold').grid(row=3,column=0,sticky='nsew')
        Label(Logout_Window, text="in your DMAT Account.",fg="black",bg='#ffe5ff',font='Helvetica 10 bold').grid(row=4,column=0,sticky='nsew')
    else:
        Label(Logout_Window, text="Zerodha Server Not Responding.",fg="black",bg='#ffe5ff',font='Helvetica 10 bold').grid(row=2,column=0,sticky='nsew')
        Label(Logout_Window, text="Maximum Reconnection Attempts Reached.",fg="black",bg='#ffe5ff',font='Helvetica 10 bold').grid(row=3,column=0,sticky='nsew')
        Label(Logout_Window, text="Want to Logg Out ?",fg="black",bg='#ffe5ff',font='Helvetica 10 bold').grid(row=4,column=0,sticky='nsew')

    Label(Logout_Window, text="",font=('Arial', 3),fg="black",bg='#ffe5ff').grid(row=5,column=0,sticky='nsew')

    Label(Logout_Window, text="",font=('Arial', 5),fg="black",bg='#ffe5ff').grid(row=6,column=0,sticky='nsew')

    var1 = tk.IntVar()
    var1.set(1)
    if (GUI.Trade_Flag == "DEMO_ONLY"):
        tk.Checkbutton(Logout_Window, text="Save Opening & Live Balance",font='Helvetica 10 bold',fg="blue",bg='#ffe5ff',variable=var1, onvalue=1, offvalue=0,).grid(row=6,column=0)
    elif (GUI.Trade_Flag == "DEMO_REAL_BOTH"):
        Label(Logout_Window, text="",font='Helvetica 10 bold',fg="blue",bg='#ffe5ff').grid(row=6,column=0,sticky='nsew')
    else:pass

    Label(Logout_Window, text="",font=('Arial', 2),fg="black",bg='#e5e5ff').grid(row=7,column=0,sticky='nsew')

    def logout_selection():
        if (GUI.Trade_Flag == "DEMO_ONLY") and (var1.get() == 1):
            if os.path.exists(TxtSettingPath):
                with open(TxtSettingPath, 'w') as file:
                    file.write(f"{ARM.Opening_Balance}\n")
                    file.write(f"{ARM.Live_Balance}\n")
                    file.close()
                print("Saved Opening & Live Balance, for next trading session.")
            else:
                print("Could not Save Opening & Live Balance, setting file not found")
            #print("{0} Opening Balance: {1}, Live Balance: {2}".format(ARM.your_user_id,ARM.Opening_Balance,ARM.Live_Balance))
            print("-------------------------------------------------------------------------------------") #--85 White Spaces
            print("Trading Session End                          {}".format(datetime.now().strftime("%H:%M:%S, %d %b %Y, %a")))
            print("-------------------------------------------------------------------------------------")
        else:
            #print("{0} Opening Balance: {1}, Live Balance: {2}".format(ARM.your_user_id,ARM.Opening_Balance,ARM.Live_Balance))
            print("-------------------------------------------------------------------------------------") #--85 White Spaces
            print("Trading Session End                          {}".format(datetime.now().strftime("%H:%M:%S, %d %b %Y, %a")))
            print("-------------------------------------------------------------------------------------")

        Logout_Window.destroy()
        if(len(ARM.My_AllTrades)>0):
            filename = f"My_AllTrades_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            filepath = os.path.join(os.getcwd(),'Log', filename)
            ARM.My_AllTrades.to_csv(filepath, index=False)
        else:pass
        root.destroy()
        sys.stdout.close()
        if (os.path.exists(TxtLogFilePath)):
            convert_txt_to_pdf(TxtLogFilePath, PdfLogFilePath)
            os.remove(TxtLogFilePath)
        else:pass
        sys.exit()

    def cancel_selection():
        ARM.Kws_ReconnAttempt = ARM.Kws_DelayCounter = 0                #--Manually Resetted to try again
        Logout_Window.destroy()
        if(R0C0.get() == "Login"): R0C0entry.configure(state="normal")  #--Re-enable the OptionMenu
        Enable_Root_Window()
        root.deiconify()  #--Ensure that the main window is not minimized
        root.focus_set()
        return

    Label(Logout_Window, text="",bg='#e5e5ff').grid(row=8,column=0,sticky='nsew')
    button_frame = tk.Frame(Logout_Window,bg='#e5e5ff')
    button_frame.grid(row=8, column=0)
    #-----------------------------------------------------
    Logout_button = tk.Button(button_frame, image=logout_button_image, bg='#e5e5ff', command=logout_selection, borderwidth=0)
    Logout_button.pack(side=tk.LEFT)
    Logout_button.bind('<Return>', lambda event: logout_selection())

    cancel_button = tk.Button(button_frame,image=cancel_button_image,bg='#e5e5ff',command=cancel_selection,borderwidth=0)
    cancel_button.pack(side=tk.RIGHT)
    cancel_button.bind('<Return>', lambda event: cancel_selection())

    Label(Logout_Window, text="",font=('Arial', 2),bg='#e5e5ff').grid(row=9,column=0,sticky='nsew')
    Label(Logout_Window, text="",font=('Arial', 3),bg='#e5e5ff').grid(row=10,column=0,sticky='nsew')

    Logout_Window.protocol("WM_DELETE_WINDOW", cancel_selection)

    Logout_Window.wait_window()  #--Wait for the popup window to be closed
def POPUP_LOGIN():
    login_window = tk.Toplevel(root)
    login_window.title("Logging...")
    x = root.winfo_x()
    y = root.winfo_y()
    #print("x= {}, y= {}".format(x,y))
    login_window.geometry(f"420x250+{x}+{y}")
    login_window.wm_attributes("-topmost", 1)  #--Set the window to stay on top
    login_window.resizable(False,False)
    #login_window.iconbitmap(cwd + '\\Images\\Login.ico')
    login_window.iconbitmap(os.path.join(cwd, 'Images', 'Login.ico'))

    login_window.grid_columnconfigure((0,1,2,3), weight=1) #--Columns are seperated by equal distance

    def Disable_Login_Close():
        pass
    global UID_Value
    UID_Value = None

    Label(login_window, text="Login Window",font='Helvetica 12 bold',fg="black",bg='#e5e5ff').grid(row=0,column=0,columnspan=4,sticky='nsew')

    Label(login_window, text=""

        ,fg="black",bg='#e5e5ff',justify="left",wraplength=450).grid(row=1,column=0,columnspan=4,sticky='nsew')
    #-------------------------------------
    radio_vars = []
    Label(login_window, text="",font='Helvetica 10 bold',fg="black",bg='#ffe5ff').grid(row=3,column=0, columnspan=4,sticky='nsew')
    previous_trade_selection = tk.StringVar()  #--OT Order Type (DEMO / BOTH)
    Label(login_window, text="Trade",font='Helvetica 10 bold',fg="black",bg='#ffe5ff').grid(row=3,column=0,columnspan=4,sticky='nsew')
    Label(login_window, text="",bg='#ffe5ff').grid(row=4,column=0,columnspan=4,sticky='nsew')
    Label(login_window, text="",bg='#ffe5ff').grid(row=5,column=0,columnspan=4,sticky='nsew')
    def on_radiobutton_clickMN(*args):
        if (ARM.NIFTY_LEG):  #--Replace with your actual condition
            radio_vars[0].set(previous_trade_selection.get())
            DEMO_radio_button['state'] = tk.DISABLED
            BOTH_radio_button['state'] = tk.DISABLED
        else:pass

    radio_vars.append(tk.StringVar(value=f"{GUI.Trade_Flag}"))
    DEMO_radio_button = tk.Radiobutton(login_window, text="DEMO", variable=radio_vars[0], value="DEMO_ONLY",fg="black",bg='#ffe5ff')
    DEMO_radio_button.grid(row=4,column=0,columnspan=4, padx=10, pady=5)
    BOTH_radio_button = tk.Radiobutton(login_window, text="DEMO & REAL", variable=radio_vars[0], value="DEMO_REAL_BOTH",fg="black",bg='#ffe5ff')
    BOTH_radio_button.grid(row=5,column=0,columnspan=4, padx=10, pady=5)

    previous_trade_selection.set(GUI.Trade_Flag)
    radio_vars[0].trace("w", lambda *args: on_radiobutton_clickMN())

    Label(login_window, text="",font=('Arial', 2),bg='#e5e5ff').grid(row=6,column=0,columnspan=4,sticky='nsew')
    Label(login_window, text="",font=('Arial', 2),bg='#e5e5ff').grid(row=7,column=0,columnspan=4,sticky='nsew')

    def save_selection():
        global UID_Value
        if ARM.Exchange_Data_Process is not None:
            if not ARM.Exchange_Data_Process.is_alive():
                ARM.Exchange_Data_Process.join()
                ARM.instrument_nse = ns.df1
                ARM.instrument_nfo = ns.df2
                Exchange_Status["text"]="Downloaded Exchange Data"

                if UID_Value is None:
                    #print("UID_Value",UID_Value)
                    #print("your_user_id",ARM.your_user_id)
                    UID_Entered()
                else:pass
                ARM.Login_selected_flag = False
                for i, var in enumerate(radio_vars):
                    if i==0:GUI.Trade_Flag = var.get()
                ARM.License_Type = "P" if (GUI.Trade_Flag == 'DEMO_ONLY') else "R"
                login_window.destroy()
                print("Downloaded Exchange Data")
            else:
                Exchange_Status["text"]="Downloading Exchange Data, Please wait . . . ."
        else:
            print("Exchange Data Downloading not Started, Restart ARM")

    def cancel_selection():
        if ARM.Exchange_Data_Process.is_alive():
            ARM.Exchange_Data_Process.terminate()  #--Terminate the child process
            ARM.Exchange_Data_Process.join()       #--Ensure the process has terminated
        else:pass
        login_window.destroy()
        root.destroy()

    Label(login_window, text="",bg='#e5e5ff').grid(row=8,column=0,columnspan=4,sticky='nsew')
    userid_frame = tk.Frame(login_window,bg='#e5e5ff')
    userid_frame.grid(row=8, column=0,columnspan=4)

    tk.Label(userid_frame,text="User-ID",font='Helvetica 10 bold',fg="black",bg='#e5e5ff',borderwidth=0).pack(side=tk.LEFT,padx=10)

    def UID_Entered():
        global UID_Value
        try:
            UID_Value = UIDentry.get()
            if (UID_Value is None):
                UIDentry.delete(0, tk.END)
                UIDentry.insert(0, ARM.your_user_id)
                UIDentry.config(fg='grey')
            else:
                UIDentry.delete(0, tk.END)
                ARM.your_user_id = (UID_Value.upper())
                UIDentry.insert(0, ARM.your_user_id)
        except:
            UID_Value = None
            UIDentry.delete(0, tk.END)
            UIDentry.insert(0, ARM.your_user_id)
        Save_button.focus_set()

    def on_entry_click(event):
        if UIDentry.get() == ARM.your_user_id:
            UIDentry.delete(0, tk.END)
            UIDentry.config(fg='black')  #--Change text color to black

    def on_focus_out(event):
        if not UIDentry.get():
            UIDentry.insert(0, ARM.your_user_id)
            UIDentry.config(fg='grey')  #--Change text color to grey

    UIDentry = tk.Entry(userid_frame,font='Helvetica 10',fg='grey',width=10)
    UIDentry.insert(0, ARM.your_user_id)
    UIDentry.bind("<FocusIn>", on_entry_click)
    UIDentry.bind("<FocusOut>", on_focus_out)
    UIDentry.bind('<Return>', lambda event: UID_Entered())
    UIDentry.pack(side=tk.RIGHT,padx=10)

    Label(login_window, text="",font=('Arial', 2),bg='#e5e5ff').grid(row=9,column=0,columnspan=4,sticky='nsew')

    Label(login_window, text="",bg='#e5e5ff').grid(row=10,column=0,columnspan=4,sticky='nsew')
    button_frame = tk.Frame(login_window,bg='#e5e5ff')
    button_frame.grid(row=10, column=0,columnspan=4)

    Save_button = tk.Button(button_frame,image=save_button_image,bg='#e5e5ff',command=save_selection,borderwidth=0)
    Save_button.pack(side=tk.LEFT)
    Save_button.bind('<Return>', lambda event: save_selection())

    cancel_button = tk.Button(button_frame,image=cancel_button_image,bg='#e5e5ff',command=cancel_selection,borderwidth=0)
    cancel_button.pack(side=tk.RIGHT)
    cancel_button.bind('<Return>', lambda event: cancel_selection())

    #Label(login_window, text="",font=('Arial', 2),bg='#e5e5ff').grid(row=11,column=0,columnspan=4,sticky='nsew')
    Exchange_Status = Label(login_window, text="",fg="black",font='Helvetica 8 bold',bg='#e5e5ff')
    Exchange_Status.grid(row=11,column=0,columnspan=4,sticky='nsew')
    Label(login_window, text="",font=('Arial', 2),bg='#e5e5ff').grid(row=12,column=0,columnspan=4,sticky='nsew')

    Save_button.focus_set()
    login_window.protocol("WM_DELETE_WINDOW", Disable_Login_Close)

def POPUP_SETTING():
    setting_window = tk.Toplevel(root)
    setting_window.title("Setting...")
    x = root.winfo_x()
    y = root.winfo_y()
    setting_window.geometry(f"420x230+{x}+{y}")
    setting_window.wm_attributes("-topmost", 1)
    setting_window.resizable(False,False)
    setting_window.iconbitmap(os.path.join(cwd, 'Images', 'Setting.ico'))

    setting_window.grid_columnconfigure((0,1,2,3), weight=1) #--Columns are seperated by equal distance

    def Disable_Login_Close():
        pass
    global Gap_Value
    Gap_Value = ARM.BuySellSpread

    Label(setting_window, text="Setting Window",font='Helvetica 12 bold',fg="black",bg='#e5e5ff').grid(row=0,column=0,columnspan=4,sticky='nsew')

    Label(setting_window, text="",fg="black",bg='#e5e5ff',justify="left",wraplength=450).grid(row=1,column=0,columnspan=4,sticky='nsew')
    #-------------------------------------
    Label(setting_window, text="",font='Helvetica 10 bold',fg="black",bg='#ffe5ff').grid(row=3,column=0, columnspan=4,sticky='nsew')
    Label(setting_window, text="",font='Helvetica 10 bold',fg="black",bg='#ffe5ff').grid(row=3,column=0,columnspan=4,sticky='nsew')
    Label(setting_window, text="",bg='#ffe5ff').grid(row=4,column=0,columnspan=4,sticky='nsew')
    Label(setting_window, text="",bg='#ffe5ff').grid(row=5,column=0,columnspan=4,sticky='nsew')

    Label(setting_window, text="",font=('Arial', 2),bg='#e5e5ff').grid(row=6,column=0,columnspan=4,sticky='nsew')
    Label(setting_window, text="",font=('Arial', 2),bg='#e5e5ff').grid(row=7,column=0,columnspan=4,sticky='nsew')

    def save_selection():
        global Gap_Value
        Gap_Entered()
        if Gap_Value != ARM.BuySellSpread:ARM.BuySellSpread = Gap_Value

        print("Entered Gap Value",ARM.BuySellSpread)
        #R0C0options[0] = ARM.your_user_id
        R0C0.set(ARM.your_user_id)
        setting_window.destroy()

    def cancel_selection():
        #R0C0options[0] = ARM.your_user_id
        R0C0.set(ARM.your_user_id)
        setting_window.destroy()

    Label(setting_window, text="",bg='#e5e5ff').grid(row=8,column=0,columnspan=4,sticky='nsew')
    userid_frame = tk.Frame(setting_window,bg='#e5e5ff')
    userid_frame.grid(row=8, column=0,columnspan=4)

    tk.Label(userid_frame,text="Buy/Sell Spread (%)",font='Helvetica 10 bold',fg="black",bg='#e5e5ff',borderwidth=0).pack(side=tk.LEFT,padx=10)

    def Gap_Entered():
        global Gap_Value
        try:
            Gap_Value = float(Gapentry.get())   #--Convert the input to a float
            if -5.0 <= Gap_Value <= 5.0:        #--Check if the value is within the allowed range
                Gapentry.delete(0, tk.END)
                Gapentry.insert(0, Gap_Value)
            else:  #--If the value is outside the range, reset it to ARM.BuySellSpread
                Gap_Value = ARM.BuySellSpread
                Gapentry.delete(0, tk.END)
                Gapentry.insert(0, Gap_Value)
                Gapentry.config(fg='grey')
        except ValueError:  #--Handle invalid input
            Gap_Value = ARM.BuySellSpread
            Gapentry.delete(0, tk.END)
            Gapentry.insert(0, Gap_Value)
            Gapentry.config(fg='grey')
        Save_button.focus_set()

    def on_entry_click(event):
        #if float(Gapentry.get()) == ARM.BuySellSpread:
        Gapentry.delete(0, tk.END)
        Gapentry.config(fg='black')         #--Change text color to black
        #else:pass
        text_length = len(Gapentry.get())
        Gapentry.icursor(text_length // 2)  #--Move cursor to the center of the text

    def on_focus_out(event):
        if not Gapentry.get():
            Gapentry.insert(0, ARM.BuySellSpread)
            Gapentry.config(fg='grey')  #--Change text color to grey

    Gapentry = tk.Entry(userid_frame, font='Helvetica 10', fg='grey', justify='center', width=10)
    Gapentry.insert(0, ARM.BuySellSpread)
    Gapentry.bind("<FocusIn>", on_entry_click)
    Gapentry.bind("<FocusOut>", on_focus_out)
    Gapentry.bind('<Return>', lambda event: Gap_Entered())
    Gapentry.pack(side=tk.RIGHT,padx=10)

    Label(setting_window, text="",font=('Arial', 2),bg='#e5e5ff').grid(row=9,column=0,columnspan=4,sticky='nsew')

    Label(setting_window, text="",bg='#e5e5ff').grid(row=10,column=0,columnspan=4,sticky='nsew')
    button_frame = tk.Frame(setting_window,bg='#e5e5ff')
    button_frame.grid(row=10, column=0,columnspan=4)

    Save_button = tk.Button(button_frame,image=save_button_image,bg='#e5e5ff',command=save_selection,borderwidth=0)
    Save_button.pack(side=tk.LEFT)
    Save_button.bind('<Return>', lambda event: save_selection())

    cancel_button = tk.Button(button_frame,image=cancel_button_image,bg='#e5e5ff',command=cancel_selection,borderwidth=0)
    cancel_button.pack(side=tk.RIGHT)
    cancel_button.bind('<Return>', lambda event: cancel_selection())

    Exchange_Status = Label(setting_window, text="",fg="black",font='Helvetica 8 bold',bg='#e5e5ff')
    Exchange_Status.grid(row=11,column=0,columnspan=4,sticky='nsew')
    Label(setting_window, text="",font=('Arial', 2),bg='#e5e5ff').grid(row=12,column=0,columnspan=4,sticky='nsew')

    Save_button.focus_set()
    setting_window.protocol("WM_DELETE_WINDOW", Disable_Login_Close)

def NF_FUT_SYMBOL():
    #-------------------------NIFTY---------------------------------
    result = ARM.instrument_nse[ARM.instrument_nse['tradingsymbol'] == "NIFTY 50"]
    ARM.SPT_Token = int(result.iloc[0]['instrument_token']) if not result.empty else 0
    return

def STD_StrikeToSubscribe():
    symbol_row_ce = ARM.instrument_nfo[    #--Filter the dataframe based on the provided inputs
                                                (ARM.instrument_nfo['name'] == "NIFTY") &
                                                (ARM.instrument_nfo['expiry'] == ARM.STD_STR_Exp) &
                                                (ARM.instrument_nfo['strike'] == ARM.STD_Strike) &
                                                (ARM.instrument_nfo['instrument_type'] == "CE")
                                            ]
    symbol_row_pe = ARM.instrument_nfo[    #--Filter the dataframe based on the provided inputs
                                                (ARM.instrument_nfo['name'] == "NIFTY") &
                                                (ARM.instrument_nfo['expiry'] == ARM.STD_STR_Exp) &
                                                (ARM.instrument_nfo['strike'] == ARM.STD_Strike) &
                                                (ARM.instrument_nfo['instrument_type'] == "PE")
                                            ]

    if not symbol_row_ce.empty and not symbol_row_pe.empty:   #--Check if there is a matching tradingsymbol
        ARM.STD_STR_Subscribe = pd.concat([
            symbol_row_ce[['instrument_token', 'tradingsymbol', 'strike', 'instrument_type']],
            symbol_row_pe[['instrument_token', 'tradingsymbol', 'strike', 'instrument_type']]
        ], ignore_index=True)

        # print(f"No. of Rows after STD: {len(ARM.STD_STR_Subscribe)}")
        # print(ARM.STD_STR_Subscribe)
    else:
        print("Straddle Strike is not in Main Trading Symbol list")
def STR_StrikeToSubscribe():
    symbol_row_ce = ARM.instrument_nfo[    #--Filter the dataframe based on the provided inputs
                                                (ARM.instrument_nfo['name'] == "NIFTY") &
                                                (ARM.instrument_nfo['expiry'] == ARM.STD_STR_Exp) &
                                                (ARM.instrument_nfo['strike'] == ARM.STR_CE_Strike) &
                                                (ARM.instrument_nfo['instrument_type'] == "CE")
                                            ]
    symbol_row_pe = ARM.instrument_nfo[    #--Filter the dataframe based on the provided inputs
                                                (ARM.instrument_nfo['name'] == "NIFTY") &
                                                (ARM.instrument_nfo['expiry'] == ARM.STD_STR_Exp) &
                                                (ARM.instrument_nfo['strike'] == ARM.STR_PE_Strike) &
                                                (ARM.instrument_nfo['instrument_type'] == "PE")
                                            ]

    if not symbol_row_ce.empty and not symbol_row_pe.empty:   #--Check if there is a matching tradingsymbol
        STR_Subscribe = pd.concat([
            symbol_row_ce[['instrument_token', 'tradingsymbol', 'strike', 'instrument_type']],
            symbol_row_pe[['instrument_token', 'tradingsymbol', 'strike', 'instrument_type']]
        ], ignore_index=True)

        ARM.STD_STR_Subscribe = pd.concat([ARM.STD_STR_Subscribe, STR_Subscribe], ignore_index=True)

        # print(f"No. of Rows after STR: {len(ARM.STD_STR_Subscribe)}")
        # print(ARM.STD_STR_Subscribe)
    else:
        print("Strangle Strike is not in Main Trading Symbol list")

def SYMBOL_TO_TRADE(BuySell_Strike,BuySell_CE_PE):
    mytrade_exp_ip = datetime.strptime(R1C2.get(), '%Y-%m-%d').date()
    mystrike1_ip = BuySell_Strike
    mytype_cepe_ip = BuySell_CE_PE

    symbol_row_cepe = ARM.instrument_nfo[   #--Filter the dataframe based on the provided inputs
                                            (ARM.instrument_nfo['name'] == ARM.oc_symbol) &
                                            (ARM.instrument_nfo['expiry'] == mytrade_exp_ip) &
                                            (ARM.instrument_nfo['strike'] == mystrike1_ip) &
                                            (ARM.instrument_nfo['instrument_type'] == mytype_cepe_ip)
                                        ]

    if not symbol_row_cepe.empty:           #--Check if there is a matching tradingsymbol
        ARM.NF_OPT_Trade_Symbol = symbol_row_cepe['tradingsymbol'].values[0]
    else:
        print(f"No Matching {ARM.oc_symbol} Tradingsymbol CE/PE Found.")
    return

def Clear_ErrorExcptionInfo():
    R2C8entry.config(fg="black",bg="#F0F0F0")
    R2C8.set("")
    ARM.ErrorExcptionInfo_scheduled = None
    return

def check_internet():
    for interface, addrs in psutil.net_if_addrs().items():
        for addr in addrs:
            #--Check for IPv4 addresses (ignore localhost)
            if addr.family == socket.AF_INET and not addr.address.startswith('127.'):
                #--Check if the interface is up
                if interface in psutil.net_if_stats() and psutil.net_if_stats()[interface].isup:
                    if ARM.KWS_RestartCheck:
                        ARM.KWS_RestartCheck = False
                        ARM.KWS_URestart = True
                        print("{} - ! ! !- .. Recovered Internet Connection .. - ! ! ! -".format(time.strftime("%H:%M:%S", time.localtime())))
                        R2C8.set("Recovered Internet Connection")
                        R2C8entry.config(fg="white",bg="green")
                        #if ARM.ErrorExcptionInfo_scheduled is None: ARM.ErrorExcptionInfo_scheduled = root.after(2000,Clear_ErrorExcptionInfo)
                        ARM.InternetLost_Once = True
                    else:pass
                    return False    #--Returns False if Internet is Connected
    return True                     #--Returns True if Internet is Not Connected

def OTM_Blink_Status():
    #-----------------PART-I--------------------
    ARM.Internet_NotConnected = check_internet()
    if ARM.Internet_NotConnected and ARM.InternetLost_Once:
        print("{} - ! ! !- ..   Lost Internet Connection    .. - ! ! ! -".format(time.strftime("%H:%M:%S", time.localtime())))
        ARM.InternetLost_Once = False
    else:pass
    if(R0C0.get() != "Login"):
        if ARM.Internet_NotConnected:
            R2C8.set("Lost Internet Connection")
            R2C8entry.config(fg="white",bg="red")
            #if ARM.ErrorExcptionInfo_scheduled is None: ARM.ErrorExcptionInfo_scheduled = root.after(2000,Clear_ErrorExcptionInfo)
            ARM.KWS_RestartCheck = True
        else:
            if (GUI.R0C5flag == "DEMO"):
                if (not ARM.Program_Running):
                    R0C5.configure(image=demod_image,borderwidth=0)
                    ARM.Program_Running = True
                else:
                    R0C5.configure(image=demof_image,borderwidth=0)
                    ARM.Program_Running = False
            elif (GUI.R0C5flag == "REAL"):
                if (not ARM.Program_Running):
                    R0C5.configure(image=reald_image,borderwidth=0)
                    ARM.Program_Running = True
                else:
                    R0C5.configure(image=realf_image,borderwidth=0)
                    ARM.Program_Running = False
            else:pass
    else:pass
    #-----------------PART-II -------------------
    if ARM.Trade_Start_Time <= datetime.now().time() <= ARM.Trade_Stop_Time:
        if not ARM.KWS_Functional:
            if ARM.Internet_NotConnected:       #--Executes True body if Internet is Not Connected
                ARM.KWS_Functional = False
                ARM.KWS_Check = 0
            else:                               #--Executes False body if Internet is Connected
                if ARM.KWS_Check >= 120:
                    kws_canvas.itemconfig(square, fill='red')
                    ARM.KWS_URestart = True
                else:
                    ARM.KWS_Check += 1
        else:
            ARM.KWS_Functional = False
            ARM.KWS_Check = 0
    else:pass
    #-----------------PART-III -------------------
    if ARM.STD_Strike > 0 and ARM.STD_CE_Symbol != "" and ARM.STD_PE_Symbol != "":
        if ARM.STD_STR_Subscribe['tradingsymbol'].isin(ARM.KWS_OPT_Data).all() and len(ARM.STD_STR_Subscribe) >= 4:
            ARM.STD_CE_Price_Cur = ARM.KWS_OPT_Data[ARM.STD_CE_Symbol]['ltp']
            ARM.STD_PE_Price_Cur = ARM.KWS_OPT_Data[ARM.STD_PE_Symbol]['ltp']
            ARM.STR_CE_Price_Cur = ARM.KWS_OPT_Data[ARM.STR_CE_Symbol]['ltp']
            ARM.STR_PE_Price_Cur = ARM.KWS_OPT_Data[ARM.STR_PE_Symbol]['ltp']

            ARM.STD_UDS_Cur = round((((ARM.STD_CE_Price_Cur - ARM.STD_CE_Price_Org) - (ARM.STD_PE_Price_Cur - ARM.STD_PE_Price_Org))*100/(ARM.STD_CE_Price_Org + ARM.STD_PE_Price_Org)), ndigits = 2)
            # ARM.STD_UDS_df.loc[len(ARM.STD_UDS_df), 'UDS'] = ARM.STD_UDS_Cur    #--Not Implemented, so commented

            ARM.STR_UDS_Cur = round((((ARM.STR_CE_Price_Cur - ARM.STR_CE_Price_Org) - (ARM.STR_PE_Price_Cur - ARM.STR_PE_Price_Org))*100/(ARM.STR_CE_Price_Org + ARM.STR_PE_Price_Org)), ndigits = 2)
            # ARM.STR_UDS_df.loc[len(ARM.STR_UDS_df), 'UDS'] = ARM.STR_UDS_Cur    #--Not Implemented, so commented

            ARM.AVG_UDS_Cur = round(((ARM.STD_UDS_Cur + ARM.STR_UDS_Cur)/2), ndigits = 2)
            ARM.AVG_UDS_df.loc[len(ARM.AVG_UDS_df), 'UDS'] = ARM.AVG_UDS_Cur

            # ARM.RISING_EDGE_df.loc[len(ARM.RISING_EDGE_df),'BUY_CE'] = ARM.CONFIRM_BUYCE_Cur        #--Not Implemented, so commented
            # ARM.FALLING_EDGE_df.loc[len(ARM.FALLING_EDGE_df), 'BUY_PE'] = ARM.CONFIRM_BUYPE_Cur     #--Not Implemented, so commented

            # Step through values comparing with previous
            # store the previous UDS value and compare the

            Uds_Diff = (ARM.AVG_UDS_Cur - ARM.AVG_UDS_Prev)
            #--rising edge logic
            if Uds_Diff > ARM.STD_STR_MIN_diff and ARM.AVG_UDS_Cur > ARM.UDS_DEADBAND:
                ARM.R_EDGE_Counter += 1
            else:
                ARM.R_EDGE_Counter = 0

            if ARM.R_EDGE_Counter >= ARM.CONSECUTIVE_BAR_value:
                ARM.CONFIRM_BUYCE_Cur = True
            else:
                ARM.CONFIRM_BUYCE_Cur = False

            #--falling edge logic
            if Uds_Diff < -ARM.STD_STR_MIN_diff and ARM.AVG_UDS_Cur < -ARM.UDS_DEADBAND:
                ARM.F_EDGE_Counter += 1
            else:
                ARM.F_EDGE_Counter = 0

            if ARM.F_EDGE_Counter >= ARM.CONSECUTIVE_BAR_value:
                ARM.CONFIRM_BUYPE_Cur = True
            else:
                ARM.CONFIRM_BUYPE_Cur = False

            ARM.AVG_UDS_Prev = ARM.AVG_UDS_Cur

            # Append to CSV (non-blocking)
            if ARM.std_str_csv_writer:
                timestamp = time.strftime("%H:%M:%S")
                ARM.std_str_csv_writer.writerow([
                                            timestamp,
                                            ARM.STD_CE_Price_Cur,
                                            ARM.STD_PE_Price_Cur,
                                            ARM.STD_UDS_Cur,
                                            ARM.STR_CE_Price_Cur,
                                            ARM.STR_PE_Price_Cur,
                                            ARM.STR_UDS_Cur,
                                            ARM.AVG_UDS_Cur,
                                            ARM.CONFIRM_BUYCE_Cur,
                                            ARM.CONFIRM_BUYPE_Cur,
                                            ARM.R_EDGE_Counter,
                                            ARM.F_EDGE_Counter,
                                            f"{Uds_Diff:.3f}"
                                        ])
                ARM.std_str_file.flush()  # Ensure write to disk without closing

            ARM.STD_CombinePrice_Cur = round((ARM.STD_CE_Price_Cur + ARM.STD_PE_Price_Cur), ndigits = 2)
            ARM.STR_CombinePrice_Cur = round((ARM.STR_CE_Price_Cur + ARM.STR_PE_Price_Cur), ndigits = 2)

            ARM.STD_CE_Per_Change = ((ARM.STD_CE_Price_Cur - ARM.STD_CE_Price_Org) / ARM.STD_CE_Price_Org) * 100
            ARM.STD_PE_Per_Change = ((ARM.STD_PE_Price_Cur - ARM.STD_PE_Price_Org) / ARM.STD_PE_Price_Org) * 100
            ARM.STD_Combine_Per_Change = ((ARM.STD_CombinePrice_Cur - ARM.STD_CombPrice_Org) / ARM.STD_CombPrice_Org) * 100

            ARM.STR_CE_Per_Change = ((ARM.STR_CE_Price_Cur - ARM.STR_CE_Price_Org) / ARM.STR_CE_Price_Org) * 100
            ARM.STR_PE_Per_Change = ((ARM.STR_PE_Price_Cur - ARM.STR_PE_Price_Org) / ARM.STR_PE_Price_Org) * 100
            ARM.STR_Combine_Per_Change = ((ARM.STR_CombinePrice_Cur - ARM.STR_CombPrice_Org) / ARM.STR_CombPrice_Org) * 100
            #------------- Indicate Directional Trend based on UDS Start -------------
            if not ARM.CONFIRM_BUYCE_Cur and not ARM.CONFIRM_BUYPE_Cur:
                if ARM.Directional_Trend_UDS:
                    uds_ce_canvas.itemconfig(circle, fill="yellow")     #--Sideways → yellow    🟡
                    uds_pe_canvas.itemconfig(circle, fill="yellow")     #--Sideways → yellow    🟡
                    ARM.Directional_Trend_UDS = False
                else:pass
            elif ARM.CONFIRM_BUYCE_Cur and not ARM.CONFIRM_BUYPE_Cur:
                if not ARM.Directional_Trend_UDS:
                    uds_ce_canvas.itemconfig(circle, fill="green")      #--Uptrend → green (BUY CE) 🟢
                    uds_pe_canvas.itemconfig(circle, fill="#FCDFFF")
                    ARM.Directional_Trend_UDS = True
                    ARM.R_EDGE_Counter = 0
                    on_buy_click(event=None, row=15, column=4)          #--Click on ATM CE to Buy Automatically ()
                    # root.after(0, lambda: on_buy_click(None, 15, 4))  #--Schedule after current function execution
                else:pass
            elif not ARM.CONFIRM_BUYCE_Cur and ARM.CONFIRM_BUYPE_Cur:
                if not ARM.Directional_Trend_UDS:
                    uds_ce_canvas.itemconfig(circle, fill="#FCDFFF")
                    uds_pe_canvas.itemconfig(circle, fill="green")      #--Downtrend → green (BUY PE) 🟢
                    ARM.Directional_Trend_UDS = True
                    ARM.F_EDGE_Counter = 0
                    on_buy_click(event=None, row=15, column=6)          #--Click on ATM PE to Buy Automatically
                    # root.after(0, lambda: on_buy_click(None, 15, 6))  #--Schedule after current function execution
                else:pass
            else:
                print("Something is wrong check / verify your logic uds ce/pe canvas ")
            #------------- Indicate Directional Trend based on UDS End -------------

            #------------- Indicate Directional Trend based on STD STR Start -------------
            if (ARM.Direction_Signal_On == "IND"):
                if (ARM.STD_CE_Per_Change > ARM.Individual_Per):
                    if not ARM.Directional_Trend_STD_STR:
                        ce_canvas.itemconfig(circle, fill="green")      #--Uptrend → green (BUY CE) 🟢 🔴 🟡
                        pe_canvas.itemconfig(circle, fill="#FCDFFF")
                        ARM.Directional_Trend_STD_STR = True
                    else:pass
                elif (ARM.STD_PE_Per_Change > ARM.Individual_Per):
                    if not ARM.Directional_Trend_STD_STR:
                        ce_canvas.itemconfig(circle, fill="#FCDFFF")
                        pe_canvas.itemconfig(circle, fill="green")      #--Downtrend → green (BUY PE) 🟢 🔴 🟡
                        ARM.Directional_Trend_STD_STR = True
                    else:pass
                else:
                    if ARM.Directional_Trend_STD_STR:
                        ce_canvas.itemconfig(circle, fill="yellow")    #--Sideways → yellow    🟢 🔴 🟡
                        pe_canvas.itemconfig(circle, fill="yellow")    #--Sideways → yellow    🟢 🔴 🟡
                        ARM.Directional_Trend_STD_STR = False
                    else:pass
            elif (ARM.Direction_Signal_On == "COM"):
                if ARM.STD_Combine_Per_Change <= ARM.Combine_Per:
                    if ARM.Directional_Trend_STD_STR:
                        ce_canvas.itemconfig(circle, fill="yellow")    #--Sideways → yellow    🟢 🔴 🟡
                        pe_canvas.itemconfig(circle, fill="yellow")    #--Sideways → yellow    🟢 🔴 🟡
                        ARM.Directional_Trend_STD_STR = False
                    else:pass
                elif ARM.STD_Combine_Per_Change > ARM.Combine_Per:
                    if ARM.STD_CE_Price_Cur > ARM.STD_CE_Price_Org and ARM.STD_PE_Price_Cur < ARM.STD_PE_Price_Org:
                        if not ARM.Directional_Trend_STD_STR:
                            # bg_process.Play_Beep()
                            ce_canvas.itemconfig(circle, fill="green")      #--Uptrend → green (BUY CE) 🟢 🔴 🟡
                            pe_canvas.itemconfig(circle, fill="#FCDFFF")
                            ARM.Directional_Trend_STD_STR = True
                        else:pass
                    elif ARM.STD_PE_Price_Cur > ARM.STD_PE_Price_Org and ARM.STD_CE_Price_Cur < ARM.STD_CE_Price_Org:
                        if not ARM.Directional_Trend_STD_STR:
                            # bg_process.Play_Beep()
                            ce_canvas.itemconfig(circle, fill="#FCDFFF")
                            pe_canvas.itemconfig(circle, fill="green")      #--Downtrend → green (BUY PE) 🟢 🔴 🟡
                            ARM.Directional_Trend_STD_STR = True
                        else:pass
                    else:pass
                else:pass
            elif (ARM.Direction_Signal_On == "HYB"):
                if ARM.STD_Combine_Per_Change <= ARM.Combine_Per:
                    if ARM.Directional_Trend_STD_STR:
                        ce_canvas.itemconfig(circle, fill="yellow")    #--Sideways → yellow    🟢 🔴 🟡
                        pe_canvas.itemconfig(circle, fill="yellow")    #--Sideways → yellow    🟢 🔴 🟡
                        ARM.Directional_Trend_STD_STR = False
                    else:pass
                elif ARM.STD_Combine_Per_Change > ARM.Combine_Per:
                    if (ARM.STD_CE_Per_Change > ARM.Individual_Per):
                        if not ARM.Directional_Trend_STD_STR:
                            ce_canvas.itemconfig(circle, fill="green")      #--Uptrend → green (BUY CE) 🟢 🔴 🟡
                            pe_canvas.itemconfig(circle, fill="#FCDFFF")
                            ARM.Directional_Trend_STD_STR = True
                        else:pass
                    elif (ARM.STD_PE_Per_Change > ARM.Individual_Per):
                        if not ARM.Directional_Trend_STD_STR:
                            ce_canvas.itemconfig(circle, fill="#FCDFFF")
                            pe_canvas.itemconfig(circle, fill="green")      #--Downtrend → green (BUY PE) 🟢 🔴 🟡
                            ARM.Directional_Trend_STD_STR = True
                        else:pass
                    else:pass
                else:pass
            else:
                print("Something is wrong check / verify your logic STD/STR ce/pe canvas ")
            #------------- Logic to change Yellow-Green-Yellow : Extreem Right & Left Signal End -------------
        else:
            ce_canvas.itemconfig(circle, fill="#FCDFFF")
            pe_canvas.itemconfig(circle, fill="#FCDFFF")
            # A0.setText("")
            # B5.setText("")
            # F5.setText("")
            print("ATM Straddle CE/PE Symbol Token not Subscribed/Un-Subscribed")
    else:pass
    #---------------------------------------------
    root.after(1000,OTM_Blink_Status)

def extract_info_from_symbol(symbol):
    result = ARM.instrument_nfo[ARM.instrument_nfo["tradingsymbol"] == symbol][["name", "strike", "expiry", "instrument_type"]]
    if not result.empty:
        return result.iloc[0]["name"], result.iloc[0]["strike"], result.iloc[0]["instrument_type"],result.iloc[0]["expiry"]
    else:
        print("Symbol Extraction failed")
        return None, None, None, None

def APH_MAIN(stop_scheduled=False):
    global keep_scheduled,kws_nse

    if ARM.MY_TRADE == "REAL":
        if not check_internet():
            try:
                Margin_Info = kite.margins()
                ARM.Live_Balance = Margin_Info["equity"]["available"]["live_balance"]
            except:pass
        else:pass
        R1C4.set("{:.2f}".format(ARM.Live_Balance))
    elif ARM.MY_TRADE == "DEMO":ARM.Live_Balance = R1C4.get()
    else:pass

    if not stop_scheduled:
        if not ARM.Ready_Once:ARM.oc_symbol = ARM.APH_SharedPVar[0]

        #=====================================================================
        if (ARM.pre_oc_symbol != ARM.oc_symbol or ARM.pre_oc_opt_expiry != ARM.oc_opt_expiry): #--Symbol/ Future / Option Expiry Changed
            if (ARM.pre_oc_symbol != ARM.oc_symbol):
                ARM.instrument_dict_opt = {}
                #---------------------------
                df = copy.deepcopy(ARM.fut_exchange_nfo)
                df = df[df["name"] == ARM.oc_symbol]
                #---------------------------
                ARM.oc_opt_expiry = datetime.strptime(R1C2.get(), '%Y-%m-%d').date()
                # print("Changed Symbol to NIFTY")

                ARM.Symbol_Changed = True
            elif(ARM.oc_symbol == "NIFTY"):
                if (ARM.pre_oc_opt_expiry != ARM.oc_opt_expiry):
                    ARM.instrument_dict_opt = {}

                    ARM.oc_opt_expiry = datetime.strptime(R1C2.get(), '%Y-%m-%d').date()
                    print(f"Changed Option Expiry of NIFTY to {ARM.oc_opt_expiry}")

                    ARM.Opt_Exp_Changed = True
                else:pass
            else:pass

            ARM.stop_thread_opt = True
            ARM.stop_thread_fut = True
            ARM.pre_oc_symbol = ARM.oc_symbol
            ARM.pre_oc_opt_expiry = ARM.oc_opt_expiry
        else:pass

        if (ARM.oc_symbol != ""):
            #try:
            if not ARM.instrument_dict_opt and ARM.oc_opt_expiry is not None: #--Find Option Lot
                df_opt1 = copy.deepcopy(ARM.opt_exchange_nfo)
                df_opt1 = df_opt1[df_opt1["name"] == ARM.oc_symbol]
                #df_opt1 = df_opt1[df_opt1["expiry"] == ARM.oc_opt_expiry.date()]
                df_opt1 = df_opt1[df_opt1["expiry"] == ARM.oc_opt_expiry]
                ARM.found_lotsize_opt = list(df_opt1["lot_size"])[0]      #--Added for OC
                #print("found_lotsize_opt",ARM.found_lotsize_opt)
                #print(df_fut)
                for i in df_opt1.index:
                    ARM.instrument_dict_opt[f'NFO:{df_opt1["tradingsymbol"][i]}'] = {"strikePrice": float(df_opt1["strike"][i]),
                                                                    "instrumentType": df_opt1["instrument_type"][i],
                                                                    "token": df_opt1["instrument_token"][i]}

                ARM.stop_thread_opt = False
                thread2 = threading.Thread(target=get_oi_opt, args=(ARM.instrument_dict_opt,)) #--Check Thread for Option/Future
                thread2.start()
            else:pass
            #------------------------------------------------------------------------
            if (ARM.oc_opt_expiry is not None and ARM.KWS_OPT_Data and not ARM.kws_option_flag and
                not ARM.Symbol_Changed and not ARM.Opt_Exp_Changed and not ARM.OptFut_Exp_Changed):
                #----Then Process Option Data-------------------------------------------------------------------

                df_opt = pd.DataFrame(ARM.KWS_OPT_Data).transpose()

                df_opt = df_opt[df_opt.index.str.startswith('NIFTY')]

                df_opt = df_opt[(df_opt["expiry"] == ARM.oc_opt_expiry)]
                # df_opt.to_csv("df_opt.csv", index=False)
                # df_opt = df_opt[~df_opt.index.isin(ARM.symbols_in_net_positions)].reset_index()

                #----------------------------
                # if not ARM.prev_day_oi_opt_df.empty:
                #     #df_opt.to_csv("df_opt.csv", index=False)
                #     df_opt = df_opt.reset_index().rename(columns={'index': 'TradeSymbol'})
                #     #df_opt.to_csv("df_opt_before.csv", index=False) #--Export Before finding Change in OI
                #     df_opt = df_opt.merge(ARM.prev_day_oi_opt_df, on='TradeSymbol', how='left')  # Merge to get Prv_OI
                #     df_opt['changeinOpenInterest'] = df_opt.apply(
                #                             lambda row: (row['changeinOpenInterest'] - row['Prv_OI']) / ARM.found_lotsize_opt
                #                             if not pd.isna(row['Prv_OI']) else row['changeinOpenInterest'],
                #                             axis=1
                #                         )
                #     df_opt = df_opt.drop(columns=['Prv_OI','TradeSymbol'])
                #     #df_opt.to_csv("df_opt_after.csv", index=False) #--Export After finding Change in OI
                # else:pass
                #----------------------------

                ce_df = df_opt[df_opt["instrumentType"] == "CE"]
                ce_df = ce_df[["change", "ltp", "bid","ask","averagePrice", "strikePrice"]]
                ce_df = ce_df.rename(columns={"change": "CE_CLTP",
                                            "ltp": "CE_LTP","bid":"CE_Bid",
                                            "ask":"CE_Ask","averagePrice":"CE_ATP"})
                ce_df.index = ce_df["strikePrice"]
                ce_df = ce_df.drop(["strikePrice"], axis=1)
                ce_df["STRIKE"] = ce_df.index

                pe_df = df_opt[df_opt["instrumentType"] == "PE"]
                pe_df = pe_df[["strikePrice", "bid", "ask", "averagePrice", "ltp", "change"]]
                pe_df = pe_df.rename(columns={"change": "PE_CLTP",
                                            "ltp": "PE_LTP","bid":"PE_Bid",
                                            "ask":"PE_Ask","averagePrice":"PE_ATP"})
                pe_df.index = pe_df["strikePrice"]
                pe_df = pe_df.drop(["strikePrice"], axis=1)

                df_opt = pd.concat([ce_df, pe_df], axis=1).sort_index()
                df_opt = df_opt.replace(np.nan, 0)
                df_opt["STRIKE"] = df_opt.index
                df_opt = df_opt.reset_index(drop=True) #--Reset Index
                #print("Total Number of Strikes",len(df_opt))

                atm_index = df_opt['STRIKE'].sub(ARM.Spot_Value).abs().astype('float64').idxmin()   #--Find ATM Index Location
                maxlen = len(df_opt)
                df_opt.drop(df_opt.loc[0:atm_index-ARM.OC_Strikes-1].index, inplace=True)           #--Maximum Strike on both side of ATM is 10
                df_opt.drop(df_opt.loc[atm_index+ARM.OC_Strikes+1:maxlen-1].index, inplace=True)    #--Maximum Strike on both side of ATM is 10
                df_opt = df_opt.reset_index(drop=True) # Reset Index

                # df_opt.to_csv("Available_Strikes3.csv", index=True)
                # print(f"len(df_opt): {len(df_opt)}, OC_Strikes: {ARM.OC_Strikes}, Display_Row :{ARM.Display_Row}")
                if ((len(df_opt) == (2*ARM.OC_Strikes)+1) and (1 <= ARM.OC_Strikes <= 10)):
                    df_opt = df_opt[["CE_CLTP","CE_LTP","CE_Bid","CE_ATP","CE_Ask",
                                    "STRIKE",
                                    "PE_Ask","PE_ATP","PE_Bid","PE_LTP","PE_CLTP"]]

                    #print(df_opt)
                    if not ARM.df_opt_final.empty:
                        ARM.df_opt_final.drop(ARM.df_opt_final.index, inplace=True)
                    ARM.df_opt_final = df_opt.copy(deep=True)
                    # print(ARM.df_opt_final)
                    Update_Popup_OC()
                    #print("OPtion Data Printed")
                    if ARM.STD_CE_Symbol == "" and ARM.STD_PE_Symbol == "":
                        if datetime.now().time() >= ARM.Straddle_Start_Time:
                            if not ARM.Use_Initial_Strikes:
                                #-- Use ARM.df_opt_final, and create ARM.Initial_Strikes_df
                                if ARM.STD_STR_Subscribe.empty:
                                    if len(ARM.Initial_Strikes_df) <= 0:
                                        #--------------------------------
                                        ARM.Initial_Strikes_df = (
                                                                    ARM.df_opt_final[["STRIKE", "CE_LTP", "PE_LTP"]]
                                                                    .copy()
                                                                    .reset_index(drop=True)
                                                                )

                                        # Optional safety check
                                        if len(ARM.Initial_Strikes_df) != 21:
                                            raise ValueError(
                                                                f"Expected {21} strikes, "
                                                                f"found {len(ARM.Initial_Strikes_df)}"
                                                            )

                                        ARM.Initial_Strikes_df.to_csv(Initial_Strikes_Path, index=False)
                                        print("Created Initial_Strikes.csv")
                                        #--------------------------------
                                    else:
                                        print("Some problem")
                                        print("There should not be any Data in ARM.Initial_Strikes_df")
                                    if not ARM.STD_Strike_Change:
                                        Synthetic_Fut = ARM.df_opt_final.loc[10, 'STRIKE'] + ARM.df_opt_final.loc[10, 'CE_LTP'] - ARM.df_opt_final.loc[10, 'PE_LTP']
                                        ARM.STD_Strike = round(Synthetic_Fut / 50) * 50
                                        ARM.STD_STR_Exp = datetime.strptime(R1C2.get(), '%Y-%m-%d').date()
                                    else:pass
                                    print(f"STD_Strike: {ARM.STD_Strike}")
                                    STD_StrikeToSubscribe()
                                elif ARM.STD_STR_Subscribe['tradingsymbol'].isin(ARM.KWS_OPT_Data).all() and ARM.STR_CE_Strike == 0  and ARM.STR_PE_Strike == 0:
                                    STD_CE_Symbol = ARM.STD_STR_Subscribe.loc[0, 'tradingsymbol']
                                    STD_PE_Symbol = ARM.STD_STR_Subscribe.loc[1, 'tradingsymbol']

                                    STD_CE_Price_Org = ARM.KWS_OPT_Data[STD_CE_Symbol]['ltp']
                                    STD_PE_Price_Org = ARM.KWS_OPT_Data[STD_PE_Symbol]['ltp']

                                    ARM.STR_CE_Strike = round((ARM.STD_Strike + STD_CE_Price_Org + STD_PE_Price_Org) / 50) * 50
                                    ARM.STR_PE_Strike = round((ARM.STD_Strike - STD_CE_Price_Org - STD_PE_Price_Org) / 50) * 50

                                    print(f"STR_CE_Strike = {ARM.STR_CE_Strike}")
                                    print(f"STR_PE_Strike = {ARM.STR_PE_Strike}")

                                    STR_StrikeToSubscribe()
                                elif ARM.STD_STR_Subscribe['tradingsymbol'].isin(ARM.KWS_OPT_Data).all() and len(ARM.STD_STR_Subscribe) >= 4:
                                    ARM.STD_CE_Symbol = ARM.STD_STR_Subscribe.loc[0, 'tradingsymbol']
                                    ARM.STD_PE_Symbol = ARM.STD_STR_Subscribe.loc[1, 'tradingsymbol']
                                    ARM.STR_CE_Symbol = ARM.STD_STR_Subscribe.loc[2, 'tradingsymbol']
                                    ARM.STR_PE_Symbol = ARM.STD_STR_Subscribe.loc[3, 'tradingsymbol']

                                    ARM.STD_CE_Price_Org = ARM.KWS_OPT_Data[ARM.STD_CE_Symbol]['ltp']
                                    ARM.STD_PE_Price_Org = ARM.KWS_OPT_Data[ARM.STD_PE_Symbol]['ltp']
                                    ARM.STR_CE_Price_Org = ARM.KWS_OPT_Data[ARM.STR_CE_Symbol]['ltp']
                                    ARM.STR_PE_Price_Org = ARM.KWS_OPT_Data[ARM.STR_PE_Symbol]['ltp']

                                    # ARM.STD_UDS_df.loc[len(ARM.STD_UDS_df), 'UDS'] = 0  #--Not Implemented, so commented
                                    # ARM.STR_UDS_df.loc[len(ARM.STR_UDS_df), 'UDS'] = 0  #--Not Implemented, so commented
                                    ARM.AVG_UDS_df.loc[len(ARM.AVG_UDS_df), 'UDS'] = 0
                                    # ARM.RISING_EDGE_df.loc[len(ARM.RISING_EDGE_df),'BUY_CE'] = False    #--Not Implemented, so commented
                                    # ARM.RISING_EDGE_df.loc[len(ARM.RISING_EDGE_df),'BUY_CE'] = False    #--Not Implemented, so commented

                                    # Append to CSV (non-blocking)
                                    if ARM.std_str_csv_writer:
                                        timestamp = time.strftime("%H:%M:%S")
                                        ARM.std_str_csv_writer.writerow([
                                                                    timestamp,
                                                                    ARM.STD_CE_Symbol,
                                                                    ARM.STD_PE_Symbol,
                                                                    "UDS",
                                                                    ARM.STR_CE_Symbol,
                                                                    ARM.STR_PE_Symbol,
                                                                    "UDS",
                                                                    "UDS",
                                                                    "Signal",
                                                                    "Signal",
                                                                    "Rising",
                                                                    "Falling",
                                                                    "Diff"
                                                                ])
                                        ARM.std_str_csv_writer.writerow([
                                                                    timestamp,
                                                                    ARM.STD_CE_Price_Org,
                                                                    ARM.STD_PE_Price_Org,
                                                                    0,
                                                                    ARM.STR_CE_Price_Org,
                                                                    ARM.STR_PE_Price_Org,
                                                                    0,
                                                                    0,
                                                                    "False",
                                                                    "False",
                                                                    0,
                                                                    0,
                                                                    0
                                                                ])
                                        ARM.std_str_file.flush()  # Ensure write to disk without closing

                                    ARM.STD_CombPrice_Org = round((ARM.STD_CE_Price_Org + ARM.STD_PE_Price_Org), ndigits = 2)
                                    ARM.STR_CombPrice_Org = round((ARM.STR_CE_Price_Org + ARM.STR_PE_Price_Org), ndigits = 2)

                                    ARM.STD_Dis_MBEP = ARM.STD_Strike - ARM.STD_CombPrice_Org
                                    ARM.STD_Dis_STD = ARM.Spot_Value
                                    ARM.STD_Dis_PBEP = ARM.STD_Strike + ARM.STD_CombPrice_Org
                                    ARM.STD_Dis_Time = time.strftime("%H:%M:%S")    #--"12:12:12"

                                    ARM.STD_Strike_Change = False
                                    print("--- All 4 (STD-2 & STR-2) Strikes Subscribed & Tick Data captured/started ---")
                                else:pass
                            else:
                                #-- Use ARM.Initial_Strikes_df, already created ARM.Initial_Strikes_df
                                if ARM.STD_STR_Subscribe.empty:
                                    if not ARM.STD_Strike_Change:
                                        Synthetic_Fut = ARM.Initial_Strikes_df.loc[10, 'STRIKE'] + ARM.Initial_Strikes_df.loc[10, 'CE_LTP'] - ARM.Initial_Strikes_df.loc[10, 'PE_LTP']
                                        ARM.STD_Strike = round(Synthetic_Fut / 50) * 50
                                        ARM.STD_STR_Exp = datetime.strptime(R1C2.get(), '%Y-%m-%d').date()
                                    else:pass
                                    print(f"STD_Strike: {ARM.STD_Strike}")
                                    STD_StrikeToSubscribe()
                                elif ARM.STD_STR_Subscribe['tradingsymbol'].isin(ARM.KWS_OPT_Data).all() and ARM.STR_CE_Strike == 0  and ARM.STR_PE_Strike == 0:
                                    STD_CE_Symbol = ARM.STD_STR_Subscribe.loc[0, 'tradingsymbol']
                                    STD_PE_Symbol = ARM.STD_STR_Subscribe.loc[1, 'tradingsymbol']

                                    ARM.STD_CE_Price_Org = ARM.Initial_Strikes_df.loc[ARM.Initial_Strikes_df["STRIKE"] == ARM.STD_Strike,"CE_LTP"].iloc[0]
                                    ARM.STD_PE_Price_Org = ARM.Initial_Strikes_df.loc[ARM.Initial_Strikes_df["STRIKE"] == ARM.STD_Strike,"PE_LTP"].iloc[0]

                                    print(f"STD CE_LTP : {ARM.STD_CE_Price_Org}, CE_LTP : {ARM.STD_PE_Price_Org}")

                                    ARM.STR_CE_Strike = round((ARM.STD_Strike + ARM.STD_CE_Price_Org + ARM.STD_PE_Price_Org) / 50) * 50
                                    ARM.STR_PE_Strike = round((ARM.STD_Strike - ARM.STD_CE_Price_Org - ARM.STD_PE_Price_Org) / 50) * 50

                                    print(f"STR_CE_Strike = {ARM.STR_CE_Strike}")
                                    print(f"STR_PE_Strike = {ARM.STR_PE_Strike}")

                                    STR_StrikeToSubscribe()
                                elif ARM.STD_STR_Subscribe['tradingsymbol'].isin(ARM.KWS_OPT_Data).all() and len(ARM.STD_STR_Subscribe) >= 4:
                                    ARM.STD_CE_Symbol = ARM.STD_STR_Subscribe.loc[0, 'tradingsymbol']
                                    ARM.STD_PE_Symbol = ARM.STD_STR_Subscribe.loc[1, 'tradingsymbol']
                                    ARM.STR_CE_Symbol = ARM.STD_STR_Subscribe.loc[2, 'tradingsymbol']
                                    ARM.STR_PE_Symbol = ARM.STD_STR_Subscribe.loc[3, 'tradingsymbol']

                                    ARM.STR_CE_Price_Org = ARM.Initial_Strikes_df.loc[ARM.Initial_Strikes_df["STRIKE"] == ARM.STR_CE_Strike,"CE_LTP"].iloc[0]
                                    ARM.STR_PE_Price_Org = ARM.Initial_Strikes_df.loc[ARM.Initial_Strikes_df["STRIKE"] == ARM.STR_PE_Strike,"PE_LTP"].iloc[0]

                                    print(f"STR CE_LTP : {ARM.STR_CE_Price_Org}, CE_LTP : {ARM.STR_PE_Price_Org}")

                                    # ARM.STD_UDS_df.loc[len(ARM.STD_UDS_df), 'UDS'] = 0  #--Not Implemented, so commented
                                    # ARM.STR_UDS_df.loc[len(ARM.STR_UDS_df), 'UDS'] = 0  #--Not Implemented, so commented
                                    ARM.AVG_UDS_df.loc[len(ARM.AVG_UDS_df), 'UDS'] = 0
                                    # ARM.RISING_EDGE_df.loc[len(ARM.RISING_EDGE_df),'BUY_CE'] = False    #--Not Implemented, so commented
                                    # ARM.RISING_EDGE_df.loc[len(ARM.RISING_EDGE_df),'BUY_CE'] = False    #--Not Implemented, so commented

                                    # Append to CSV (non-blocking)
                                    if ARM.std_str_csv_writer:
                                        timestamp = time.strftime("%H:%M:%S")
                                        ARM.std_str_csv_writer.writerow([
                                                                    timestamp,
                                                                    ARM.STD_CE_Symbol,
                                                                    ARM.STD_PE_Symbol,
                                                                    "UDS",
                                                                    ARM.STR_CE_Symbol,
                                                                    ARM.STR_PE_Symbol,
                                                                    "UDS",
                                                                    "UDS",
                                                                    "Signal",
                                                                    "Signal",
                                                                    "Rising",
                                                                    "Falling",
                                                                    "Diff"
                                                                ])
                                        ARM.std_str_csv_writer.writerow([
                                                                    timestamp,
                                                                    ARM.STD_CE_Price_Org,
                                                                    ARM.STD_PE_Price_Org,
                                                                    0,
                                                                    ARM.STR_CE_Price_Org,
                                                                    ARM.STR_PE_Price_Org,
                                                                    0,
                                                                    0,
                                                                    "False",
                                                                    "False",
                                                                    0,
                                                                    0,
                                                                    0
                                                                ])
                                        ARM.std_str_file.flush()  # Ensure write to disk without closing

                                    ARM.STD_CombPrice_Org = round((ARM.STD_CE_Price_Org + ARM.STD_PE_Price_Org), ndigits = 2)
                                    ARM.STR_CombPrice_Org = round((ARM.STR_CE_Price_Org + ARM.STR_PE_Price_Org), ndigits = 2)

                                    ARM.STD_Dis_MBEP = ARM.STD_Strike - ARM.STD_CombPrice_Org
                                    ARM.STD_Dis_STD = ARM.Spot_Value
                                    ARM.STD_Dis_PBEP = ARM.STD_Strike + ARM.STD_CombPrice_Org
                                    ARM.STD_Dis_Time = time.strftime("%H:%M:%S")    #--"12:12:12"

                                    ARM.STD_Strike_Change = False
                                    print("--- All 4 (STD-2 & STR-2) Strikes Subscribed & Tick Data captured/started ---")
                                else:pass
                        else:pass
                        # print("--------------- ATM Straddle Values ---------------")
                        # print(f"{ARM.STD_CE_Price_Org:.2f} : {ARM.STD_CE_Symbol}")
                        # print(f"{ARM.STD_PE_Price_Org:.2f} : {ARM.STD_PE_Symbol}")
                        # print(f"{ARM.STD_CombPrice_Org:.2f} : Nifty Straddle Combine Price")
                        # print(f"{ARM.Combine_Per} % : Nifty Straddle Combine Stop-Loss")
                        # print("----------------------------------------------------")
                    else:pass
                else:
                    R2C8.set("Wait...... Inconsistant Strikes")
                    R2C8entry.config(fg="white",bg="red")
                    if ARM.ErrorExcptionInfo_scheduled is None: ARM.ErrorExcptionInfo_scheduled = root.after(1000,Clear_ErrorExcptionInfo)
            else:pass
            if ARM.Symbol_Changed:
                NF_FUT_SYMBOL()
                ARM.New_SYM_Tokens = {}
                ARM.New_FUT_Tokens = {}

                ARM.New_SYM_Tokens = {  int(ARM.SPT_Token): {'symbol': "NIFTY 50"} }
                # print(ARM.New_Subscribe_Tokens)
                # print(ARM.New_Subscribe_Tokens.keys())           #--Prints dict_keys([256265, 260105, 257801])
                # print(list(ARM.New_Subscribe_Tokens.keys()))     #--Prints [256265, 260105, 257801]
                ARM.New_Token_Ready = True
            elif ARM.Opt_Exp_Changed:
                ARM.New_Token_Ready = True
            else:pass
            # except Exception as e: #--Enable these three statements after checking
            #     print("Error Main",e)
        else:pass
        #============================================================================================================================

        try:
            if not kws_nse.is_connected():#--Check for kws connection lost
                if (R0C0.get() != "Logout"):
                    if ARM.Kws_ReconnAttempt < ARM.Kws_MaxReconAttempt:
                        #--Calculate the required delay (in terms of number of calls)
                        required_delay = 3 ** ARM.Kws_ReconnAttempt #--Try 2 or 3
                        if ARM.Kws_DelayCounter >= required_delay:
                            print("{}, Attempting to Re-Store Zerodha Connection (Attempt {})..."
                                    .format(time.strftime("%H:%M:%S", time.localtime()),ARM.Kws_ReconnAttempt + 1))
                            kws_nse.on_ticks = on_ticks_ARMKWS              #--Assign callback to on_ticks_ARMKWS
                            kws_nse.on_connect = on_connect_ARMKWS          #--Subscribe & Set Mode
                            kws_nse.connect(threaded=True)                  #--Restart WebSocket
                            ARM.Kws_ReconnAttempt += 1
                            ARM.Kws_DelayCounter = 0
                            ARM.KWS_Resterted = True
                        else:
                            ARM.Kws_DelayCounter += 1
                            #print(f"Waiting for {required_delay - ARM.Kws_DelayCounter} more calls before next reconnect attempt...")
                    else:
                        print("{}, Zerodha Server Not Responding.".format(time.strftime("%H:%M:%S", time.localtime())))
                        #------------------------------------------------------
                        ARM.ZLatency, ZJitter = measure_latency_and_jitter('www.zerodha.com')
                        GLatency, GJtter = measure_latency_and_jitter('www.google.com')

                        if ARM.ZLatency is not None and ZJitter is not None:#--Latency(RTT), Standard Deviation of the latency(RTT)
                            print(f"          Zerodha RTT : {ARM.ZLatency:.2f} ms, SD RTT: {ZJitter:.2f} ms")
                        else:
                            print("          Failed to measure Zerodha RTT & SD-RTT")

                        if GLatency is not None and GJtter is not None:#--Latency(RTT), Standard Deviation of the latency(RTT)
                            print(f"          Google  RTT : {GLatency:.2f} ms, SD RTT : {GJtter:.2f} ms")
                        else:
                            print("          Failed to measure Google RTT & SD-RTT")
                        #------------------------------------------------------
                        print("          Maximum Reconnection Attempts Reached (KWS-UND). Logging Out.")
                        R0C0.set("Logout")
                else:pass
            else:
                if ARM.KWS_URestart:
                    #kws_nse.stop()
                    kws_nse.close()                      #--Close Websocket
                    ARM.KWS_URestart = False
                    ARM.KWS_Functional = False
                    ARM.KWS_Check = 0
                    ARM.Kws_ReconnAttempt = 0
                    ARM.Kws_DelayCounter = 0
                    ARM.Delay_KWS_Restart = True
                    ARM.KWS_Closed = True
                    #print("--------- kws websocket Connection Closed for Restart ----------------")
                elif ARM.KWS_Closed and ARM.KWS_Resterted:
                    print("{}, Zerodha Connection Re-Stored Successfully".format(time.strftime("%H:%M:%S", time.localtime())))
                    R2C8.set("Kite Connection Re-Stored")
                    R2C8entry.config(fg="white",bg="green")
                    #if ARM.ErrorExcptionInfo_scheduled is None: ARM.ErrorExcptionInfo_scheduled = root.after(1000,Clear_ErrorExcptionInfo)
                    ARM.KWS_Closed = False
                    ARM.KWS_Resterted = False
                else:
                    # print("Assign callback to on_ticks_ARMKWS")
                    kws_nse.on_ticks = on_ticks_ARMKWS              #--Assign callback to on_ticks_ARMKWS
                ARM.Kws_ReconnAttempt = ARM.Kws_DelayCounter = 0    #--Reset attempts if connection is successful
        except:print("{}, Underlaying Data KWS Connection Error".format(time.strftime("%H:%M:%S", time.localtime())))
        # except ConnectionError as e:
        #     print(f"Connection error: {e}")
        # except TimeoutError as e:
        #     print(f"Timeout error: {e}")
        # except socket.error as e:
        #     print(f"Socket error: {e}")
        # except Exception as e:
        #     print(f"Unexpected error: {e}")
        #--------------------------------------------
        if ARM.Delay_KWS_Restart and ARM.ZLatency >= 40:    #--Use/check ARM.KWS_Resterted with Delay_KWS_Restart
            # print("Schedule APH_MAIN after 5 Sec")
            keep_scheduled = root.after(5000, APH_MAIN)
        elif ARM.New_Token_Ready:
            # print("Schedule APH_MAIN after 2 Sec")
            keep_scheduled = root.after(2000, APH_MAIN)
        else:
            keep_scheduled = root.after(5000 if ARM.ZLatency is None else 3000 if ARM.ZLatency >= 40 else 1000, APH_MAIN)
    else:
        #print("APH_MAIN Schedule Canceled")
        root.after_cancel(keep_scheduled) #--Cancel APH_MAIN() Schedule

def process_sell_trades():
    NewSell_Entry = len(ARM.Opt_Sold1_df)-1
    def call_process_sell_trades():
        Bought_Entries = len(ARM.Opt_Bought1_df)
        Diff_Qty = 0
        for i in range(Bought_Entries):
            # print("Sold Length:{},Bought Length:{}".format(NewSell_Entry,len(ARM.Opt_Bought1_df)))
            # print(ARM.Opt_Sold1_df.loc[NewSell_Entry,'Symbol'])
            # print(ARM.Opt_Bought1_df.loc[i,'Symbol'])
            if ARM.Opt_Bought1_df.loc[i,'Symbol'] == ARM.Opt_Sold1_df.loc[NewSell_Entry,'Symbol']:
                Diff_Qty = ARM.Opt_Bought1_df.loc[i,'Qty'] - ARM.Opt_Sold1_df.loc[NewSell_Entry,'Qty']
                if Diff_Qty == 0:
                    ARM.Opt_Bought1_df = ARM.Opt_Bought1_df.drop(ARM.Opt_Bought1_df.index[i])
                    ARM.Opt_Sold1_df = ARM.Opt_Sold1_df.drop(ARM.Opt_Sold1_df.index[NewSell_Entry])
                    ARM.Opt_Bought1_df = ARM.Opt_Bought1_df.reset_index(drop=True)
                    ARM.Opt_Sold1_df = ARM.Opt_Sold1_df.reset_index(drop=True)
                    #print("Diff_Qty == 0",Diff_Qty)
                    break
                elif Diff_Qty > 0:
                    ARM.Opt_Bought1_df.loc[i,'Qty'] = Diff_Qty
                    ARM.Opt_Sold1_df = ARM.Opt_Sold1_df.drop(ARM.Opt_Sold1_df.index[NewSell_Entry])
                    ARM.Opt_Bought1_df = ARM.Opt_Bought1_df.reset_index(drop=True)
                    ARM.Opt_Sold1_df = ARM.Opt_Sold1_df.reset_index(drop=True)
                    #print("Diff_Qty > 0",Diff_Qty)
                    break
                elif Diff_Qty < 0:
                    ARM.Opt_Sold1_df.loc[NewSell_Entry,'Qty'] = -Diff_Qty
                    ARM.Opt_Bought1_df = ARM.Opt_Bought1_df.drop(ARM.Opt_Bought1_df.index[i])
                    ARM.Opt_Bought1_df = ARM.Opt_Bought1_df.reset_index(drop=True)
                    ARM.Opt_Sold1_df = ARM.Opt_Sold1_df.reset_index(drop=True)
                    #print("Diff_Qty < 0",Diff_Qty)
                    break
                else:pass
            else:pass
        if Diff_Qty >= 0:return
        else: call_process_sell_trades()
    call_process_sell_trades()
    return
def process_buy_trades():
    NewBuy_Entry = len(ARM.Opt_Bought1_df)-1
    def call_process_buy_trades():
        Sold_Entries = len(ARM.Opt_Sold1_df)
        Diff_Qty = 0
        for i in range(Sold_Entries):
            # print("Bought Length:{}, Sold Length:{}".format(NewBuy_Entry,len(ARM.Opt_Sold1_df)))
            # print(ARM.Opt_Bought1_df.loc[NewBuy_Entry,'Symbol'])
            # print(ARM.Opt_Sold1_df.loc[i,'Symbol'])
            if ARM.Opt_Sold1_df.loc[i,'Symbol'] == ARM.Opt_Bought1_df.loc[NewBuy_Entry,'Symbol']:
                Diff_Qty = ARM.Opt_Sold1_df.loc[i,'Qty'] - ARM.Opt_Bought1_df.loc[NewBuy_Entry,'Qty']
                if Diff_Qty == 0:
                    ARM.Opt_Sold1_df = ARM.Opt_Sold1_df.drop(ARM.Opt_Sold1_df.index[i])
                    ARM.Opt_Bought1_df = ARM.Opt_Bought1_df.drop(ARM.Opt_Bought1_df.index[NewBuy_Entry])
                    ARM.Opt_Sold1_df = ARM.Opt_Sold1_df.reset_index(drop=True)
                    ARM.Opt_Bought1_df = ARM.Opt_Bought1_df.reset_index(drop=True)
                    #print("Diff_Qty == 0",Diff_Qty)
                    break
                elif Diff_Qty > 0:
                    ARM.Opt_Sold1_df.loc[i,'Qty'] = Diff_Qty
                    ARM.Opt_Bought1_df = ARM.Opt_Bought1_df.drop(ARM.Opt_Bought1_df.index[NewBuy_Entry])
                    ARM.Opt_Sold1_df = ARM.Opt_Sold1_df.reset_index(drop=True)
                    ARM.Opt_Bought1_df = ARM.Opt_Bought1_df.reset_index(drop=True)
                    #print("Diff_Qty > 0",Diff_Qty)
                    break
                elif Diff_Qty < 0:
                    ARM.Opt_Bought1_df.loc[NewBuy_Entry,'Qty'] = -Diff_Qty
                    ARM.Opt_Sold1_df = ARM.Opt_Sold1_df.drop(ARM.Opt_Sold1_df.index[i])
                    ARM.Opt_Sold1_df = ARM.Opt_Sold1_df.reset_index(drop=True)
                    ARM.Opt_Bought1_df = ARM.Opt_Bought1_df.reset_index(drop=True)
                    #print("Diff_Qty < 0",Diff_Qty)
                    break
                else:pass
            else:pass
        if Diff_Qty >= 0:return
        else: call_process_buy_trades()
    call_process_buy_trades()
    return
def Desplay_NetPosition():
    num_rows, num_columns = ARM.Display_NetBuySell_df.shape
    # print(f"Number of rows: {num_rows}")
    # print(f"Number of columns: {num_columns}")
    if (num_rows > 0):
        #ARM.Display_NetBuySell_df.to_csv("Net_Position.csv", index=False)
        #print("Update Status",ARM.Stop_Trade_Update)
        if ARM.Stop_Trade_Update and (GUI.R2C4flag == "RESET"):
            for i in range(28,32):
                for j in range(0,10):       #--range(2,14)
                    if j == 8:
                        OC_Cell_Value[i][j].set(-100.00)
                        OC_Cell[i][j].config(fg="#F0F0F0",bg="#F0F0F0")
                    elif j == 9:
                        OC_Cell_Value[i][j].set(100.00)
                        OC_Cell[i][j].config(fg="#F0F0F0",bg="#F0F0F0")
                    else:
                        OC_Cell[i][j].config(text="",fg="black",bg="#F0F0F0")
            OC_Cell[28][10].config(text="",fg="#F0F0F0",bg="#F0F0F0")
            OC_Cell[29][10].config(text="",fg="#F0F0F0",bg="#F0F0F0")
            OC_Cell[30][10].config(text="",fg="#F0F0F0",bg="#F0F0F0")
            OC_Cell[31][10].config(text="",fg="#F0F0F0",bg="#F0F0F0")
            OC_Cell[32][10].config(text="",fg="#F0F0F0",bg="#F0F0F0")
            # print("----------------------------------------------------------")
            # print(ARM.Opt_Bought1_df)
            # print("----------------------------------------------------------")
            # print(ARM.Opt_Bought2_df)
            print("----------------------------------------------------------")
            print(ARM.Display_NetBuySell_df)
            print("----------------------------------------------------------")
            ARM.Stop_Trade_Update = False
            GUI.R2C4flag = "RUN"
            R2C4.configure(fg="blue")
            ARM.Edit_SL_Target = False
            #print("Trade Display Refreshed......")
        else:pass

        for i in range(num_rows):
            if (ARM.Display_NetBuySell_df.loc[i,"Action"]=="BUY"):
                try:    #--For Bid Price
                    ARM.Display_NetBuySell_df.loc[i,'SqOffPrice'] = ARM.KWS_OPT_Data[ARM.Display_NetBuySell_df.loc[i,'Symbol']]['bid']
                    ARM.Display_NetBuySell_df.loc[i,'P/L(P)'] = round((ARM.Display_NetBuySell_df.loc[i,'SqOffPrice'] - ARM.Display_NetBuySell_df.loc[i,'Avg_Price']), ndigits = 2)
                    ARM.Display_NetBuySell_df.loc[i,'P/L(Rs)'] = round((ARM.Display_NetBuySell_df.loc[i,'P/L(P)'] * ARM.Display_NetBuySell_df.loc[i,'Total_Qty']), ndigits = 2)
                except:pass
            else:
                try:    #--For Ask Price
                    ARM.Display_NetBuySell_df.loc[i,'SqOffPrice'] = ARM.KWS_OPT_Data[ARM.Display_NetBuySell_df.loc[i,'Symbol']]['ask']
                    ARM.Display_NetBuySell_df.loc[i,'P/L(P)'] = round((ARM.Display_NetBuySell_df.loc[i,'Avg_Price'] - ARM.Display_NetBuySell_df.loc[i,'SqOffPrice']), ndigits = 2)
                    ARM.Display_NetBuySell_df.loc[i,'P/L(Rs)'] = round((ARM.Display_NetBuySell_df.loc[i,'P/L(P)'] * ARM.Display_NetBuySell_df.loc[i,'Total_Qty']), ndigits = 2)
                except:pass
        for i in range(28,num_rows+28): #--DISPLAY / UPDATE OPTION CHAIN
            # if (ARM.Display_NetBuySell_df.loc[i-28, "Action"]== "SELL"):
            #     OC_Cell[i][j].config(text=value,bg="#FFCCD5")
            # else:
            #     OC_Cell[i][j].config(text=value,bg="#BEF6BC")
            for j in range(0,num_columns):
                if (ARM.Display_NetBuySell_df.loc[i-28, "Action"]== "SELL"):
                    if j in [0,1,5]:
                        value = str(ARM.Display_NetBuySell_df.iloc[i-28, j])              #--2-2=0,3-2=1 (Action, Symbol)
                        OC_Cell[i][j].config(text=value,bg="#FFCCD5")
                    elif j in [2,4,6]:
                        value = "{:.2f}".format(ARM.Display_NetBuySell_df.iloc[i-28, j])  #--(Avg Price, SqOffPrice,P/L(P))
                        OC_Cell[i][j].config(text=value,bg="#FFCCD5")
                    elif j in [7]:
                        if (ARM.Display_NetBuySell_df.iloc[i-28, j] > 0):
                            value = "{:.2f}".format(ARM.Display_NetBuySell_df.iloc[i-28, j])  #--(P/L(Rs))
                            OC_Cell[i][j].config(text=value,fg="white",bg="green")
                        elif (ARM.Display_NetBuySell_df.iloc[i-28, j] < 0):
                            value = "{:.2f}".format(ARM.Display_NetBuySell_df.iloc[i-28, j])  #--(P/L(Rs))
                            OC_Cell[i][j].config(text=value,fg="white",bg="red")
                        else:
                            value = "{:.2f}".format(ARM.Display_NetBuySell_df.iloc[i-28, j])  #--(P/L(Rs))
                            OC_Cell[i][j].config(text=value,fg="black",bg="#FFCCD5")
                    elif j in [3]:
                        value = str(ARM.Display_NetBuySell_df.iloc[i-28, j])                  #--5-2=3 (Qty)
                        OC_Cell[i][j].config(text=value,bg="#FFCCD5")
                    elif j in [8,9]:
                        value = "{:.2f}".format(ARM.Display_NetBuySell_df.iloc[i-28, j])      #--9-2=7, 10-2=8
                        if not ARM.Edit_SL_Target:
                            OC_Cell_Value[i][j].set(value)
                            OC_Cell[i][j].config(fg="blue",bg="#FFCCD5")
                        else:pass
                    elif j in [10]:
                        value = str(ARM.Display_NetBuySell_df.iloc[i-28, j])                  #--(YES/NO)
                        OC_Cell[i][j].config(text=value,fg="blue",bg="#FFCCD5")
                    else:
                        OC_Cell[i][j].config(bg="#FFCCD5")
                else:
                    if j in [0,1,5]:
                        value = str(ARM.Display_NetBuySell_df.iloc[i-28, j])              #--2-2=0,3-2=1 (Action, Symbol)
                        OC_Cell[i][j].config(text=value,bg="#BEF6BC")
                    elif j in [2,4,6]:
                        value = "{:.2f}".format(ARM.Display_NetBuySell_df.iloc[i-28, j])  #--(Avg Price, SqOffPrice,P/L(P))
                        OC_Cell[i][j].config(text=value,bg="#BEF6BC")
                    elif j in [7]:
                        if (ARM.Display_NetBuySell_df.iloc[i-28, j] > 0):
                            value = "{:.2f}".format(ARM.Display_NetBuySell_df.iloc[i-28, j])  #--(P/L(Rs))
                            OC_Cell[i][j].config(text=value,fg="white",bg="green")
                        elif (ARM.Display_NetBuySell_df.iloc[i-28, j] < 0):
                            value = "{:.2f}".format(ARM.Display_NetBuySell_df.iloc[i-28, j])  #--(P/L(Rs))
                            OC_Cell[i][j].config(text=value,fg="white",bg="red")
                        else:
                            value = "{:.2f}".format(ARM.Display_NetBuySell_df.iloc[i-28, j])  #--(P/L(Rs))
                            OC_Cell[i][j].config(text=value,fg="black",bg="#BEF6BC")
                    elif (j == 3):
                        value = str(ARM.Display_NetBuySell_df.iloc[i-28, j])                  #--5-2=3 (Qty)
                        OC_Cell[i][j].config(text=value,bg="#BEF6BC")
                    elif (j == 8) or (j == 9):
                        value = "{:.2f}".format(ARM.Display_NetBuySell_df.iloc[i-28, j])      #--9-2=7, 10-2=8
                        if not ARM.Edit_SL_Target:
                            OC_Cell_Value[i][j].set(value)
                            OC_Cell[i][j].config(fg="blue",bg="#BEF6BC")
                        else:pass
                    elif (j == 10):
                        value = str(ARM.Display_NetBuySell_df.iloc[i-28, j])                  #--(YES/NO)
                        OC_Cell[i][j].config(text=value,fg="blue",bg="#BEF6BC")
                    else:pass

        if not ARM.Target_SL_Hit_Flag:
            for i in range(num_rows):
                if (ARM.Display_NetBuySell_df.loc[i,'P/L(P)'] > ARM.Display_NetBuySell_df.loc[i,'Target(P)']):
                    if (ARM.Display_NetBuySell_df.loc[i,'Action'] == 'BUY'):
                        ARM.Sell_Symbol_SqOff = ARM.Display_NetBuySell_df.loc[i,'Symbol']
                        ARM.oc_symbol_SqOff, ARM.Sell_Strike_SqOff, ARM.Sell_CE_PE_SqOff, ARM.Sell_Exp_SqOff = extract_info_from_symbol(ARM.Sell_Symbol_SqOff)
                        ARM.Sell_BidPrice_SqOff = ARM.Display_NetBuySell_df.loc[i,'SqOffPrice']
                        ARM.Sell_MIS_NRML_SqOff = GUI.R0C1flag
                        ARM.Sell_MARKET_LIMIT_SqOff = GUI.R1C1flag
                        ARM.Sell_Qty_SqOff = ARM.Display_NetBuySell_df.loc[i,'Total_Qty']
                        print("#----------------------------------------------------")
                        print(f"{ARM.Sell_Symbol_SqOff} - Target Hit, Auto Sell (Squiring Off) Bought Position")
                        print("#----------------------------------------------------")
                        ARM.Target_SL_Hit_Flag = True
                        #APH_SELL()
                        root.after(1,APH_SELL)
                        # try:
                        #     winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
                        # except:pass
                        R2C8.set("Square-Off TR/SL Hit")
                        R2C8entry.config(fg="white",bg="green")
                        if ARM.ErrorExcptionInfo_scheduled is None: ARM.ErrorExcptionInfo_scheduled = root.after(1000,Clear_ErrorExcptionInfo)
                        ARM.Desplay_NetPosition_flag = False
                        #print("Stop Display 1" )
                        return
                    elif (ARM.Display_NetBuySell_df.loc[i,'Action'] == 'SELL'):
                        ARM.Buy_Symbol_SqOff = ARM.Display_NetBuySell_df.loc[i,'Symbol']
                        ARM.oc_symbol_SqOff, ARM.Buy_Strike_SqOff, ARM.Buy_CE_PE_SqOff, ARM.Buy_Exp_SqOff = extract_info_from_symbol(ARM.Buy_Symbol_SqOff)
                        ARM.Buy_AskPrice_SqOff = ARM.Display_NetBuySell_df.loc[i,'SqOffPrice']
                        ARM.Buy_MIS_NRML_SqOff = GUI.R0C1flag
                        ARM.Buy_MARKET_LIMIT_SqOff = GUI.R1C1flag
                        ARM.Buy_Qty_SqOff = ARM.Display_NetBuySell_df.loc[i,'Total_Qty']
                        print("#----------------------------------------------------")
                        print(f"{ARM.Buy_Symbol_SqOff} - Target Hit, Auto Buy (Squiring Off) Sold Position")
                        print("#----------------------------------------------------")
                        ARM.Target_SL_Hit_Flag = True
                        #APH_BUY()
                        root.after(1,APH_BUY)
                        # try:
                        #     winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
                        # except:pass
                        R2C8.set("Square-Off TR/SL Hit")
                        R2C8entry.config(fg="white",bg="green")
                        if ARM.ErrorExcptionInfo_scheduled is None: ARM.ErrorExcptionInfo_scheduled = root.after(1000,Clear_ErrorExcptionInfo)
                        ARM.Desplay_NetPosition_flag = False
                        #print("Stop Display 2" )
                        return
                    else:pass
                elif (ARM.Display_NetBuySell_df.loc[i,'P/L(P)'] < ARM.Display_NetBuySell_df.loc[i,'SL(P)']):
                    if (ARM.Display_NetBuySell_df.loc[i,'Action'] == 'BUY'):
                        ARM.Sell_Symbol_SqOff = ARM.Display_NetBuySell_df.loc[i,'Symbol']
                        ARM.oc_symbol_SqOff, ARM.Sell_Strike_SqOff, ARM.Sell_CE_PE_SqOff, ARM.Sell_Exp_SqOff = extract_info_from_symbol(ARM.Sell_Symbol_SqOff)
                        ARM.Sell_BidPrice_SqOff = ARM.Display_NetBuySell_df.loc[i,'SqOffPrice']
                        ARM.Sell_MIS_NRML_SqOff = GUI.R0C1flag
                        ARM.Sell_MARKET_LIMIT_SqOff = GUI.R1C1flag
                        ARM.Sell_Qty_SqOff = ARM.Display_NetBuySell_df.loc[i,'Total_Qty']
                        print("#----------------------------------------------------")
                        print(f"{ARM.Sell_Symbol_SqOff} - SL Hit, Auto Sell (Squiring Off) Bought Position")
                        print("#----------------------------------------------------")
                        ARM.Target_SL_Hit_Flag = True
                        #APH_SELL()
                        root.after(1,APH_SELL)
                        # try:
                        #     winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
                        # except:pass
                        R2C8.set("Square-Off TR/SL Hit")
                        R2C8entry.config(fg="white",bg="green")
                        if ARM.ErrorExcptionInfo_scheduled is None: ARM.ErrorExcptionInfo_scheduled = root.after(1000,Clear_ErrorExcptionInfo)
                        ARM.Desplay_NetPosition_flag = False
                        #print("Stop Display 3" )
                        return
                    elif (ARM.Display_NetBuySell_df.loc[i,'Action'] == 'SELL'):
                        ARM.Buy_Symbol_SqOff = ARM.Display_NetBuySell_df.loc[i,'Symbol']
                        ARM.oc_symbol_SqOff, ARM.Buy_Strike_SqOff, ARM.Buy_CE_PE_SqOff, ARM.Buy_Exp_SqOff = extract_info_from_symbol(ARM.Buy_Symbol_SqOff)
                        ARM.Buy_AskPrice_SqOff = ARM.Display_NetBuySell_df.loc[i,'SqOffPrice']
                        ARM.Buy_MIS_NRML_SqOff = GUI.R0C1flag
                        ARM.Buy_MARKET_LIMIT_SqOff = GUI.R1C1flag
                        ARM.Buy_Qty_SqOff = ARM.Display_NetBuySell_df.loc[i,'Total_Qty']
                        print("#----------------------------------------------------")
                        print(f"{ARM.Buy_Symbol_SqOff} - SL Hit, Auto Buy (Squiring Off) Sold Position")
                        print("#----------------------------------------------------")
                        ARM.Target_SL_Hit_Flag = True
                        #APH_BUY()
                        root.after(1,APH_BUY)
                        # try:
                        #     winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
                        # except:pass
                        R2C8.set("Square-Off TR/SL Hit")
                        R2C8entry.config(fg="white",bg="green")
                        if ARM.ErrorExcptionInfo_scheduled is None: ARM.ErrorExcptionInfo_scheduled = root.after(1000,Clear_ErrorExcptionInfo)
                        ARM.Desplay_NetPosition_flag = False
                        #print("Stop Display 4" )
                        return
                    else:pass
                else:pass
        else:pass
        root.after(1000,Desplay_NetPosition)
    else:
        ARM.Desplay_NetPosition_flag = False
        #print("Stop Display 5" )
        return
def Auto_Buy_Status():
    if ARM.Buy_Symbol in ARM.Display_NetBuySell_df['Symbol'].values:
        #--Select specific columns
        columns_to_copy = ['Time_In','Action','Avg_Price','Total_Qty','SqOffPrice','Symbol','P/L(P)','P/L(Rs)']
        current_closed_trade = ARM.Display_NetBuySell_df.loc[ARM.Display_NetBuySell_df['Symbol'] == ARM.Buy_Symbol, columns_to_copy].copy()
        if current_closed_trade.at[current_closed_trade.index[0], 'Action'] == "SELL":
            print("AUTO Yes, Sold Earlier, now Buy (Square-Off)")            
        else:
            print("AUTO No, not Sold Earlier, Buy Again (Add Position)")
    else:
        print("AUTO No, not Sold Earlier, Fresh Buy (New Position)")
def APH_BUY():
    NW_PO_BuyError = False
    if (ARM.order_id_buy is None):        #--Fresh New Order First Time
        if ARM.Target_SL_Hit_Flag:
            ARM.Buy_AskPrice = ARM.Buy_AskPrice_SqOff
            ARM.Buy_Symbol = ARM.Buy_Symbol_SqOff
            ARM.Buy_Qty = ARM.Buy_Qty_SqOff
            ARM.Sell_MIS_NRML = ARM.Sell_MIS_NRML_SqOff
            ARM.Sell_MARKET_LIMIT = ARM.Sell_MARKET_LIMIT_SqOff
        else:pass
        if (ARM.Buy_MARKET_LIMIT == "LIMIT"):
            try:
                position_buy_at = (int((ARM.Buy_AskPrice + (ARM.Buy_AskPrice/(100/ARM.BuySellSpread)))*100)//5*5)/100.0 #--Try to buy 2% More than market price
                if (GUI.R0C5flag == "REAL"):
                        ARM.order_id_buy = kite.place_order(variety=kite.VARIETY_REGULAR,
                                exchange=kite.EXCHANGE_NFO,
                                tradingsymbol=ARM.Buy_Symbol,
                                transaction_type="BUY",
                                quantity=int(ARM.Buy_Qty),
                                product=ARM.Buy_MIS_NRML,
                                order_type=kite.ORDER_TYPE_LIMIT,
                                price=position_buy_at,
                                validity=None,
                                disclosed_quantity=None,
                                trigger_price=None,
                                squareoff=None,
                                stoploss=None,
                                trailing_stoploss=None,
                                tag=None)
                else:
                    ARM.order_id_buy = 2888888888
                ARM.Order_WaitTime = 10
            except Exception as e:
                NW_PO_BuyError = True
                print("Technical error wile Buying, Check Zerodha / Internet")
            if not NW_PO_BuyError:
                ARM.price_try = position_buy_at
                print("{} Buy Order Tried at Rs.{}".format(ARM.Buy_Symbol,position_buy_at))
                time.sleep(0.5)
        else:pass
    else:pass
    if not NW_PO_BuyError:
        myrejectedorder_status = ARM.myopenorder_status = mycancleorder_status = mycompleteorder_status = False
        if (ARM.MY_TRADE == "REAL"):                    #--REAL Mode Trade
            ARM.Zerodha_Orders_Flag = False
            try:
                ARM.My_DayOrders = kite.orders()
            except Exception as e:
                ARM.Zerodha_Orders_Flag = True
                if not ARM.ZDOrder_Info:
                    ARM.ZDOrder_Info = True
                    print("{}, Not getting Order Status/Information from Zerodha".format(time.strftime("%H:%M:%S", time.localtime())))
                    if ARM.Buy_CE_PE == "CE":
                        R26C4["text"] = f"Wait"
                        R26C4["bg"] = "yellow"
                    else:
                        R26C6["text"] = f"Wait"
                        R26C6["bg"] = "yellow"
                else:pass
            if not ARM.Zerodha_Orders_Flag:
                ARM.ZDOrder_Info = False
                for individual_order in ARM.My_DayOrders:
                    if int(individual_order['order_id']) == int(ARM.order_id_buy):
                        myrejectedorder_status = (individual_order['status'] == "REJECTED")
                        ARM.myopenorder_status = (individual_order['status'] == "OPEN")
                        mycancleorder_status = (individual_order['status'] == "CANCELLED")
                        if(individual_order['status'] == "COMPLETE"):
                            mycompleteorder_status = True
                            ARM.BoughtPrice = round(float(individual_order['average_price']),ndigits = 2)
                        else:
                            mycompleteorder_status = False
                        break
                    else:pass
            else:
                mycancleorder_status = False
                ARM.myopenorder_status = False
                myrejectedorder_status = False
                mycompleteorder_status = False
        else:
            myrejectedorder_status = False              #--PAPER Mode Trade
            if (ARM.price_try >= ARM.Buy_AskPrice):
                ARM.myopenorder_status = False          #--COMPLETE
                mycompleteorder_status = True
                ARM.BoughtPrice = ARM.Buy_AskPrice
            else:
                ARM.myopenorder_status = True           #--Keep OPEN
                mycompleteorder_status = False
            mycancleorder_status = False
        if (mycancleorder_status):
            ARM.price_try = 0.0
            ARM.order_id_buy = None
            ARM.myopenorder_status = False
            print("{0}, NW-PO {1} OPEN Order Cancelled in 2nd attempt"
                    .format(time.strftime("%H:%M:%S", time.localtime()),ARM.Buy_Symbol))
            if ARM.Buy_CE_PE == "CE":
                GUI.R3C4_Value = 1
                R3C4.set("{:.0f}".format(GUI.R3C4_Value))
                R26C4["text"] = f"CALCELLED"
                R26C4["bg"] = "yellow"
            else:
                GUI.R3C6_Value = 1
                R3C6.set("{:.0f}".format(GUI.R3C6_Value))
                R26C6["text"] = f"CALCELLED"
                R26C6["bg"] = "yellow"
            ARM.Buy_CE_PE = ""
            ARM.Buy_AskPrice = 0
            ARM.Buy_Strike = 0
            ARM.Buy_Qty = 0
            if ARM.Target_SL_Hit_Flag: ARM.Target_SL_Hit_Flag=False
        elif (ARM.myopenorder_status):
            if (ARM.Order_WaitTime >= 0):#--Wait for 5 Seconds
                if (ARM.Order_WaitTime == 10):print("{0}, NW-PO {1} Order is OPEN, (Qty-{2}), Wait for {3} Sec"#--5 4 3 2 1 sec in every iteration
                                                .format(time.strftime("%H:%M:%S", time.localtime()),ARM.Buy_Symbol,ARM.Buy_Qty,ARM.Order_WaitTime))
                if ARM.Buy_CE_PE == "CE":
                    R26C4["text"] = f"Wait ({ARM.Order_WaitTime} s)"
                    R26C4["bg"] = "yellow"
                else:
                    R26C6["text"] = f"Wait ({ARM.Order_WaitTime} s)"
                    R26C6["bg"] = "yellow"
                ARM.Order_WaitTime = ARM.Order_WaitTime - 1
            else:
                print("{0}, NW-PO {1} Could't get Buy Price, Canceling Order"
                            .format(time.strftime("%H:%M:%S", time.localtime()),ARM.Buy_Symbol))
                CANTechnicalError = False
                try:
                    if (ARM.MY_TRADE == "REAL"):        #--REAL Mode Trade
                        ARM.order_id_buy = kite.cancel_order(variety=kite.VARIETY_REGULAR,
                                                        order_id = ARM.order_id_buy,
                                                        parent_order_id=None)
                    else:pass
                except Exception as e:
                    CANTechnicalError = True
                    print("{0}, NW-PO {1} Buy Order Cancellation failed (TechnicalError)"
                                .format(time.strftime("%H:%M:%S", time.localtime()),ARM.Buy_Symbol))

                if (not CANTechnicalError):
                    mycancleorder_status = False
                    if (ARM.MY_TRADE == "REAL"):        #--REAL Mode Trade
                        time.sleep(0.5)
                        ARM.Zerodha_Orders_Flag = False
                        try:
                            ARM.My_DayOrders = kite.orders()
                        except Exception as e:
                            ARM.Zerodha_Orders_Flag = True
                        if not ARM.Zerodha_Orders_Flag:
                            for individual_order in ARM.My_DayOrders:
                                if int(individual_order['order_id']) == int(ARM.order_id_buy):
                                    mycancleorder_status = (individual_order['status'] == "CANCELLED")
                                    break
                                else:pass
                        else:
                            mycancleorder_status = False
                            ARM.myopenorder_status = False
                            myrejectedorder_status = False
                            mycompleteorder_status = False
                    else:
                        mycancleorder_status = True     #--PAPER Mode Trade
                    if mycancleorder_status:
                        ARM.price_try = 0.0
                        ARM.order_id_buy = None
                        ARM.myopenorder_status = False
                        print("{0}, NW-PO {1} OPEN Order Cancelled in 1st attempt"
                                .format(time.strftime("%H:%M:%S", time.localtime()),ARM.Buy_Symbol))
                        if ARM.Buy_CE_PE == "CE":
                            GUI.R3C4_Value = 1
                            R3C4.set("{:.0f}".format(GUI.R3C4_Value))
                            R26C4["text"] = f"CALCELLED"
                            R26C4["bg"] = "yellow"
                        else:
                            GUI.R3C6_Value = 1
                            R3C6.set("{:.0f}".format(GUI.R3C6_Value))
                            R26C6["text"] = f"CALCELLED"
                            R26C6["bg"] = "yellow"
                        ARM.Buy_CE_PE = ""
                        ARM.Buy_AskPrice = 0
                        ARM.Buy_Strike = 0
                        ARM.Buy_Qty = 0
                        if ARM.Target_SL_Hit_Flag: ARM.Target_SL_Hit_Flag=False
                    else:pass
                else:pass
        elif (myrejectedorder_status):
            ARM.price_try = 0.0
            ARM.order_id_buy = None
            ARM.myopenorder_status = False
            print("{0}, NW-PO {1} Order Placement Rejected"
                            .format(time.strftime("%H:%M:%S", time.localtime()),ARM.Buy_Symbol))
            if ARM.Buy_CE_PE == "CE":
                GUI.R3C4_Value = 1
                R3C4.set("{:.0f}".format(GUI.R3C4_Value))
                R26C4["text"] = f"REJECTED"
                R26C4["bg"] = "yellow"
            else:
                GUI.R3C6_Value = 1
                R3C6.set("{:.0f}".format(GUI.R3C6_Value))
                R26C6["text"] = f"REJECTED"
                R26C6["bg"] = "yellow"
            ARM.Buy_CE_PE = ""
            ARM.Buy_AskPrice = 0
            ARM.Buy_Strike = 0
            ARM.Buy_Qty = 0
            if ARM.Target_SL_Hit_Flag: ARM.Target_SL_Hit_Flag=False
        #----------------------------------------
        elif (mycompleteorder_status):
            ARM.price_try = 0.0
            ARM.order_id_buy = None
            ARM.myopenorder_status = False
            print("Buy Order Completed at Rs.{} Check Zerodha Account".format(ARM.BoughtPrice))
            #--------------------------------------------------------------------------------------------
            if ARM.Buy_Symbol in ARM.Display_NetBuySell_df['Symbol'].values:
                #--Select specific columns
                columns_to_copy = ['Time_In','Action','Avg_Price','Total_Qty','SqOffPrice','Symbol','P/L(P)','P/L(Rs)']
                current_closed_trade = ARM.Display_NetBuySell_df.loc[ARM.Display_NetBuySell_df['Symbol'] == ARM.Buy_Symbol, columns_to_copy].copy()
                if current_closed_trade.at[current_closed_trade.index[0], 'Action'] == "SELL":
                    # print("Yes, Sold Earlier, now Buy (Square-Off)")
                    #--Update cell values
                    current_closed_trade.at[current_closed_trade.index[0], 'Total_Qty'] = ARM.Buy_Qty
                    current_closed_trade.at[current_closed_trade.index[0], 'SqOffPrice'] = ARM.BoughtPrice
                    current_closed_trade.at[current_closed_trade.index[0], 'P/L(P)'] = round((current_closed_trade.iloc[0]['Avg_Price'] - ARM.BoughtPrice), 2)
                    current_closed_trade.at[current_closed_trade.index[0], 'P/L(Rs)'] = round((current_closed_trade.iloc[0]['P/L(P)'] * ARM.Buy_Qty), 2)

                    #--Add Time Column
                    current_time = datetime.now().strftime('%H:%M:%S')
                    current_closed_trade.insert(0, 'Time_Out', current_time)

                    #--Display and Append
                    #print(current_closed_trade)
                    ARM.My_AllTrades = pd.concat([ARM.My_AllTrades, current_closed_trade], ignore_index=True)
                else:pass
                    # print("No, not Sold Earlier, Buy Again (Add Position)")
            else:pass
                # print("No, not Sold Earlier, Fresh Buy (New Position)")
            #ARM.My_AllTrades = ARM.Display_NetBuySell_df.copy()
            #--------------------------------------------------------------------------------------------
            if ARM.Target_SL_Hit_Flag:
                key_to_search = ("BUY", ARM.oc_symbol_SqOff, ARM.Buy_Strike_SqOff, ARM.Buy_CE_PE_SqOff, ARM.Buy_Exp_SqOff)
            else:
                key_to_search = ("BUY", ARM.oc_symbol, ARM.Buy_Strike, ARM.Buy_CE_PE, ARM.oc_opt_expiry)
            if ARM.Option_Bought:
                key_to_update = None
                for key in ARM.Option_Bought:
                    if key == key_to_search:
                        key_to_update = key
                        break
                if key_to_update is not None:
                    ARM.Option_Bought[key_to_update][5] += ARM.Buy_Qty
                else:
                    ARM.Option_Bought[key_to_search] = list(key_to_search) + [ARM.Buy_Qty]
            else:
                ARM.Option_Bought[key_to_search] = list(key_to_search) + [ARM.Buy_Qty]
            #--------------------------------------------------------------------------------------------
            #ARM.BoughtPrice = ARM.Buy_AskPrice #--Price Tried
            ARM.TotalBuyOrders += 1
            Time_In = datetime.now().strftime('%H:%M:%S')

            LPT_Pressed()
            PPT_Pressed()
            StopLossP = round(GUI.LPT_Value/ARM.found_lotsize_opt,ndigits = 2)
            TargetP = round(GUI.PPT_Value/ARM.found_lotsize_opt,ndigits = 2)
            Option_Buy_Info = (ARM.TotalBuyOrders,Time_In,'BUY',ARM.BoughtPrice,ARM.Buy_Qty,ARM.Buy_SqOffPrice,
                        ARM.Buy_Symbol,ARM.Buy_PL_P,ARM.Buy_PL_Rs,-StopLossP,TargetP,"NO")

            # Option_Buy_Info = (ARM.TotalBuyOrders,'BUY',ARM.Buy_Symbol,ARM.BoughtPrice,ARM.Buy_Qty,
            #             ARM.Buy_SqOffPrice,ARM.Buy_PL_P,ARM.Buy_PL_Rs,-100.00,100.00,"NO")

            if ARM.NetBuySell_len > 0:
                for i in range(ARM.NetBuySell_len):
                    for j in range(len(ARM.Opt_Bought1_df)):
                        if ((ARM.Display_NetBuySell_df.loc[i,'Action'] == ARM.Opt_Bought1_df.loc[j,'Action']) and
                            (ARM.Display_NetBuySell_df.loc[i,'Symbol'] == ARM.Opt_Bought1_df.loc[j,'Symbol'])):
                            ARM.Opt_Bought1_df.loc[j,'SL(P)'] = ARM.Display_NetBuySell_df.loc[i,'SL(P)']
                            ARM.Opt_Bought1_df.loc[j,'Target(P)'] = ARM.Display_NetBuySell_df.loc[i,'Target(P)']
                        else:pass
                    if ((ARM.Display_NetBuySell_df.loc[i,'Action'] == "BUY") and
                        (ARM.Display_NetBuySell_df.loc[i,'Symbol'] == ARM.Buy_Symbol)):
                        Option_Buy_Info = (ARM.TotalBuyOrders,ARM.Display_NetBuySell_df.loc[i,'Time_In'],'BUY',ARM.BoughtPrice,ARM.Buy_Qty,ARM.Buy_SqOffPrice,
                                          ARM.Buy_Symbol,ARM.Buy_PL_P,ARM.Buy_PL_Rs,
                                        ARM.Display_NetBuySell_df.loc[i,'SL(P)'],ARM.Display_NetBuySell_df.loc[i,'Target(P)'],"NO")
                    else:pass
            else:pass
            ARM.Opt_Bought1_df.loc[len(ARM.Opt_Bought1_df)] = Option_Buy_Info
            # print("---------- Bought Option Contract ----------")
            # print(ARM.Opt_Bought1_df)
            # print("---------- Check above ----------")
            if(len(ARM.Opt_Sold1_df)>0):
                #print("---------- Checking Sell Positions in Data ----------")
                # print(ARM.Opt_Bought1_df)
                # print(ARM.Opt_Sold1_df)
                # print("----")
                process_buy_trades()
                # print(ARM.Opt_Bought1_df)
                # print(ARM.Opt_Sold1_df)
                # print("----")
                #--------------------------------------------------------------------------------------------
                #--Combines Similar Sell Clicks which has same Symbol,Strike,Type & Price
                ARM.Opt_Sold2_df = ARM.Opt_Sold1_df.groupby(['Action','Symbol','Price','Qty'], as_index=False).agg({'Qty':'sum', 'OrderNo':'first',
                                        'Time_In':'first','SqOffPrice':'first','P/L(P)':'first','P/L(Rs)':'first','SL(P)':'first','Target(P)':'first','SQOff':'first'})
                ARM.Opt_Sold2_df = ARM.Opt_Sold2_df.sort_values(by=['Symbol','OrderNo'])
                ARM.Opt_Sold2_df = ARM.Opt_Sold2_df[['OrderNo','Time_In','Action','Price','Qty','SqOffPrice','Symbol','P/L(P)','P/L(Rs)','SL(P)','Target(P)','SQOff']]
                ARM.Opt_Sold2_df = ARM.Opt_Sold2_df.reset_index(drop=True)
                #print(ARM.Opt_Sold2_df)
                #--------------------------------------------------------------------------------------------
                #--Combines Similar Sell Clicks which has same Symbol,Strike,Type, has Avg_Price and Total_Qty
                ARM.Opt_Sold3_df = ARM.Opt_Sold2_df.copy()
                ARM.Opt_Sold3_df = ARM.Opt_Sold3_df.drop(columns=['OrderNo'])
                ARM.Opt_Sold3_df['Total_Qty'] = ARM.Opt_Sold3_df.groupby(['Symbol'])['Qty'].transform('sum')
                weighted_sum = ARM.Opt_Sold3_df['Price'] * ARM.Opt_Sold3_df['Qty']
                grouped_sum = weighted_sum.groupby([ARM.Opt_Sold3_df['Symbol']]).transform('sum')
                # ARM.Opt_Sold3_df['Avg_Price'] = grouped_sum / ARM.Opt_Sold3_df.groupby(['Symbol'])['Qty'].transform('sum')
                ARM.Opt_Sold3_df['Avg_Price'] = (grouped_sum / ARM.Opt_Sold3_df.groupby(['Symbol'])['Qty'].transform('sum')).round(2)
                mask = ARM.Opt_Sold3_df.duplicated(['Symbol'], keep=False)
                ARM.Opt_Sold3_df.loc[mask, 'Qty'] = ARM.Opt_Sold3_df.loc[mask, 'Total_Qty']
                ARM.Opt_Sold3_df.loc[mask, 'Price'] = ARM.Opt_Sold3_df.loc[mask, 'Avg_Price']    #--FutureWarning, so tried below statement
                ARM.Opt_Sold3_df['Price'] = ARM.Opt_Sold3_df['Price'].astype('float64')
                ARM.Opt_Sold3_df = ARM.Opt_Sold3_df[['Time_In','Action','Avg_Price','Total_Qty','SqOffPrice','Symbol','P/L(P)','P/L(Rs)','SL(P)','Target(P)','SQOff']]
                #ARM.Opt_Sold3_df = ARM.Opt_Sold3_df.drop_duplicates().reset_index(drop=True)
                ARM.Opt_Sold3_df = ARM.Opt_Sold3_df.drop_duplicates(subset=['Action', 'Symbol', 'Avg_Price']).reset_index(drop=True)
                #print(ARM.Opt_Sold3_df)
                #--------------------------------------------------------------------------------------------
            else:pass
            #--------------------------------------------------------------------------------------------
            #--Combines Similar Buy Clicks which has same Symbol,Strike,Type & Price
            ARM.Opt_Bought2_df = ARM.Opt_Bought1_df.groupby(['Action','Symbol','Price','Qty'], as_index=False).agg({'Qty':'sum', 'OrderNo':'first',
                                        'Time_In':'first','SqOffPrice':'first','P/L(P)':'first','P/L(Rs)':'first','SL(P)':'first','Target(P)':'first','SQOff':'first'})
            ARM.Opt_Bought2_df = ARM.Opt_Bought2_df.sort_values(by=['Symbol','OrderNo'])
            ARM.Opt_Bought2_df = ARM.Opt_Bought2_df[['OrderNo','Time_In','Action','Price','Qty','SqOffPrice','Symbol','P/L(P)','P/L(Rs)','SL(P)','Target(P)','SQOff']]
            ARM.Opt_Bought2_df = ARM.Opt_Bought2_df.reset_index(drop=True)
            #print(ARM.Opt_Bought2_df)
            #--------------------------------------------------------------------------------------------
            #--Combines Similar Buy Clicks which has same Symbol,Strike,Type, has Avg_Price and Total_Qty
            ARM.Opt_Bought3_df = ARM.Opt_Bought2_df.copy()
            ARM.Opt_Bought3_df = ARM.Opt_Bought3_df.drop(columns=['OrderNo'])
            ARM.Opt_Bought3_df['Total_Qty'] = ARM.Opt_Bought3_df.groupby(['Symbol'])['Qty'].transform('sum')
            weighted_sum = ARM.Opt_Bought3_df['Price'] * ARM.Opt_Bought3_df['Qty']
            grouped_sum = weighted_sum.groupby([ARM.Opt_Bought3_df['Symbol']]).transform('sum')
            ARM.Opt_Bought3_df['Avg_Price'] = (grouped_sum / ARM.Opt_Bought3_df.groupby(['Symbol'])['Qty'].transform('sum')).round(2)
            mask = ARM.Opt_Bought3_df.duplicated(['Symbol'], keep=False)
            ARM.Opt_Bought3_df.loc[mask, 'Qty'] = ARM.Opt_Bought3_df.loc[mask, 'Total_Qty']
            ARM.Opt_Bought3_df.loc[mask, 'Price'] = ARM.Opt_Bought3_df.loc[mask, 'Avg_Price'] #--FutureWarning, so tried below statement
            ARM.Opt_Bought3_df['Price'] = ARM.Opt_Bought3_df['Price'].astype('float64')
            ARM.Opt_Bought3_df = ARM.Opt_Bought3_df[['Time_In','Action','Avg_Price','Total_Qty','SqOffPrice','Symbol','P/L(P)','P/L(Rs)','SL(P)','Target(P)','SQOff']]
            #ARM.Opt_Bought3_df = ARM.Opt_Bought3_df.drop_duplicates().reset_index(drop=True)
            ARM.Opt_Bought3_df = ARM.Opt_Bought3_df.drop_duplicates(subset=['Action', 'Symbol', 'Avg_Price']).reset_index(drop=True)
            #print(ARM.Opt_Bought3_df)
            #--------------------------------------------------------------------------------------------
            ARM.Display_NetBuySell_df = pd.concat([ARM.Opt_Bought3_df, ARM.Opt_Sold3_df], ignore_index=True)
            ARM.NetBuySell_len = len(ARM.Display_NetBuySell_df)
            #--------------------------------------------------------------------------------------------
            # print(ARM.Option_Bought)
            # print(ARM.Option_Sold)
            # print(ARM.Option_NetBuySell)
            for buy_key, buy_values in ARM.Option_Bought.items():
                sell_key = ('SELL',) + buy_key[1:]
                if sell_key not in ARM.Option_Sold:
                    ARM.Option_NetBuySell[buy_key] = buy_values
                else:pass
            for sell_key, sell_values in ARM.Option_Sold.items():
                buy_key = ('BUY',) + sell_key[1:]
                if buy_key not in ARM.Option_Bought:
                    ARM.Option_NetBuySell[sell_key] = sell_values
                else:pass
            NetBuySell_modified = False
            for buy_key, buy_values in ARM.Option_Bought.items():
                for sell_key, sell_values in ARM.Option_Sold.items():
                    if buy_key[1:] == sell_key[1:]:
                        quantity = buy_values[5] - sell_values[5]
                        if quantity > 0:
                            ARM.Option_Bought[buy_key] = buy_values[:5] + [quantity]
                            ARM.Option_Sold.pop(sell_key, None)
                            ARM.Option_NetBuySell.pop(sell_key, None)
                            ARM.Option_NetBuySell[buy_key] = buy_values[:5] + [quantity]
                            NetBuySell_modified = True
                            break
                        elif quantity < 0:
                            quantity = abs(quantity)
                            ARM.Option_Bought.pop(buy_key, None)
                            ARM.Option_Sold[sell_key] = sell_values[:5] + [quantity]
                            ARM.Option_NetBuySell.pop(buy_key, None)
                            ARM.Option_NetBuySell[sell_key] = sell_values[:5] + [quantity]
                            NetBuySell_modified = True
                            break
                        elif quantity == 0:
                            ARM.Option_Bought.pop(buy_key, None)
                            ARM.Option_Sold.pop(sell_key, None)
                            ARM.Option_NetBuySell.pop(buy_key, None)
                            ARM.Option_NetBuySell.pop(sell_key, None)
                            NetBuySell_modified = True
                            break
                        else:pass
                    else:pass
                if NetBuySell_modified:break
            # print(ARM.Option_Bought)
            # print(ARM.Option_Sold)
            # print(ARM.Option_NetBuySell)
            #--------------------------------------------------------------------------------------------
            # print("Trades in Display_NetBuySell_df: ",ARM.NetBuySell_len)
            # print(ARM.Display_NetBuySell_df)
            if (ARM.NetBuySell_len > 0):
                if not ARM.Desplay_NetPosition_flag :
                    ARM.Desplay_NetPosition_flag = True
                    Desplay_NetPosition()
                else:pass
                num_rows, num_columns = ARM.Display_NetBuySell_df.shape
                for i in range(28,ARM.pev_num_rows+28): #--Clear Previous
                    for j in range(0,ARM.prev_num_columns):
                        if j == 8:
                            # print(f"SL Row No {i}")
                            OC_Cell_Value[i][j].set(-100.00)
                            OC_Cell[i][j].config(fg="#F0F0F0",bg="#F0F0F0")
                        elif j == 9:
                            # print(f"Target Row No {i}")
                            OC_Cell_Value[i][j].set(100.00)
                            OC_Cell[i][j].config(fg="#F0F0F0",bg="#F0F0F0")
                        else:
                            OC_Cell[i][j].config(text="",fg="black",bg="#F0F0F0")
                    #OC_Cell[i][15].config(text="",fg="#F0F0F0",bg="#F0F0F0")

                OC_Cell[28][10].config(text="",fg="#F0F0F0",bg="#F0F0F0")
                OC_Cell[29][10].config(text="",fg="#F0F0F0",bg="#F0F0F0")
                OC_Cell[30][10].config(text="",fg="#F0F0F0",bg="#F0F0F0")
                OC_Cell[31][10].config(text="",fg="#F0F0F0",bg="#F0F0F0")
                OC_Cell[32][10].config(text="",fg="#F0F0F0",bg="#F0F0F0")
                ARM.pev_num_rows = num_rows
                ARM.prev_num_columns = num_columns
            else:
                for i in range(28,32): #--Clear Previous
                    for j in range(0,10):
                        if j == 8:
                            OC_Cell_Value[i][j].set(-100.00)
                            OC_Cell[i][j].config(fg="#F0F0F0",bg="#F0F0F0")
                        elif j == 9:
                            OC_Cell_Value[i][j].set(100.00)
                            OC_Cell[i][j].config(fg="#F0F0F0",bg="#F0F0F0")
                        else:
                            OC_Cell[i][j].config(text="",fg="black",bg="#F0F0F0")
                OC_Cell[28][10].config(text="",fg="#F0F0F0",bg="#F0F0F0")
                OC_Cell[29][10].config(text="",fg="#F0F0F0",bg="#F0F0F0")
                OC_Cell[30][10].config(text="",fg="#F0F0F0",bg="#F0F0F0")
                OC_Cell[31][10].config(text="",fg="#F0F0F0",bg="#F0F0F0")
                OC_Cell[32][10].config(text="",fg="#F0F0F0",bg="#F0F0F0")
                ARM.Desplay_NetPosition_flag = False
                #print("Stop Display 6" )
            #--------------------------------------------------------------------------------------------
            #print("Clearing Buy Veriables")
            if ARM.Buy_CE_PE == "CE":
                GUI.R3C4_Value = 1
                R3C4.set("{:.0f}".format(GUI.R3C4_Value))
            else:
                GUI.R3C6_Value = 1
                R3C6.set("{:.0f}".format(GUI.R3C6_Value))
            ARM.Buy_CE_PE = ""
            ARM.Buy_AskPrice = 0
            ARM.Buy_Strike = 0
            ARM.Buy_Qty = 0
            if ARM.Target_SL_Hit_Flag: ARM.Target_SL_Hit_Flag=False
            #--------------------------------------------------------------------------------------------
            #print("Ending Net Option Buy Sell")
        else:pass
    if ARM.myopenorder_status:root.after(1000,APH_BUY)
def APH_SELL():
    NW_PO_SellError = False
    if (ARM.order_id_sell is None):        #--Fresh New Order First Time
        if ARM.Target_SL_Hit_Flag:
            ARM.Sell_BidPrice = ARM.Sell_BidPrice_SqOff
            ARM.Sell_Symbol = ARM.Sell_Symbol_SqOff
            ARM.Sell_Qty = ARM.Sell_Qty_SqOff
            ARM.Sell_MIS_NRML = ARM.Sell_MIS_NRML_SqOff
            ARM.Sell_MARKET_LIMIT = ARM.Sell_MARKET_LIMIT_SqOff
        else:pass
        if (ARM.Sell_MARKET_LIMIT == "LIMIT"):
            try:
                position_sell_at = (int((ARM.Sell_BidPrice - (ARM.Sell_BidPrice/(100/ARM.BuySellSpread)))*100)//5*5)/100.0 #--2% Less than market price
                if (GUI.R0C5flag == "REAL"):
                        ARM.order_id_sell = kite.place_order(variety=kite.VARIETY_REGULAR,
                                exchange=kite.EXCHANGE_NFO,
                                tradingsymbol=ARM.Sell_Symbol,
                                transaction_type="SELL",
                                quantity=int(ARM.Sell_Qty),
                                product=ARM.Sell_MIS_NRML,
                                order_type=kite.ORDER_TYPE_LIMIT,
                                price=position_sell_at,
                                validity=None,
                                disclosed_quantity=None,
                                trigger_price=None,
                                squareoff=None,
                                stoploss=None,
                                trailing_stoploss=None,
                                tag=None)
                else:
                    ARM.order_id_sell = 2111111111
                ARM.Order_WaitTime = 10
            except Exception as e:
                NW_PO_SellError = True
                print("Technical error wile Selling, Check Zerodha / Internet")
            if not NW_PO_SellError:
                ARM.price_try = position_sell_at
                print("Sell Order Tried at Rs.{}".format(position_sell_at))
                time.sleep(0.5)
        else:pass
    else:pass
    if not NW_PO_SellError:
        myrejectedorder_status = ARM.myopenorder_status = mycancleorder_status = mycompleteorder_status = False
        if (ARM.MY_TRADE == "REAL"):                    #--REAL Mode Trade
            ARM.Zerodha_Orders_Flag = False
            try:
                ARM.My_DayOrders = kite.orders()
            except Exception as e:
                ARM.Zerodha_Orders_Flag = True
                if not ARM.ZDOrder_Info:
                    ARM.ZDOrder_Info = True
                    print("{}, Not getting Order Status/Information from Zerodha".format(time.strftime("%H:%M:%S", time.localtime())))
                    if ARM.Sell_CE_PE == "CE":
                        R26C2["text"] = f"Wait"
                        R26C2["bg"] = "yellow"
                    else:
                        R26C8["text"] = f"Wait"
                        R26C8["bg"] = "yellow"
                else:pass
            if not ARM.Zerodha_Orders_Flag:
                ARM.ZDOrder_Info = False
                for individual_order in ARM.My_DayOrders:
                    if int(individual_order['order_id']) == int(ARM.order_id_sell):
                        myrejectedorder_status = (individual_order['status'] == "REJECTED")
                        ARM.myopenorder_status = (individual_order['status'] == "OPEN")
                        mycancleorder_status = (individual_order['status'] == "CANCELLED")
                        if(individual_order['status'] == "COMPLETE"):
                            mycompleteorder_status = True
                            ARM.SoldPrice = round(float(individual_order['average_price']),ndigits = 2)
                        else:
                            mycompleteorder_status = False
                        break
                    else:pass
            else:
                mycancleorder_status = False
                ARM.myopenorder_status = False
                myrejectedorder_status = False
                mycompleteorder_status = False
        else:
            myrejectedorder_status = False              #--PAPER Mode Trade
            if (ARM.price_try <= ARM.Sell_BidPrice):
                ARM.myopenorder_status = False          #--COMPLETE
                mycompleteorder_status = True
                ARM.SoldPrice = ARM.Sell_BidPrice
            else:
                ARM.myopenorder_status = True           #--Keep OPEN
                mycompleteorder_status = False
            mycancleorder_status = False
        if (mycancleorder_status):
            ARM.price_try = 0.0
            ARM.order_id_sell = None
            ARM.myopenorder_status = False
            print("{0}, NW-PO {1} OPEN Order Cancelled in 2nd attempt"
                    .format(time.strftime("%H:%M:%S", time.localtime()),ARM.Sell_Symbol))
            if ARM.Sell_CE_PE == "CE":
                GUI.R3C2_Value = 1
                R3C2.set("{:.0f}".format(GUI.R3C2_Value))
                R26C2["text"] = f"CALCELLED"
                R26C2["bg"] = "yellow"
            else:
                GUI.R3C8_Value = 1
                R3C8.set("{:.0f}".format(GUI.R3C8_Value))
                R26C8["text"] = f"CALCELLED"
                R26C8["bg"] = "yellow"
            ARM.Sell_CE_PE = ""
            ARM.Sell_BidPrice = 0
            ARM.Sell_Strike = 0
            ARM.Sell_Qty = 0
            if ARM.Target_SL_Hit_Flag: ARM.Target_SL_Hit_Flag=False
        elif (ARM.myopenorder_status):
            if (ARM.Order_WaitTime >= 0):#--Wait for 5 Seconds
                if (ARM.Order_WaitTime == 10):print("{0}, NW-PO {1} Order is OPEN, (Qty-{2}), Wait for {3} Sec"#--5 4 3 2 1 sec in every iteration
                                                .format(time.strftime("%H:%M:%S", time.localtime()),ARM.Sell_Symbol,ARM.Sell_Qty,ARM.Order_WaitTime))
                if ARM.Sell_CE_PE == "CE":
                    R26C2["text"] = f"Wait ({ARM.Order_WaitTime} s)"
                    R26C2["bg"] = "yellow"
                else:
                    R26C8["text"] = f"Wait ({ARM.Order_WaitTime} s)"
                    R26C8["bg"] = "yellow"
                ARM.Order_WaitTime = ARM.Order_WaitTime - 1
            else:
                print("{0}, NW-PO {1} Could't get Sell Price, Canceling Order"
                            .format(time.strftime("%H:%M:%S", time.localtime()),ARM.Sell_Symbol))
                CANTechnicalError = False
                try:
                    if (ARM.MY_TRADE == "REAL"):        #--REAL Mode Trade
                        ARM.order_id_sell = kite.cancel_order(variety=kite.VARIETY_REGULAR,
                                                        order_id = ARM.order_id_sell,
                                                        parent_order_id=None)
                    else:pass
                except Exception as e:
                    CANTechnicalError = True
                    print("{0}, NW-PO {1} Sell Order Cancellation failed (TechnicalError)"
                                .format(time.strftime("%H:%M:%S", time.localtime()),ARM.Sell_Symbol))

                if (not CANTechnicalError):
                    mycancleorder_status = False
                    if (ARM.MY_TRADE == "REAL"):        #--REAL Mode Trade
                        time.sleep(0.5)
                        ARM.Zerodha_Orders_Flag = False
                        try:
                            ARM.My_DayOrders = kite.orders()
                        except Exception as e:
                            ARM.Zerodha_Orders_Flag = True
                        if not ARM.Zerodha_Orders_Flag:
                            for individual_order in ARM.My_DayOrders:
                                if int(individual_order['order_id']) == int(ARM.order_id_sell):
                                    mycancleorder_status = (individual_order['status'] == "CANCELLED")
                                    break
                                else:pass
                        else:
                            mycancleorder_status = False
                            ARM.myopenorder_status = False
                            myrejectedorder_status = False
                            mycompleteorder_status = False
                    else:
                        mycancleorder_status = True     #--PAPER Mode Trade
                    if mycancleorder_status:
                        ARM.price_try = 0.0
                        ARM.order_id_sell = None
                        ARM.myopenorder_status = False
                        print("{0}, NW-PO {1} OPEN Order Cancelled in 1st attempt"
                                .format(time.strftime("%H:%M:%S", time.localtime()),ARM.Sell_Symbol))
                        if ARM.Sell_CE_PE == "CE":
                            GUI.R3C2_Value = 1
                            R3C2.set("{:.0f}".format(GUI.R3C2_Value))
                            R26C2["text"] = f"CALCELLED"
                            R26C2["bg"] = "yellow"
                        else:
                            GUI.R3C8_Value = 1
                            R3C8.set("{:.0f}".format(GUI.R3C8_Value))
                            R26C8["text"] = f"CALCELLED"
                            R26C8["bg"] = "yellow"
                        ARM.Sell_CE_PE = ""
                        ARM.Sell_BidPrice = 0
                        ARM.Sell_Strike = 0
                        ARM.Sell_Qty = 0
                        if ARM.Target_SL_Hit_Flag: ARM.Target_SL_Hit_Flag=False
                    else:pass
                else:pass
        elif (myrejectedorder_status):
            ARM.price_try = 0.0
            ARM.order_id_sell = None
            ARM.myopenorder_status = False
            print("{0}, NW-PO {1} Order Placement Rejected"
                            .format(time.strftime("%H:%M:%S", time.localtime()),ARM.Sell_Symbol))
            if ARM.Sell_CE_PE == "CE":
                GUI.R3C2_Value = 1
                R3C2.set("{:.0f}".format(GUI.R3C2_Value))
                R26C2["text"] = f"REJECTED"
                R26C2["bg"] = "yellow"
            else:
                GUI.R3C8_Value = 1
                R3C8.set("{:.0f}".format(GUI.R3C8_Value))
                R26C8["text"] = f"REJECTED"
                R26C8["bg"] = "yellow"
            ARM.Sell_CE_PE = ""
            ARM.Sell_BidPrice = 0
            ARM.Sell_Strike = 0
            ARM.Sell_Qty = 0
            if ARM.Target_SL_Hit_Flag: ARM.Target_SL_Hit_Flag=False
        #----------------------------------------
        elif (mycompleteorder_status):
            ARM.price_try = 0.0
            ARM.order_id_sell = None
            ARM.myopenorder_status = False
            print("Sell Order Completed at Rs.{} Check Zerodha Account".format(ARM.SoldPrice))
            #--------------------------------------------------------------------------------------------
            if ARM.Sell_Symbol in ARM.Display_NetBuySell_df['Symbol'].values:
                #--Select specific columns
                columns_to_copy = ['Time_In','Action','Avg_Price','Total_Qty','SqOffPrice','Symbol','P/L(P)','P/L(Rs)']
                current_closed_trade = ARM.Display_NetBuySell_df.loc[ARM.Display_NetBuySell_df['Symbol'] == ARM.Sell_Symbol, columns_to_copy].copy()
                if current_closed_trade.at[current_closed_trade.index[0], 'Action'] == "BUY":
                    print("Yes, Bought Earlier, now Sell (Square-Off)")
                    #--Update cell values
                    current_closed_trade.at[current_closed_trade.index[0], 'Total_Qty'] = ARM.Sell_Qty
                    current_closed_trade.at[current_closed_trade.index[0], 'SqOffPrice'] = ARM.SoldPrice
                    current_closed_trade.at[current_closed_trade.index[0], 'P/L(P)'] = round((ARM.SoldPrice - current_closed_trade.iloc[0]['Avg_Price']), 2)
                    current_closed_trade.at[current_closed_trade.index[0], 'P/L(Rs)'] = round((current_closed_trade.iloc[0]['P/L(P)'] * ARM.Sell_Qty), 2)

                    #--Add Time Column
                    current_time = datetime.now().strftime('%H:%M:%S')
                    current_closed_trade.insert(0, 'Time_Out', current_time)

                    #--Display and Append
                    #print(current_closed_trade)
                    ARM.My_AllTrades = pd.concat([ARM.My_AllTrades, current_closed_trade], ignore_index=True)
                    print(ARM.My_AllTrades)
                else:
                    print("No, not Bought Earlier, Sell Again (Add Position)")
            else:
                print("No, not Bought Earlier, Fresh Sell (New Position)")
            #ARM.My_AllTrades = ARM.Display_NetBuySell_df.copy()
            #--------------------------------------------------------------------------------------------
            if ARM.Target_SL_Hit_Flag:
                key_to_search = ("SELL", ARM.oc_symbol_SqOff, ARM.Sell_Strike_SqOff, ARM.Sell_CE_PE_SqOff, ARM.Sell_Exp_SqOff)
            else:
                key_to_search = ("SELL", ARM.oc_symbol, ARM.Sell_Strike, ARM.Sell_CE_PE, ARM.oc_opt_expiry)
            if ARM.Option_Sold:
                key_to_update = None
                for key in ARM.Option_Sold:
                    if key == key_to_search:
                        key_to_update = key
                        break
                if key_to_update is not None:
                    ARM.Option_Sold[key_to_update][5] += ARM.Sell_Qty
                else:
                    ARM.Option_Sold[key_to_search] = list(key_to_search) + [ARM.Sell_Qty]
            else:
                ARM.Option_Sold[key_to_search] = list(key_to_search) + [ARM.Sell_Qty]
            #--------------------------------------------------------------------------------------------
            #ARM.SoldPrice = ARM.Sell_BidPrice  #--Price Tried
            ARM.TotalSellOrders += 1
            Time_In = datetime.now().strftime('%H:%M:%S')

            LPT_Pressed()
            PPT_Pressed()
            StopLossP = round(GUI.LPT_Value/ARM.found_lotsize_opt,ndigits = 2)
            TargetP = round(GUI.PPT_Value/ARM.found_lotsize_opt,ndigits = 2)
            Option_Sell_Info = (ARM.TotalSellOrders,Time_In,'SELL',ARM.SoldPrice,ARM.Sell_Qty,ARM.Sell_SqOffPrice,
                                ARM.Sell_Symbol,ARM.Sell_PL_P,ARM.Sell_PL_Rs,-StopLossP,TargetP,"NO")

            # Option_Sell_Info = (ARM.TotalSellOrders,'SELL',ARM.Sell_Symbol,ARM.SoldPrice,ARM.Sell_Qty,
            #                     ARM.Sell_SqOffPrice,ARM.Sell_PL_P,ARM.Sell_PL_Rs,-100.00,100.00,"NO")
            if ARM.NetBuySell_len > 0:
                for i in range(ARM.NetBuySell_len):
                    for j in range(len(ARM.Opt_Sold1_df)):
                        if ((ARM.Display_NetBuySell_df.loc[i,'Action'] == ARM.Opt_Sold1_df.loc[j,'Action']) and
                            (ARM.Display_NetBuySell_df.loc[i,'Symbol'] == ARM.Opt_Sold1_df.loc[j,'Symbol'])):
                            ARM.Opt_Sold1_df.loc[j,'SL(P)'] = ARM.Display_NetBuySell_df.loc[i,'SL(P)']
                            ARM.Opt_Sold1_df.loc[j,'Target(P)'] = ARM.Display_NetBuySell_df.loc[i,'Target(P)']
                        else:pass
                    if ((ARM.Display_NetBuySell_df.loc[i,'Action'] == "SELL") and
                        (ARM.Display_NetBuySell_df.loc[i,'Symbol'] == ARM.Sell_Symbol)):
                        Option_Sell_Info = (ARM.TotalSellOrders,ARM.Display_NetBuySell_df.loc[i,'Time_In'],'SELL',ARM.SoldPrice,ARM.Sell_Qty,ARM.Sell_SqOffPrice,
                                        ARM.Sell_Symbol,ARM.Sell_PL_P,ARM.Sell_PL_Rs,
                                        ARM.Display_NetBuySell_df.loc[i,'SL(P)'],ARM.Display_NetBuySell_df.loc[i,'Target(P)'],"NO")
                    else:pass
            else:pass
            ARM.Opt_Sold1_df.loc[len(ARM.Opt_Sold1_df)] = Option_Sell_Info
            #print("---------- Sold Option Contract ----------")
            #print(ARM.Opt_Sold1_df)
            if(len(ARM.Opt_Bought1_df)>0):
                #print("---------- Checking Buy Positions in Data ----------")
                # print(ARM.Opt_Bought1_df)
                # print(ARM.Opt_Sold1_df)
                # print("----")
                process_sell_trades()
                # print(ARM.Opt_Bought1_df)
                # print(ARM.Opt_Sold1_df)
                # print("----")
                #--------------------------------------------------------------------------------------------
                #--Combines Similar Buy Clicks which has same Symbol,Strike,Type & Price
                ARM.Opt_Bought2_df = ARM.Opt_Bought1_df.groupby(['Action','Symbol','Price','Qty'], as_index=False).agg({'Qty':'sum', 'OrderNo':'first',
                                        'Time_In':'first','SqOffPrice':'first','P/L(P)':'first','P/L(Rs)':'first','SL(P)':'first','Target(P)':'first','SQOff':'first'})
                ARM.Opt_Bought2_df = ARM.Opt_Bought2_df.sort_values(by=['Symbol','OrderNo'])
                ARM.Opt_Bought2_df = ARM.Opt_Bought2_df[['OrderNo','Time_In','Action','Price','Qty','SqOffPrice','Symbol','P/L(P)','P/L(Rs)','SL(P)','Target(P)','SQOff']]
                ARM.Opt_Bought2_df = ARM.Opt_Bought2_df.reset_index(drop=True)
                #print(ARM.Opt_Bought2_df)
                #--------------------------------------------------------------------------------------------
                #--Combines Similar Buy Clicks which has same Symbol,Strike,Type, has Avg_Price and Total_Qty
                ARM.Opt_Bought3_df = ARM.Opt_Bought2_df.copy()
                ARM.Opt_Bought3_df = ARM.Opt_Bought3_df.drop(columns=['OrderNo'])
                ARM.Opt_Bought3_df['Total_Qty'] = ARM.Opt_Bought3_df.groupby(['Symbol'])['Qty'].transform('sum')
                weighted_sum = ARM.Opt_Bought3_df['Price'] * ARM.Opt_Bought3_df['Qty']
                grouped_sum = weighted_sum.groupby([ARM.Opt_Bought3_df['Symbol']]).transform('sum')
                ARM.Opt_Bought3_df['Avg_Price'] = (grouped_sum / ARM.Opt_Bought3_df.groupby(['Symbol'])['Qty'].transform('sum')).round(2)
                mask = ARM.Opt_Bought3_df.duplicated(['Symbol'], keep=False)
                ARM.Opt_Bought3_df.loc[mask, 'Qty'] = ARM.Opt_Bought3_df.loc[mask, 'Total_Qty']
                ARM.Opt_Bought3_df['Price'] = ARM.Opt_Bought3_df['Price'].astype('float64')
                ARM.Opt_Bought3_df = ARM.Opt_Bought3_df[['Time_In','Action','Avg_Price','Total_Qty','SqOffPrice','Symbol','P/L(P)','P/L(Rs)','SL(P)','Target(P)','SQOff']]
                #ARM.Opt_Bought3_df = ARM.Opt_Bought3_df.drop_duplicates().reset_index(drop=True)
                ARM.Opt_Bought3_df = ARM.Opt_Bought3_df.drop_duplicates(subset=['Action', 'Symbol', 'Avg_Price']).reset_index(drop=True)
                #print(ARM.Opt_Bought3_df)
                #--------------------------------------------------------------------------------------------
            else:pass
            #--------------------------------------------------------------------------------------------
            #--Combines Similar Sell Clicks which has same Symbol,Strike,Type & Price
            ARM.Opt_Sold2_df = ARM.Opt_Sold1_df.groupby(['Action','Symbol','Price','Qty'], as_index=False).agg({'Qty':'sum', 'OrderNo':'first',
                                        'Time_In':'first','SqOffPrice':'first','P/L(P)':'first','P/L(Rs)':'first','SL(P)':'first','Target(P)':'first','SQOff':'first'})
            ARM.Opt_Sold2_df = ARM.Opt_Sold2_df.sort_values(by=['Symbol','OrderNo'])
            ARM.Opt_Sold2_df = ARM.Opt_Sold2_df[['OrderNo','Time_In','Action','Price','Qty','SqOffPrice','Symbol','P/L(P)','P/L(Rs)','SL(P)','Target(P)','SQOff']]
            ARM.Opt_Sold2_df = ARM.Opt_Sold2_df.reset_index(drop=True)
            #print(ARM.Opt_Sold2_df)
            #--------------------------------------------------------------------------------------------
            #--Combines Similar Sell Clicks which has same Symbol,Strike,Type, has Avg_Price and Total_Qty
            ARM.Opt_Sold3_df = ARM.Opt_Sold2_df.copy()
            ARM.Opt_Sold3_df = ARM.Opt_Sold3_df.drop(columns=['OrderNo'])
            ARM.Opt_Sold3_df['Total_Qty'] = ARM.Opt_Sold3_df.groupby(['Symbol'])['Qty'].transform('sum')
            weighted_sum = ARM.Opt_Sold3_df['Price'] * ARM.Opt_Sold3_df['Qty']
            grouped_sum = weighted_sum.groupby([ARM.Opt_Sold3_df['Symbol']]).transform('sum')
            ARM.Opt_Sold3_df['Avg_Price'] = (grouped_sum / ARM.Opt_Sold3_df.groupby(['Symbol'])['Qty'].transform('sum')).round(2)
            mask = ARM.Opt_Sold3_df.duplicated(['Symbol'], keep=False)
            ARM.Opt_Sold3_df.loc[mask, 'Qty'] = ARM.Opt_Sold3_df.loc[mask, 'Total_Qty']
            ARM.Opt_Sold3_df['Price'] = ARM.Opt_Sold3_df['Price'].astype('float64')
            ARM.Opt_Sold3_df = ARM.Opt_Sold3_df[['Time_In','Action','Avg_Price','Total_Qty','SqOffPrice','Symbol','P/L(P)','P/L(Rs)','SL(P)','Target(P)','SQOff']]
            #ARM.Opt_Sold3_df = ARM.Opt_Sold3_df.drop_duplicates().reset_index(drop=True)
            ARM.Opt_Sold3_df = ARM.Opt_Sold3_df.drop_duplicates(subset=['Action', 'Symbol', 'Avg_Price']).reset_index(drop=True)
            #print(ARM.Opt_Sold3_df)
            #---------------------------------------------------------------------------------------------
            ARM.Display_NetBuySell_df = pd.concat([ARM.Opt_Bought3_df, ARM.Opt_Sold3_df], ignore_index=True)
            ARM.NetBuySell_len = len(ARM.Display_NetBuySell_df)
            #--------------------------------------------------------------------------------------------
            # print(ARM.Option_Bought)
            # print(ARM.Option_Sold)
            # print(ARM.Option_NetBuySell)
            for buy_key, buy_values in ARM.Option_Bought.items():
                sell_key = ('SELL',) + buy_key[1:]
                if sell_key not in ARM.Option_Sold:
                    ARM.Option_NetBuySell[buy_key] = buy_values
                else:pass
            for sell_key, sell_values in ARM.Option_Sold.items():
                buy_key = ('BUY',) + sell_key[1:]
                if buy_key not in ARM.Option_Bought:
                    ARM.Option_NetBuySell[sell_key] = sell_values
                else:pass
            NetBuySell_modified = False
            for buy_key, buy_values in ARM.Option_Bought.items():
                for sell_key, sell_values in ARM.Option_Sold.items():
                    if buy_key[1:] == sell_key[1:]:
                        quantity = buy_values[5] - sell_values[5]
                        if quantity > 0:
                            ARM.Option_Bought[buy_key] = buy_values[:5] + [quantity]
                            ARM.Option_Sold.pop(sell_key, None)
                            ARM.Option_NetBuySell.pop(sell_key, None)
                            ARM.Option_NetBuySell[buy_key] = buy_values[:5] + [quantity]
                            NetBuySell_modified = True
                            break
                        elif quantity < 0:
                            quantity = abs(quantity)
                            ARM.Option_Bought.pop(buy_key, None)
                            ARM.Option_Sold[sell_key] = sell_values[:5] + [quantity]
                            ARM.Option_NetBuySell.pop(buy_key, None)
                            ARM.Option_NetBuySell[sell_key] = sell_values[:5] + [quantity]
                            NetBuySell_modified = True
                            break
                        elif quantity == 0:
                            ARM.Option_Bought.pop(buy_key, None)
                            ARM.Option_Sold.pop(sell_key, None)
                            ARM.Option_NetBuySell.pop(buy_key, None)
                            ARM.Option_NetBuySell.pop(sell_key, None)
                            NetBuySell_modified = True
                            break
                        else:pass
                    else:pass
                if NetBuySell_modified:break
            # print(ARM.Option_Bought)
            # print(ARM.Option_Sold)
            # print(ARM.Option_NetBuySell)
            #--------------------------------------------------------------------------------------------
            #print("Trades in Display_NetBuySell_df: ",ARM.NetBuySell_len)
            #print(ARM.Display_NetBuySell_df)
            if (ARM.NetBuySell_len > 0):
                if not ARM.Desplay_NetPosition_flag :
                    ARM.Desplay_NetPosition_flag = True
                    Desplay_NetPosition()
                else:pass
                num_rows, num_columns = ARM.Display_NetBuySell_df.shape
                for i in range(28,ARM.pev_num_rows+28): #--Clear Previous
                    for j in range(0,ARM.prev_num_columns):
                        if j == 8:
                            OC_Cell_Value[i][j].set(-100.00)
                            OC_Cell[i][j].config(fg="#F0F0F0",bg="#F0F0F0")
                        elif j == 9:
                            OC_Cell_Value[i][j].set(100.00)
                            OC_Cell[i][j].config(fg="#F0F0F0",bg="#F0F0F0")
                        else:
                            OC_Cell[i][j].config(text="",fg="black",bg="#F0F0F0")
                    #OC_Cell[i][15].config(text="",fg="#F0F0F0",bg="#F0F0F0")

                OC_Cell[28][10].config(text="",fg="#F0F0F0",bg="#F0F0F0")
                OC_Cell[29][10].config(text="",fg="#F0F0F0",bg="#F0F0F0")
                OC_Cell[30][10].config(text="",fg="#F0F0F0",bg="#F0F0F0")
                OC_Cell[31][10].config(text="",fg="#F0F0F0",bg="#F0F0F0")
                OC_Cell[32][10].config(text="",fg="#F0F0F0",bg="#F0F0F0")
                ARM.pev_num_rows = num_rows
                ARM.prev_num_columns = num_columns
            else:
                for i in range(28,32): #--Clear Previous
                    for j in range(0,10):
                        if j == 8:
                            OC_Cell_Value[i][j].set(-100.00)
                            OC_Cell[i][j].config(fg="#F0F0F0",bg="#F0F0F0")
                        elif j == 9:
                            OC_Cell_Value[i][j].set(100.00)
                            OC_Cell[i][j].config(fg="#F0F0F0",bg="#F0F0F0")
                        else:
                            OC_Cell[i][j].config(text="",fg="black",bg="#F0F0F0")
                OC_Cell[28][10].config(text="",fg="#F0F0F0",bg="#F0F0F0")
                OC_Cell[29][10].config(text="",fg="#F0F0F0",bg="#F0F0F0")
                OC_Cell[30][10].config(text="",fg="#F0F0F0",bg="#F0F0F0")
                OC_Cell[31][10].config(text="",fg="#F0F0F0",bg="#F0F0F0")
                OC_Cell[32][10].config(text="",fg="#F0F0F0",bg="#F0F0F0")
                ARM.Desplay_NetPosition_flag = False
                #print("Stop Display 7" )
            #--------------------------------------------------------------------------------------------
            #print("Clearing Sell Veriables")
            if ARM.Sell_CE_PE == "CE":
                GUI.R3C2_Value = 1
                R3C2.set("{:.0f}".format(GUI.R3C2_Value))
            else:
                GUI.R3C8_Value = 1
                R3C8.set("{:.0f}".format(GUI.R3C8_Value))
            ARM.Sell_CE_PE = ""
            ARM.Sell_BidPrice = 0
            ARM.Sell_Strike = 0
            ARM.Sell_Qty = 0
            if ARM.Target_SL_Hit_Flag: ARM.Target_SL_Hit_Flag=False
            #--------------------------------------------------------------------------------------------
            #print("Ending Net Option Buy Sell")
        else:pass
    if ARM.myopenorder_status:root.after(1000,APH_SELL)
def Verify_Manual_Enctoken():
    global kite
    try:
        Check_etoken = kite.margins()
        if (Check_etoken == None):
            ARM.My_try = ARM.My_try + 1
            if (ARM.My_try < 3):
                try:
                    winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
                except:pass
                if ARM.Etoken_File_Exists:
                    os.remove('Etoken.txt')
                    ARM.Etoken_File_Exists = False
                else:pass
                User_New_Etoken = ""
                print("Please Provide New Enctoken, (Attempt : {0}/2)".format(ARM.My_try))
                User_New_Etoken = POPUP_ETOKEN(ARM.your_user_id)
                if (User_New_Etoken != ARM.User_Prv_Etoken):
                    ARM.enctoken = ARM.User_Prv_Etoken = User_New_Etoken
                    kite = KiteApp(enctoken=ARM.enctoken)
                    with open('Etoken.txt', 'w') as efile:
                        efile.write(ARM.your_user_id)
                        efile.write("\n")
                        efile.write(ARM.enctoken)
                    efile.close()
                    ARM.Etoken_File_Exists = True
                else:pass
            else:
                return
        else:
            return
    except:
        APH_Delay(5)
    Verify_Manual_Enctoken()
def Verify_Chrome_Enctoken():
    global kite
    try:
        Check_etoken = kite.margins()       #--Enable after Zerodha-Maintanance
        #Check_etoken =  "Arvind"           #--Disable after Zerodha-Maintanance
        if (Check_etoken == None):
            ARM.My_try = ARM.My_try + 1

            if (ARM.My_try < 3):
                if ARM.Etoken_File_Exists:
                    os.remove('Etoken.txt')
                    ARM.Etoken_File_Exists = False
                else:pass
                User_New_Etoken = ""
                if (ARM.My_try == 1):
                    if os.path.exists(Cookiesfile_path):
                        with open(Cookiesfile_path, 'r') as file:
                            cookies_str = file.read()
                        #--Extracting the part within "cookies = {...}"
                        start_index = cookies_str.find("{")
                        end_index = cookies_str.rfind("}")
                        if start_index != -1 and end_index != -1:
                            cookies_str = cookies_str[start_index:end_index + 1]
                        try:
                            cookies_dict = ast.literal_eval(cookies_str)
                            #print(cookies_dict)
                            if "user_id" in cookies_dict and cookies_dict["user_id"] == ARM.your_user_id:
                                if "enctoken" in cookies_dict:
                                    User_New_Etoken = cookies_dict["enctoken"] + "=="
                                    #print("User ID is {}, and the Modified Enctoken is:{}".format(ARM.your_user_id, User_New_Etoken))
                                else:
                                    print("Enctoken not found for user_id: {}".format(ARM.your_user_id))
                            else:pass   #--print("User ID is not {}".format(ARM.your_user_id))
                        except (SyntaxError, ValueError):
                            print("Error while parsing cookies content")
                            cookies_dict = {}
                    else:
                        print("Cookie file Not Found")

                elif (ARM.My_try > 1):
                    try:
                        winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
                    except:pass
                    print("Trying Manual Login, -- for User-ID : {}".format(ARM.your_user_id))
                    User_New_Etoken = POPUP_ETOKEN(ARM.your_user_id)
                else:pass

                if (User_New_Etoken != ARM.User_Prv_Etoken):
                    ARM.enctoken = ARM.User_Prv_Etoken = User_New_Etoken
                    kite = KiteApp(enctoken=ARM.enctoken)
                    with open('Etoken.txt', 'w') as efile:
                        efile.write(ARM.your_user_id)
                        efile.write("\n")
                        efile.write(ARM.enctoken)
                    efile.close()
                    ARM.Etoken_File_Exists = True
                else:pass
            else:
                return
        else:
            return
    except:
        APH_Delay(5)
    Verify_Chrome_Enctoken()

#====================================================================================================
def Process_KWS_Ticks(ticks):
    ARM.Wait_KWS_Process = True
    #print("********************** Processing Tick Data **********************")
    if not ARM.KWS_Functional:ARM.KWS_Functional = True
    if ARM.KWS_Working:
        kws_canvas.itemconfig(square, fill='white')
        ARM.KWS_Working = False
    else:
        kws_canvas.itemconfig(square, fill='#D6D6D2')
        ARM.KWS_Working = True
    #------------------------------------------------
    for symbol in ticks:
        try:
            #symbol_token = ARM.KWS_Subscribe_Tokens[symbol['instrument_token']]['symbol']
            if symbol['instrument_token'] in ARM.KWS_OPT_Tokens.keys():
                #print("Option Symbol Token",symbol_token)
                ARM.KWS_OPT_Data[ARM.KWS_Subscribe_Tokens[symbol['instrument_token']]['symbol']] = {"strikePrice":ARM.KWS_Subscribe_Tokens[symbol['instrument_token']]["strike"],
                                                                                            "instrumentType":ARM.KWS_Subscribe_Tokens[symbol['instrument_token']]["instrumentType"],
                                                                                            "bid":symbol['depth']['buy'][0]['price'],
                                                                                            "ask":symbol['depth']['sell'][0]['price'],
                                                                                            "ltp":symbol["last_price"],
                                                                                            "averagePrice":symbol["average_traded_price"],
                                                                                            # "totalTradedVolume":int(symbol["volume_traded"]/ARM.found_lotsize_opt),
                                                                                            # "openInterest":int(symbol["oi"]/ARM.found_lotsize_opt),
                                                                                            "change":symbol["last_price"] - symbol["ohlc"]["close"] if symbol["last_price"] != 0 else 0,
                                                                                            # # "changeinOpenInterest":int((symbol["oi"] - ARM.prev_day_oi_opt[symbol])/ARM.found_lotsize_opt),
                                                                                            # "changeinOpenInterest":symbol["oi"],
                                                                                            "expiry":ARM.KWS_Subscribe_Tokens[symbol['instrument_token']]["expiry"]
                                                                                            }
            else:
                #print("EQ / Index Symbol Token",symbol_token)
                ARM.KWS_SPT_Data[ARM.KWS_Subscribe_Tokens[symbol['instrument_token']]['symbol']] = {"Ltp":symbol["last_price"],
                                                            "Open":symbol["ohlc"]["open"],
                                                            "High":symbol["ohlc"]["high"],
                                                            "Low":symbol["ohlc"]["low"],
                                                            "Close":symbol["ohlc"]["close"]
                                                            }

        except Exception as e:
            #print(e)
            pass
    #if len (ARM.KWS_SPT_Data): print(ARM.KWS_SPT_Data)
    # print ('Received Data : ', len (ticks), '. Stored data :', len (ARM.KWS_SPT_Data))
    # print ('Received Data : ', len (ticks), '. Stored data :', len (ARM.KWS_OPT_Data))
    if ARM.kws_option_flag and ARM.KWS_OPT_Tokens: ARM.kws_option_flag = False
    #------------------------------------------------
    if ARM.oc_symbol == "NIFTY":Spot_Symbol = "NIFTY 50"
    else:pass

    if Spot_Symbol in ARM.KWS_SPT_Data:
        #if len (ARM.KWS_SPT_Data): print(ARM.KWS_SPT_Data)
        ARM.Spot_Value = ARM.KWS_SPT_Data[Spot_Symbol]["Ltp"]
        #print("Spot_Value",ARM.Spot_Value)
        OC_Cell[3][5].config(text=str(ARM.Spot_Value))

        # R1C8.set(ARM.KWS_SPT_Data[Spot_Symbol]["Open"])
        # R1C9.set(ARM.KWS_SPT_Data[Spot_Symbol]["High"])
        # R1C10.set(ARM.KWS_SPT_Data[Spot_Symbol]["Low"])
        if(ARM.Spot_Value>ARM.KWS_SPT_Data[Spot_Symbol]["Open"]):
            OC_Cell[3][5].config(fg="white",bg="green")
        elif(ARM.Spot_Value<ARM.KWS_SPT_Data[Spot_Symbol]["Open"]):
            OC_Cell[3][5].config(fg="white",bg="red")
        else:
            OC_Cell[3][5].config(fg="black",bg="yellow")
    else:(f"The key {Spot_Symbol} is not found in the Dictionary")

    #print ('Received Data : ', len (ticks), '. Stored data :', len (ARM.KWS_SPT_Data))

    #------------------------------------------------
    if ARM.Spot_Value != 0:

        oc_opt_expiry = datetime.strptime(R1C2.get(), '%Y-%m-%d').date()

        myexpiry_date = datetime.combine(ARM.oc_opt_expiry, ARM.Trade_Stop_Time)
        myexpiry_date = myexpiry_date - datetime.today()
        ARM.Script_YOE = myexpiry_date/timedelta(days=1)/365
        #print(f"Number of YOE: {ARM.Script_YOE}")   #--Just for Testing
        #------------------------------------------
        df_opt_temp = ARM.df_opt[(ARM.df_opt["name"] == ARM.oc_symbol) & (ARM.df_opt["expiry"] == oc_opt_expiry)]
        df_opt_temp.drop(columns=['index', 'name', 'expiry', 'exchange_token'], inplace=True) #--Remove the specified columns

        #--Ensure both CE and PE exist for each strike value
        #--Group by strike and filter out groups where both CE and PE do not exist
        grouped = df_opt_temp.groupby('strike').filter(lambda x: set(x['instrument_type']) == {'CE', 'PE'})

        #--Split the dataframe into CE and PE dataframes
        ce_df = grouped[grouped['instrument_type'] == 'CE'].reset_index(drop=True)
        ce_df.sort_values(by='strike', ascending=True, inplace=True)
        ce_df = ce_df.reset_index(drop=True)

        pe_df = grouped[grouped['instrument_type'] == 'PE'].reset_index(drop=True)
        pe_df.sort_values(by='strike', ascending=True, inplace=True)
        pe_df = pe_df.reset_index(drop=True)

        #--Rearrange and create the required columns
        cepe_opt_df = pd.DataFrame({
            'CEinstrument_token': ce_df['instrument_token'],
            'CEtradingsymbol': ce_df['tradingsymbol'],
            'strike': ce_df['strike'],
            'PEtradingsymbol': pe_df['tradingsymbol'],
            'PEinstrument_token': pe_df['instrument_token']
        })

        cepe_opt_df.sort_values(by='strike', ascending=True, inplace=True)
        cepe_opt_df = cepe_opt_df.reset_index(drop=True)

        #cepe_opt_df.to_csv("Available_Strikes1.csv", index=True)
        #print("Spot_Value",ARM.Spot_Value)
        atm_index = cepe_opt_df['strike'].sub(ARM.Spot_Value).abs().astype('float64').idxmin() #--Find ATM Index Location
        #print(f"First:ATM_Index: {atm_index}, ATM_Strike: {cepe_opt_df.loc[atm_index,'strike']}, Max Length: {len(cepe_opt_df)}")
        try:
            Lower_Strikes = atm_index
            Higher_Strikes = len(cepe_opt_df) - atm_index - 1
            if (Lower_Strikes < 10 or Higher_Strikes < 10):
                if Lower_Strikes == Higher_Strikes:ARM.OC_Strikes = Lower_Strikes
                elif Lower_Strikes < Higher_Strikes:ARM.OC_Strikes = Lower_Strikes
                else:ARM.OC_Strikes = Higher_Strikes
            elif (Lower_Strikes >= 10 and Higher_Strikes >= 10):
                ARM.OC_Strikes = 10
            else:pass
            ARM.Display_Row = 15 - ARM.OC_Strikes
            # print(f"OC_Strikes: {ARM.OC_Strikes}, Display_Row :{ARM.Display_Row}")
            # print(f"ATM_Strike: {cepe_opt_df.loc[atm_index,'strike']}, Lower Strikes: {Lower_Strikes},  Higher Strikes: {Higher_Strikes}, OC Strikes: ({ARM.OC_Strikes} x 2) + 1 = {(ARM.OC_Strikes*2)+1}")
        except:pass
        #--------------------------------------
        maxlen = len(cepe_opt_df)
        #print("1. Available Strikes:",maxlen)
        cepe_opt_df.drop(cepe_opt_df.loc[0:atm_index-ARM.OC_Strikes-1].index, inplace=True)
        cepe_opt_df.drop(cepe_opt_df.loc[atm_index+ARM.OC_Strikes+1:maxlen-1].index, inplace=True)
        cepe_opt_df = cepe_opt_df.reset_index(drop=True) #--Reset Index
        #cepe_opt_df.to_csv("Available_Strikes2.csv", index=True)
        #print("2. Available Strikes:",len(cepe_opt_df))
        atm_index = cepe_opt_df['strike'].sub(ARM.Spot_Value).abs().astype('float64').idxmin() #--Find ATM Index Location
        #print(f"2:ATM_Index: {atm_index}, ATM_Strike: {cepe_opt_df.loc[atm_index,'strike']}, Max Length: {len(cepe_opt_df)}")

        #print(cepe_opt_df)
        #-------------------------------------------
        #--Create CE dataframe
        ce_opt_df = cepe_opt_df[['CEinstrument_token', 'CEtradingsymbol', 'strike']]
        ce_opt_df.columns = ['instrument_token', 'tradingsymbol', 'strike']
        ce_opt_df['instrument_type'] = 'CE'

        #--Create PE dataframe
        pe_opt_df = cepe_opt_df[['PEinstrument_token', 'PEtradingsymbol', 'strike']]
        pe_opt_df.columns = ['instrument_token', 'tradingsymbol', 'strike']
        pe_opt_df['instrument_type'] = 'PE'

        #opt_df = pd.concat([ce_opt_df, pe_opt_df]).reset_index(drop=True)
        #-----------------------------------------------------------------------------
        opt_df = pd.concat([ce_opt_df, pe_opt_df, ARM.STD_STR_Subscribe], ignore_index=True)
        # print("--------- before ------------")
        # print(opt_df)
        # print(ARM.KWS_Subscribe_Tokens.keys())
        instrument_tokens = ARM.KWS_Subscribe_Tokens.keys()
        opt_df = opt_df[~opt_df['instrument_token'].isin(instrument_tokens)] #--Filter the DataFrame
        # print("--------- after ------------")
        # print(opt_df)
        # print("--------- finish ------------")
        if (len(opt_df)):
            #opt_df.to_csv("Check123.csv", index=False)
            # print("-----------------------------------")                      #--Just for Testing
            # print(f"Existing Data Tokens : {len(ARM.KWS_Subscribe_Tokens)}")  #--Just for Testing
            # print(f"Subscribe New Tokens : {len(opt_df)}")                    #--Just for Testing
            for i in opt_df.index:
                ARM.KWS_Subscribe_Tokens[int(opt_df["instrument_token"][i])] = {"symbol": opt_df["tradingsymbol"][i],
                                                                "strike": float(opt_df["strike"][i]),
                                                                "instrumentType": opt_df["instrument_type"][i],
                                                                "expiry": oc_opt_expiry}
                ARM.KWS_OPT_Tokens[int(opt_df["instrument_token"][i])] = {"symbol": opt_df["tradingsymbol"][i],
                                                                "strike": float(opt_df["strike"][i]),
                                                                "instrumentType": opt_df["instrument_type"][i],
                                                                "expiry": oc_opt_expiry,
                                                                #"type": "OPT"
                                                                }

            # print("================== OPT TOKENS NF ==================")
            # print(ARM.KWS_Subscribe_Tokens)
            # if ARM.KWS_Subscribe_Tokens:                      #--Just for Testing START
            #     print(f"Total Subscribed Tokens : {len(ARM.KWS_Subscribe_Tokens)}")
            # else: print("Token List is Empty")
            # cepe_strike = opt_df['strike'].drop_duplicates().astype(str)
            # cepe_strike_tokens = ' '.join(cepe_strike)
            # print("{} CE & PE, {} Tokens : {}".format(time.strftime("%H:%M:%S", time.localtime()),ARM.oc_symbol,cepe_strike_tokens))
            # print("-----------------------------------")      #--Just for Testing END

            # print(ARM.KWS_OPT_Tokens)
            # print(ARM.KWS_Subscribe_Tokens)
            # print(ARM.KWS_Subscribe_Tokens.keys())           #--Prints dict_keys([256265, 260105, 257801])
            # print(list(ARM.KWS_Subscribe_Tokens.keys()))     #--Prints [256265, 260105, 257801]

            kws_nse.subscribe(list(ARM.KWS_Subscribe_Tokens.keys()))
            kws_nse.set_mode(kws_nse.MODE_FULL,list(ARM.KWS_Subscribe_Tokens.keys()))

            ARM.kws_option_flag = True
        else:pass
    else:pass
    #------------------------------------------------
    ARM.Wait_KWS_Process = False
    # print("Assign callback to on_ticks_ARMKWS")
    # kws_nse.on_ticks = on_ticks_ARMKWS          #--Assign callback to on_ticks_ARMKWS
def on_ticks_ARMKWS(ws, ticks):
    if ARM.New_Token_Ready:
        # print("Tokens to Unsubscribe")
        # print(ARM.KWS_Subscribe_Tokens)
        # print("1. Count KWS_Subscribe_Tokens",len(ARM.KWS_Subscribe_Tokens))
        # print("2. Count KWS_OPT_Tokens",len(ARM.KWS_OPT_Tokens))

        if ARM.NetBuySell_len > 0:
            # print(f"Currently you have {ARM.NetBuySell_len} Running Trade")
            ARM.symbols_in_net_positions = set(ARM.Display_NetBuySell_df['Symbol'].values) #--Extract symbols from Net Positions
            retain_position_tokens = {     #--Find Retain Position Token
                                        token: details
                                        for token, details in ARM.KWS_Subscribe_Tokens.items()
                                        if details.get('symbol') in ARM.symbols_in_net_positions
                                    }
            #print(ARM.symbols_in_net_positions)
        else:
            ARM.symbols_in_net_positions = {}
            # print("Currently you Do Not have Running Trade")

        if ARM.Symbol_Changed:
            ARM.KWS_OPT_Data = {}
            ARM.kws_option_flag = False

            ARM.KWS_OPT_Tokens = {}
            #print("Total Tokens (SYM)",len(ARM.KWS_Subscribe_Tokens))
            ws.unsubscribe(list(ARM.KWS_Subscribe_Tokens.keys()))
            ARM.KWS_Subscribe_Tokens = {}
            #print("Total Tokens (SYM)",len(ARM.KWS_Subscribe_Tokens))
            ARM.KWS_Subscribe_Tokens.update(ARM.New_SYM_Tokens)    #--Then add New Symbol Tokens
            ARM.KWS_Subscribe_Tokens.update(ARM.New_FUT_Tokens)    #--Then add New Future Tokens
            if ARM.NetBuySell_len > 0:
                ARM.KWS_Subscribe_Tokens.update(retain_position_tokens)
                ARM.KWS_OPT_Tokens.update(retain_position_tokens)
            else:pass

            ws.subscribe(list(ARM.KWS_Subscribe_Tokens.keys()))
            ws.set_mode(kws_nse.MODE_FULL,list(ARM.KWS_Subscribe_Tokens.keys()))
            #print("Total Tokens (SYM)",len(ARM.KWS_Subscribe_Tokens))
            ARM.Symbol_Changed = False
            ARM.APH_SharedPVar[1] = ARM.oc_symbol
        elif ARM.Opt_Exp_Changed:
            ARM.KWS_OPT_Data = {}
            ARM.kws_option_flag = False

            ARM.KWS_OPT_Tokens = {}
            #print("Total Tokens (OPT)",len(ARM.KWS_Subscribe_Tokens))
            ws.unsubscribe(list(ARM.KWS_Subscribe_Tokens.keys()))
            ARM.KWS_Subscribe_Tokens = {}
            #print("Total Tokens (OPT)",len(ARM.KWS_Subscribe_Tokens))
            ARM.KWS_Subscribe_Tokens.update(ARM.New_SYM_Tokens)    #--Then add New Symbol Tokens
            ARM.KWS_Subscribe_Tokens.update(ARM.New_FUT_Tokens)    #--Then add New Future Tokens
            if ARM.NetBuySell_len > 0:
                ARM.KWS_Subscribe_Tokens.update(retain_position_tokens)
                ARM.KWS_OPT_Tokens.update(retain_position_tokens)
            else:pass

            ws.subscribe(list(ARM.KWS_Subscribe_Tokens.keys()))
            ws.set_mode(kws_nse.MODE_FULL,list(ARM.KWS_Subscribe_Tokens.keys()))
            #print("Total Tokens (OPT)",len(ARM.KWS_Subscribe_Tokens))
            ARM.Opt_Exp_Changed = False
        elif ARM.OptFut_Exp_Changed:
            ARM.KWS_OPT_Data = {}
            ARM.kws_option_flag = False

            ARM.KWS_OPT_Tokens = {}
            #print("Total Tokens (OPT & FUT)",len(ARM.KWS_Subscribe_Tokens))
            ws.unsubscribe(list(ARM.KWS_Subscribe_Tokens.keys()))
            ARM.KWS_Subscribe_Tokens = {}
            #print("Total Tokens (OPT & FUT)",len(ARM.KWS_Subscribe_Tokens))
            ARM.KWS_Subscribe_Tokens.update(ARM.New_SYM_Tokens)    #--Then add New Symbol Tokens
            ARM.KWS_Subscribe_Tokens.update(ARM.New_FUT_Tokens)    #--Then add New Future Tokens
            if ARM.NetBuySell_len > 0:
                ARM.KWS_Subscribe_Tokens.update(retain_position_tokens)
                ARM.KWS_OPT_Tokens.update(retain_position_tokens)
            else:pass

            ws.subscribe(list(ARM.KWS_Subscribe_Tokens.keys()))
            ws.set_mode(kws_nse.MODE_FULL,list(ARM.KWS_Subscribe_Tokens.keys()))
            #print("Total Tokens (OPT & FUT)",len(ARM.KWS_Subscribe_Tokens))
            ARM.OptFut_Exp_Changed = False
        else:
            print("! ! ! - - Something Went Wrong (111) - - ! ! !")
        ARM.New_Token_Ready = False
    elif not ARM.Wait_KWS_Process:
        root.after(0, lambda: Process_KWS_Ticks(ticks))
    else:
        pass
def on_connect_ARMKWS(ws,response):
    ws.subscribe(list(ARM.KWS_Subscribe_Tokens.keys()))
    ws.set_mode(ws.MODE_FULL,list(ARM.KWS_Subscribe_Tokens.keys())) #--(_FULL, _QUOTE, _LTP)
#====================================================================================================

def TT_Login():
    global kite,R0C0entry,R0C0options,R1C2entry #--(destroy)
    if (ARM.User_Type == ""):
        if(R0C0.get() == "Manual Login"):
            ARM.User_Type = "Manual"
            R0C0options[0] = "Logging"
            R0C0.set(R0C0options[0])
            R0C0entry.configure(state="disabled")  #--Disable the OptionMenu until Save is clicked
            POPUP_LOGIN()
        elif(R0C0.get() == "Browser Login"):
            ARM.User_Type = "Browser"
            R0C0options[0] = "Logging"
            R0C0.set(R0C0options[0])
            R0C0entry.configure(state="disabled")  #--Disable the OptionMenu until Save is clicked
            POPUP_LOGIN()
        elif(R0C0.get() == "Logout"):
            R0C0.set("Login")
            R0C0entry.configure(state="disabled")  #--Disable the OptionMenu until Save is clicked
            POPUP_LOGOUT()
        else:pass
    else:pass
    if (ARM.User_Type != ""):
        if (not ARM.Login_selected_flag):
            #================== Added for Server Database Connection START ===========================
            try:
                if ((ARM.your_user_id == "DN4808") or (ARM.your_user_id == "AZM042") or #--Gaba
                    (ARM.your_user_id == "MV1394") or (ARM.your_user_id == "MX0597") or #--Arvind
                    (ARM.your_user_id == "ERL063") or #--Sneha
                    (ARM.your_user_id == "XR9475")  or #--Tambat
                    (ARM.your_user_id == "QR8003")  # --Meena

                    ):
                    pass
                else:
                    print("\nVerifying User : ",ARM.your_user_id)
                    print("You have not Subscribed for this ARM Version,")
                    print("\nTo Subscribe visit https://prakashgaba.com")
                    print("To Subscribe contact gabamentoring@gmail.com")
                    sys.exit()
            except:
                sys.stdout.close()
                if (os.path.exists(TxtLogFilePath)):
                    convert_txt_to_pdf(TxtLogFilePath, PdfLogFilePath)
                    os.remove(TxtLogFilePath)
                else:pass
                sys.exit()

            if not (os.path.exists('Etoken.txt')):
                #print("Etoken File Not Exists")
                ARM.Etoken_File_Exists = False
                ARM.enctoken = "Arvind"
            else:
                with open('Etoken.txt', 'r') as efile:
                    lines = efile.readlines()
                efile.close()
                if(len(lines) == 2):
                    if (lines[0][0:6] == ARM.your_user_id):
                        ARM.enctoken = lines[1]
                        ARM.Etoken_File_Exists = True
                    else:
                        ARM.enctoken = "Arvind"
                        ARM.Etoken_File_Exists = False
                else:
                    ARM.enctoken = "Arvind"
                    ARM.Etoken_File_Exists = False
            ARM.User_Prv_Etoken = ARM.enctoken

            try:
                kite = KiteApp(enctoken=ARM.enctoken)
            except:
                print("Please Check Internet Connection, Not Connecting to Zerodha ( KAPI )\n")

            ARM.My_try = 0
            print("{} {} Login".format(ARM.your_user_id,ARM.User_Type))
            if (ARM.User_Type == "Manual"):
                Verify_Manual_Enctoken()
            elif (ARM.User_Type == "Browser"):
                Verify_Chrome_Enctoken()
            else:pass

            try:
                Cookiesfile_to_del = glob.glob(os.path.join(os.getenv("USERPROFILE"),"Downloads","cookies*.txt"))
                if Cookiesfile_to_del: #--Remove cookies*.txt from Download Folder
                    for delfile_path in Cookiesfile_to_del:
                        os.remove(delfile_path)
            except:pass

            try:
                Login = kite.margins() #--Enable after Zerodha-Maintanance
                time.sleep(1)
                user_info = kite.profile()["user_id"]
                user_id_enctoken = True if ((user_info == ARM.your_user_id) and (Login != None)) else False  #--Enable after Zerodha-Maintanance
                #user_id_enctoken = True if ((user_info == ARM.your_user_id) ) else False     #--Disable after Zerodha-Maintanance
            except:
                user_id_enctoken = False
                print("\nYou have not provided correct Etoken / User-ID\n")

            if not user_id_enctoken:
                try:
                    winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
                except:pass

                print("-------------------------------------------------------------------------------------") #--85 White Spaces
                print("Trading Session End                         {}".format(datetime.now().strftime("%H:%M:%S, %d %b %Y, %a")))
                print("-------------------------------------------------------------------------------------")
                sys.stdout.close()
                if (os.path.exists(TxtLogFilePath)):
                    convert_txt_to_pdf(TxtLogFilePath, PdfLogFilePath)
                    os.remove(TxtLogFilePath)
                else:pass
                sys.exit()
            else:
                try:
                    kite = KiteApp(enctoken=ARM.enctoken,userid=ARM.your_user_id)
                except:
                    print("Please Check Internet Connection, Not Connecting to Zerodha ( KWS )\n")
            #================== Added for Server Database Connection END ===========================
            root.title("{} ({}L-{})".format(ARM.ARM_Version,ARM.License_Type,ARM.your_user_id))
            if (ARM.License_Type == "R"):
                if (GUI.Trade_Flag == 'DEMO_ONLY'):
                    ARM.License_Type = "P"
                    GUI.Trade_Flag = 'DEMO_ONLY'
                    print("{} Subscription Real License (Mode: DEMO)".format(ARM.your_user_id))
                else:
                    ARM.License_Type = "R"
                    GUI.Trade_Flag = 'DEMO_REAL_BOTH'
                    print("{} Subscription Real License (Mode: DEMO/REAL)".format(ARM.your_user_id))
                ARM.MY_TRADE = GUI.R0C5flag
            else:
                GUI.Trade_Flag = 'DEMO_ONLY'
                ARM.MY_TRADE = "DEMO"
                print("{} Subscription Demo License".format(ARM.your_user_id))

            if ARM.MY_TRADE == "REAL":
                ARM.NO_Response = 2
                Get_OB_LB()        #--Enable after Zerodha-Maintanance
                ARM.Opening_Balance = GUI.R1C3_Value
                ARM.Live_Balance = GUI.R1C4_Value
            else:
                R1C3.set("{:.2f}".format(GUI.R1C3_Value))
                R1C4.set("{:.2f}".format(GUI.R1C4_Value))

            #-----------------------------------
            # ARM.instrument_nse.to_csv("Instrument_NSE.csv", index=False)
            # ARM.instrument_nfo.to_csv("Instrument_NFO.csv", index=False)
            ARM.todays_server_date = datetime.now().date()
            Total_120Days_Exp = ARM.todays_server_date + timedelta(days=120)
            #----------------
            filtered_df = ARM.instrument_nfo[
                    (ARM.instrument_nfo['expiry'] >= ARM.todays_server_date) &
                    (ARM.instrument_nfo['expiry'] <= Total_120Days_Exp)]
            ARM.instrument_nfo = filtered_df.reset_index(drop=True)
            #ARM.instrument_nfo.to_csv("Instrument_NFO.csv", index=False)

            ARM.fut_exchange_nfo = ARM.instrument_nfo[ARM.instrument_nfo["segment"] == "NFO-FUT"]
            ARM.fut_exchange_nfo.reset_index(inplace=True)
            ARM.opt_exchange_nfo = ARM.instrument_nfo[ARM.instrument_nfo["segment"] == "NFO-OPT"]
            ARM.FNO_Symbol = ({"FNO_Symbol":list(ARM.opt_exchange_nfo["name"].unique())})['FNO_Symbol']
            ARM.opt_exchange_nfo.reset_index(inplace=True)
            #----------------
            ARM.df_opt = copy.deepcopy(ARM.opt_exchange_nfo)
            ARM.df_opt.drop(columns=["last_price","tick_size","segment","exchange"], inplace=True)
            #ARM.df_opt.drop(columns=["last_price","tick_size","lot_size","segment","exchange"], inplace=True)
            #-----------------------------------
            R1C2entry.destroy()

            if ARM.Ready_Once:
                #--Set Current OPT Expiry of NIFTY ------------------
                ARM.opt_expiries_list = []
                R1C2options = []
                df = copy.deepcopy(ARM.opt_exchange_nfo)
                df = df[df["name"] == "NIFTY"]
                ARM.opt_expiries_list = sorted(list(df["expiry"].unique()))
                for i in range(len(ARM.opt_expiries_list)):
                    R1C2options.append(str(ARM.opt_expiries_list[i]))
                    if (i >= 4):break
                R1C2entry = tk.OptionMenu(root,R1C2,*R1C2options,command=OPT_EXP_R1C2)
                R1C2entry.grid(row=1, column=2,sticky='nsew')
                R1C2entry.config(fg="blue",bg="white",width=8,height=1, borderwidth=1)
                R1C2.set(R1C2options[0])
                NF_FUT_SYMBOL()
                #--------------Websocket--------------------
                if ARM.oc_symbol == "NIFTY":
                    ARM.KWS_Subscribe_Tokens = {int(ARM.SPT_Token): {'symbol': "NIFTY 50"} }
                    ARM.New_SYM_Tokens = {  int(ARM.SPT_Token): {'symbol': "NIFTY 50"} }
                else:pass
                # ARM.KWS_Subscribe_Tokens = {
                #                             256265: {'symbol': "NIFTY 50"},
                #                             260105: {'symbol': "NIFTY BANK"},
                #                             257801: {'symbol': "NIFTY FIN SERVICE"}
                #                         }

                # print("================== UND TOKENS ==================")
                # print(ARM.KWS_Subscribe_Tokens)
                # print(ARM.KWS_Subscribe_Tokens.keys())           #--Prints dict_keys([256265, 260105, 257801])
                # print(list(ARM.KWS_Subscribe_Tokens.keys()))     #--Prints [256265, 260105, 257801]
                global kws_nse
                kws_nse = kite.kws()                      #--Websocket for Underlaying (SPOT/FUTURE)
                kws_nse.on_ticks = on_ticks_ARMKWS        #--Tick Data comes
                kws_nse.on_connect = on_connect_ARMKWS    #--Subscribe & Set Mode
                kws_nse.connect(threaded=True)            #--Start Websocket for Underlaying (SPOT/FUTURE)
                #print("KWS Connection started")
                time.sleep(1)
                #----------------------------------------
                ARM.Ready_Once = False
            else:pass

            if(GUI.Trade_Flag == 'DEMO_REAL_BOTH'):R1C3entry.configure(state=DISABLED,disabledbackground="white",disabledforeground="black")

            OTM_Blink_Status()
            root.after(0, APH_MAIN)

            R0C0entry.destroy()
            R0C0options =[ARM.your_user_id,"Settings","Logout"]
            R0C0.set(R0C0options[0])
            R0C0entry = tk.OptionMenu(root, R0C0,*R0C0options,command = R0C0Pressed)
            R0C0entry.config(fg="blue",bg="white",width=8,height=2)
            R0C0entry.grid(row=0,column=0,rowspan=2)

            print("System is Ready, Opening Balance: {:.2f}, Live Balance: {:.2f}".format(ARM.Opening_Balance,ARM.Live_Balance))

            print("{}, Trading in {} Trade Mode"
                    .format(time.strftime("%H:%M:%S", time.localtime()),GUI.R0C5flag))

            return
        else:root.after(1000,TT_Login)
    else:root.after(1000,TT_Login)

def R0C1Pressed():      #--MIS/NRML Selection
    if(GUI.R0C1flag == 'NRML'):
        R0C1.configure(text='NRML',bg="#00E600")
        GUI.R0C1flag = 'NRML'
        ARM.Buy_MIS_NRML='NRML'
    else:
        R0C1.configure(text='NRML',bg="#00E600")
        GUI.R0C1flag = 'NRML'
        ARM.Buy_MIS_NRML='NRML'
    R0C5.focus_set()
def R1C1Pressed():      #--MIS/NRML Selection
    if(GUI.R1C1flag == 'LIMIT'):
        R1C1.configure(text='LIMIT',bg="#00E600")
        GUI.R1C1flag = 'LIMIT'
        ARM.Buy_MARKET_LIMIT='LIMIT'
    else:
        R1C1.configure(text='LIMIT',bg="#00E600")  #--#FF6666
        GUI.R1C1flag = 'LIMIT'
        ARM.Buy_MARKET_LIMIT='LIMIT'
    R0C5.focus_set()

def OPT_EXP_R1C2(event):    #--NIFTY Option Expiry Seceltion
    # print("----------NIFTY OPT CHANGED-------------")
    APH_MAIN(True)  #--Stop
    if (ARM.APH_SharedPVar[0] == "NIFTY") and (ARM.APH_SharedPVar[1] == "NIFTY"):
        ARM.oc_opt_expiry = datetime.strptime(R1C2.get(), '%Y-%m-%d').date()
    else:pass
    root.after(0,APH_MAIN) #--Start
    R0C5.focus_set()
    return

def R1C3Pressed():         #--Program Reads Opening Balance after Enter key
        if ((R0C0.get() == ARM.your_user_id) and (GUI.R0C5flag == "DEMO")):
            try:
                GUI.R1C3_Value = float(R1C3.get())
            except:
                GUI.R1C3_Value = 0.00
            GUI.R1C4_Value = GUI.R1C3_Value
            ARM.Opening_Balance = GUI.R1C3_Value
            ARM.Live_Balance = GUI.R1C4_Value
            R1C3.set("{:.2f}".format(GUI.R1C3_Value))
            R1C4.set("{:.2f}".format(GUI.R1C4_Value))
            R1C3entry.configure(state=DISABLED,disabledbackground="white",disabledforeground="black",borderwidth=1)
        else:
            R1C3.set("{:.2f}".format(GUI.R1C3_Value))
            R1C4.set("{:.2f}".format(GUI.R1C4_Value))
        R0C5.focus_set()
def Deadband_Pressed():
        if ((R0C0.get() == ARM.your_user_id) and (GUI.R0C5flag == "DEMO")):
            temp = ARM.UDS_DEADBAND
            try:
                ARM.UDS_DEADBAND = int(R1C8.get())
            except:
                ARM.UDS_DEADBAND = temp
        else:pass
        R1C8.set(ARM.UDS_DEADBAND)
        print(f"Deadband Modified from {temp} to {ARM.UDS_DEADBAND}")
        R0C5.focus_set()
def Min_Diff_Pressed():
        if ((R0C0.get() == ARM.your_user_id) and (GUI.R0C5flag == "DEMO")):
            temp = ARM.STD_STR_MIN_diff
            try:
                ARM.STD_STR_MIN_diff = float(R1C9.get())
            except:
                ARM.STD_STR_MIN_diff = temp
        else:pass
        R1C9.set(ARM.STD_STR_MIN_diff)
        print(f"Min Diff Modified from {temp} to {ARM.STD_STR_MIN_diff}")
        R0C5.focus_set()
def Min_Bars_Pressed():
        if ((R0C0.get() == ARM.your_user_id) and (GUI.R0C5flag == "DEMO")):
            temp = ARM.CONSECUTIVE_BAR_value
            try:
                ARM.CONSECUTIVE_BAR_value = int(R1C10.get())
            except:
                ARM.CONSECUTIVE_BAR_value = temp
        else:pass
        R1C10.set(ARM.CONSECUTIVE_BAR_value)
        print(f"Min Bars Modified from {temp} to {ARM.CONSECUTIVE_BAR_value}")
        R0C5.focus_set()
def LPT_Pressed():         #--Program Reads Opening Balance after Enter key
        if ((R0C0.get() == ARM.your_user_id) and (GUI.R0C5flag == "DEMO")):
            try:
                GUI.LPT_Value = int(R1C6.get())
            except:
                GUI.LPT_Value = 250
        else:pass
        R1C6.set(GUI.LPT_Value)
        R0C5.focus_set()
def PPT_Pressed():         #--Program Reads Opening Balance after Enter key
        if ((R0C0.get() == ARM.your_user_id) and (GUI.R0C5flag == "DEMO")):
            try:
                GUI.PPT_Value = int(R1C7.get())
            except:
                GUI.PPT_Value = 750
        else:pass
        R1C7.set(GUI.PPT_Value)
        R0C5.focus_set()
def R3C2Pressed():          #--Bid CE Qty
    try:
        GUI.R3C2_Value = R3C2.get()
        R3C2.set("{:.0f}".format(GUI.R3C2_Value))
    except:
        GUI.R3C2_Value = 0
        R3C2.set("{:.0f}".format(GUI.R3C2_Value))
    R0C5.focus_set()
def R3C4Pressed():          #--Ask CE Qty
    try:
        GUI.R3C4_Value = R3C4.get()
        R3C4.set("{:.0f}".format(GUI.R3C4_Value))
    except:
        GUI.R3C4_Value = 0
        R3C4.set("{:.0f}".format(GUI.R3C4_Value))
    R0C5.focus_set()
def R3C6Pressed():         #--Ask PE Qty
    try:
        GUI.R3C6_Value = R3C6.get()
        R3C6.set("{:.0f}".format(GUI.R3C6_Value))
    except:
        GUI.R3C6_Value = 0
        R3C6.set("{:.0f}".format(GUI.R3C6_Value))
    R0C5.focus_set()
def R3C8Pressed():         #--Bid PE Qty
    try:
        GUI.R3C8_Value = R3C8.get()
        R3C8.set("{:.0f}".format(GUI.R3C8_Value))
    except:
        GUI.R3C8_Value = 0
        R3C8.set("{:.0f}".format(GUI.R3C8_Value))
    R0C5.focus_set()
def R0C5Pressed():         #--DEMO/REAL Selection
    if (R0C0.get() == ARM.your_user_id) and (ARM.License_Type == "R"):
        Want_ChangePR = POPUP_YESNO("Trading Mode Confirmation", "Are you sure you want to change Mode?\n\nIf change Mode,\n all running Trade Information will get Deleted")
        if Want_ChangePR:
            if(GUI.R0C5flag == "DEMO"):
                R0C5.configure(image=reald_image,borderwidth=0)
                GUI.R0C5flag = "REAL"
            elif(GUI.R0C5flag == "REAL"):
                R0C5.configure(image=demod_image,borderwidth=0)
                GUI.R0C5flag = "DEMO"
            else:pass
            ARM.MY_TRADE = GUI.R0C5flag
            ARM.Opt_Bought1_df = ARM.Opt_Bought1_df.drop(ARM.Opt_Bought1_df.index)
            ARM.Opt_Bought2_df = ARM.Opt_Bought2_df.drop(ARM.Opt_Bought2_df.index)
            ARM.Opt_Bought3_df = ARM.Opt_Bought3_df.drop(ARM.Opt_Bought3_df.index)
            ARM.Opt_Sold1_df = ARM.Opt_Sold1_df.drop(ARM.Opt_Sold1_df.index)
            ARM.Opt_Sold2_df = ARM.Opt_Sold2_df.drop(ARM.Opt_Sold2_df.index)
            ARM.Opt_Sold3_df = ARM.Opt_Sold3_df.drop(ARM.Opt_Sold3_df.index)
            ARM.Display_NetBuySell_df = ARM.Display_NetBuySell_df.drop(ARM.Display_NetBuySell_df.index)
            ARM.NetBuySell_len = len(ARM.Display_NetBuySell_df)
            ARM.Option_Bought = {}
            ARM.Option_Sold = {}
            ARM.Option_NetBuySell = {}
            # Desplay_NetPosition()
            # print(ARM.Display_NetBuySell_df)
            print("{}, Trading in {} Trade Mode"
                    .format(time.strftime("%H:%M:%S", time.localtime()),GUI.R0C5flag))
        else:pass
        R0C5.focus_set()
    else:pass
def R2C4Pressed():         #--Refresh Trade Pressed
    if(GUI.R2C4flag == "RUN") and (R0C0.get() == ARM.your_user_id):
        if (ARM.NetBuySell_len>0):
            ARM.Stop_Trade_Update = True
            GUI.R2C4flag = "RESET"
            R2C4.configure(fg="red")
            # print("Refreshing Trade Display......")
        else:pass
        R26C2["text"] = R26C4["text"] = R26C6["text"] = R26C8["text"] = ""
        R26C2["bg"] = R26C4["bg"] = R26C6["bg"] = R26C8["bg"] = "#FCDFFF"
    else:pass
    R0C5.focus_set()
def R0C0Pressed(event):          #--Setting / Logout Option Menu selection
    if(R0C0.get() == "Logout"):root.after(0,POPUP_LOGOUT)
    elif(R0C0.get() == "Settings"):root.after(0,POPUP_SETTING)
    elif(R0C0.get() == ARM.your_user_id):R0C0.set(R0C0options[0])
    else:pass
    R0C5.focus_set()

def Close_OTM_Direct():
    if (R0C0.get() == "Logout"):
        POPUP_LOGOUT()
    else:
        # print("{0} Opening Balance: {1}, Live Balance: {2}".format(ARM.your_user_id,ARM.Opening_Balance,ARM.Live_Balance))
        print("-------------------------------------------------------------------------------------") #--85 White Spaces
        print("Trading Session End                           {}".format(datetime.now().strftime("%H:%M:%S, %d %b %Y, %a")))
        print("-------------------------------------------------------------------------------------")
        if(len(ARM.My_AllTrades)>0):
            filename = f"My_AllTrades_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            filepath = os.path.join(os.getcwd(),'Log', filename)
            ARM.My_AllTrades.to_csv(filepath, index=False)
        else:pass
        root.destroy()
        sys.stdout.close()
        if (os.path.exists(TxtLogFilePath)):
            convert_txt_to_pdf(TxtLogFilePath, PdfLogFilePath)
            os.remove(TxtLogFilePath)
        else:pass
        sys.exit()
def Update_Popup_OC():
    #--------------------------------------------------------------------
    if (ARM.OC_Strikes != 10) and (ARM.Display_Row != 5):
        for i in range(5,ARM.Display_Row): #--Try to Optimise these for loops
            for j in range(0,19):
                OC_Cell[i][j]["text"] = ""
                OC_Cell[i][j]["bg"] = "#F0F0F0"

        for i in range(ARM.Display_Row+(2*ARM.OC_Strikes)+1,26):
            for j in range(0,19):
                OC_Cell[i][j]["text"] = ""
                OC_Cell[i][j]["bg"] = "#F0F0F0"
    else:pass

    root.title(f"{ARM.oc_symbol} Option Chain")

    # print(ARM.df_opt_final)
    num_rows, num_columns = ARM.df_opt_final.shape
    # print(f"Number of rows: {num_rows}")
    # print(f"Number of columns: {num_columns}")

    for i in range(ARM.Display_Row,num_rows+ARM.Display_Row):   #--DISPLAY / UPDATE OPTION CHAIN (Row)
        for j in range(0,num_columns):                        #--DISPLAY / UPDATE OPTION CHAIN (Col)
            if j in [0,1,2,3,4,6,7,8,9,10]:
                value = "{:.2f}".format(ARM.df_opt_final.iloc[i-ARM.Display_Row, j])
            else:
                value = str(ARM.df_opt_final.iloc[i-ARM.Display_Row, j])

            if j in [0,10]:
                if ((ARM.df_opt_final.iloc[i-ARM.Display_Row, j]) > 0):
                    OC_Cell[i][j].config(fg="#006600")
                elif ((ARM.df_opt_final.iloc[i-ARM.Display_Row, j]) < 0):
                    OC_Cell[i][j].config(fg="#CC0000")
                else:
                    OC_Cell[i][j].config(fg="black")

            OC_Cell[i][j].config(text=value)

            if (i < 11 and j < 5):OC_Cell[i][j].config(bg="#D6D6D2")
            if (10 < i < 15 and j < 5):OC_Cell[i][j].config(bg="#BED9F0")
            if (20 > i > 15 and j < 5):OC_Cell[i][j].config(bg="#DDEBF7")
            if (i > 19 and j < 5):OC_Cell[i][j].config(bg="#F0F0F0")

            if  (i==15) or (j == 5):OC_Cell[i][j].config(bg="yellow")

            if (i < 11 and j > 5):OC_Cell[i][j].config(bg="#F0F0F0")
            if (10 < i < 15 and j > 5):OC_Cell[i][j].config(bg="#DDEBF7")
            if (20 > i > 15 and j > 5):OC_Cell[i][j].config(bg="#BED9F0")
            if (i > 19 and j > 5):OC_Cell[i][j].config(bg="#D6D6D2")
            #-------------Simple Solution-1-------------------------------------------
            # if (ARM.Bought_Strike != 0):
            #     index_of_bought_strike = ARM.df_opt_final.index[ARM.df_opt_final["STRIKE"] == ARM.Bought_Strike].tolist()
            #     #print("Bought Strike index is:",index_of_bought_strike)
            #     if(ARM.Bought_CE_PE == "CE"):
            #         OC_Cell[index_of_bought_strike[0]+5][7].config(bg="#FF6666")
            #     else:
            #         OC_Cell[index_of_bought_strike[0]+5][13].config(bg="#FF6666")
            # else:pass
            #-------------Simple Solution-2-------------------------------------------
            # if ARM.Option_Bought:
            #     Symbol_CE_Strikes = [entry[2] for entry in ARM.Option_Bought if entry[0] == "BUY" and entry[1] == ARM.oc_symbol and entry[3] == 'CE']
            #     Symbol_PE_Strikes = [entry[2] for entry in ARM.Option_Bought if entry[0] == "BUY" and entry[1] == ARM.oc_symbol and entry[3] == 'PE']
            #     for strike in Symbol_CE_Strikes:
            #         index_of_bought_strike = ARM.df_opt_final.index[ARM.df_opt_final["STRIKE"] == strike].tolist()
            #         #OC_Cell[index_of_bought_strike[0]+5][7].config(bg="#FF6666")
            #         CE_Qty_to_search = ("BUY",ARM.oc_symbol, strike, "CE")
            #         for key in ARM.Option_Bought:
            #             if key == CE_Qty_to_search:
            #                 CEDisplay_Qty = ARM.Option_Bought[key][4]
            #         OC_Cell[index_of_bought_strike[0]+5][7].config(text="{:.2f} ({})".format(ARM.df_opt_final.iloc[index_of_bought_strike[0], 5],int(CEDisplay_Qty/ARM.found_lotsize_opt)),bg="#FF6666")

            #     for strike in Symbol_PE_Strikes:
            #         index_of_bought_strike = ARM.df_opt_final.index[ARM.df_opt_final["STRIKE"] == strike].tolist()
            #         #OC_Cell[index_of_bought_strike[0]+5][13].config(bg="#FF6666")
            #         PE_Qty_to_search = ("BUY",ARM.oc_symbol, strike, "PE")
            #         for key in ARM.Option_Bought:
            #             if key == PE_Qty_to_search:
            #                 PEDisplay_Qty = ARM.Option_Bought[key][4]
            #         OC_Cell[index_of_bought_strike[0]+5][13].config(text="{:.2f} ({})".format(ARM.df_opt_final.iloc[index_of_bought_strike[0], 11],int(PEDisplay_Qty/ARM.found_lotsize_opt)),bg="#FF6666")
            # else:pass
            #-------------Perfect Solution-3-------------------------------------------
            #print(ARM.Option_NetBuySell)
            if ARM.Option_NetBuySell:
                Symbol_CE_BuyStrikes = [entry[2] for entry in ARM.Option_NetBuySell if entry[0] == "BUY" and entry[1] == ARM.oc_symbol and entry[3] == 'CE' and entry[4] == ARM.oc_opt_expiry ]
                Symbol_PE_BuyStrikes = [entry[2] for entry in ARM.Option_NetBuySell if entry[0] == "BUY" and entry[1] == ARM.oc_symbol and entry[3] == 'PE' and entry[4] == ARM.oc_opt_expiry]

                Symbol_CE_SellStrikes = [entry[2] for entry in ARM.Option_NetBuySell if entry[0] == "SELL" and entry[1] == ARM.oc_symbol and entry[3] == 'CE' and entry[4] == ARM.oc_opt_expiry]
                Symbol_PE_SellStrikes = [entry[2] for entry in ARM.Option_NetBuySell if entry[0] == "SELL" and entry[1] == ARM.oc_symbol and entry[3] == 'PE' and entry[4] == ARM.oc_opt_expiry]

                for strike in Symbol_CE_BuyStrikes:
                    index_of_bought_strike = ARM.df_opt_final.index[ARM.df_opt_final["STRIKE"] == strike].tolist()
                    CE_Qty_to_search = ("BUY",ARM.oc_symbol, strike, "CE",ARM.oc_opt_expiry)
                    for key in ARM.Option_NetBuySell:
                        if key == CE_Qty_to_search:
                            CEDisplay_Qty = ARM.Option_NetBuySell[key][5]
                    OC_Cell[index_of_bought_strike[0]+ARM.Display_Row][2].config(text="{:.2f} ({})".format(ARM.df_opt_final.iloc[index_of_bought_strike[0], 4],int(CEDisplay_Qty/ARM.found_lotsize_opt)),bg="#FF6666")

                for strike in Symbol_PE_BuyStrikes:
                    index_of_bought_strike = ARM.df_opt_final.index[ARM.df_opt_final["STRIKE"] == strike].tolist()
                    PE_Qty_to_search = ("BUY",ARM.oc_symbol, strike, "PE" ,ARM.oc_opt_expiry)
                    for key in ARM.Option_NetBuySell:
                        if key == PE_Qty_to_search:
                            PEDisplay_Qty = ARM.Option_NetBuySell[key][5]
                    OC_Cell[index_of_bought_strike[0]+ARM.Display_Row][8].config(text="{:.2f} ({})".format(ARM.df_opt_final.iloc[index_of_bought_strike[0], 6],int(PEDisplay_Qty/ARM.found_lotsize_opt)),bg="#FF6666")

                for strike in Symbol_CE_SellStrikes:
                    index_of_bought_strike = ARM.df_opt_final.index[ARM.df_opt_final["STRIKE"] == strike].tolist()
                    CE_Qty_to_search = ("SELL",ARM.oc_symbol, strike, "CE" ,ARM.oc_opt_expiry)
                    for key in ARM.Option_NetBuySell:
                        if key == CE_Qty_to_search:
                            CEDisplay_Qty = ARM.Option_NetBuySell[key][5]
                    OC_Cell[index_of_bought_strike[0]+ARM.Display_Row][4].config(text="{:.2f} ({})".format(ARM.df_opt_final.iloc[index_of_bought_strike[0], 2],int(CEDisplay_Qty/ARM.found_lotsize_opt)),bg="#00E600")

                for strike in Symbol_PE_SellStrikes:
                    index_of_bought_strike = ARM.df_opt_final.index[ARM.df_opt_final["STRIKE"] == strike].tolist()
                    PE_Qty_to_search = ("SELL",ARM.oc_symbol, strike, "PE" ,ARM.oc_opt_expiry)
                    for key in ARM.Option_NetBuySell:
                        if key == PE_Qty_to_search:
                            PEDisplay_Qty = ARM.Option_NetBuySell[key][5]
                    OC_Cell[index_of_bought_strike[0]+ARM.Display_Row][6].config(text="{:.2f} ({})".format(ARM.df_opt_final.iloc[index_of_bought_strike[0], 8],int(PEDisplay_Qty/ARM.found_lotsize_opt)),bg="#00E600")
            else:pass

def on_SL_Target_enter(event, row, column):
    ARM.Edit_SL_Target = True

    # print("----Try to Pause Trade Update----")
def on_SL_Target_leave(event, row, column):
    pass
    # ARM.Edit_SL_Target = False
    # print("----started----")
def on_SL_Target_click(event, row, column):
    ARM.Edit_SL_Target = True
def on_SL_Target_Pressed(event, row, column):
    if(column == 8):
        try:
            SET_SL = OC_Cell_Value[row][column].get()
            OC_Cell_Value[row][column].set("{:.0f}".format(SET_SL))
        except:
            SET_SL = -100.00
            OC_Cell_Value[row][column].set("{:.0f}".format(SET_SL))

        ARM.Display_NetBuySell_df.loc[row-28,'SL(P)'] = SET_SL

    elif(column == 9):
        try:
            SET_TARGET = OC_Cell_Value[row][column].get()
            OC_Cell_Value[row][column].set("{:.0f}".format(SET_TARGET))
        except:
            SET_TARGET = 100.00
            OC_Cell_Value[row][column].set("{:.0f}".format(SET_TARGET))

        ARM.Display_NetBuySell_df.loc[row-28,'Target(P)'] = SET_TARGET

    else:pass

    for i in range(ARM.NetBuySell_len):
        for J in range(len(ARM.Opt_Bought1_df)):
            if ((ARM.Display_NetBuySell_df.loc[i,'Action'] == ARM.Opt_Bought1_df.loc[J,'Action']) and
                (ARM.Display_NetBuySell_df.loc[i,'Symbol'] == ARM.Opt_Bought1_df.loc[J,'Symbol'])):
                ARM.Opt_Bought1_df.loc[J,'SL(P)'] = ARM.Display_NetBuySell_df.loc[i,'SL(P)']
                ARM.Opt_Bought1_df.loc[J,'Target(P)'] = ARM.Display_NetBuySell_df.loc[i,'Target(P)']
            else:pass

    ARM.Edit_SL_Target = False
    # Desplay_NetPosition()
    R0C5.focus_set()
#----------------------------------------------------
def on_enter(event, row, column):
    #if (R0C0.get() == "Logout"):
    if (R0C0.get() == ARM.your_user_id) and (ARM.Display_Row <= row <= ARM.Display_Row+len(ARM.df_opt_final)-1):
        event.widget.config(relief="solid",borderwidth=1)
        event.widget.config(cursor="hand2")
    else:pass
def on_leave(event, row, column):
    if (R0C0.get() == ARM.your_user_id):
        event.widget.config(relief="flat",borderwidth=1)
        event.widget.config(cursor="")
        R3C3.config(text="<== Qty ==>", anchor="center")
        R3C7.config(text="<== Qty ==>", anchor="center")
    else:pass
def on_buy_click(event, row, column):
    # print(f"row : {row}, column : {column}")
    if (ARM.order_id_buy is None) and (ARM.order_id_sell is None):
        R26C2["text"] = R26C4["text"] = R26C6["text"] = R26C8["text"] = ""
        R26C2["bg"] = R26C4["bg"] = R26C6["bg"] = R26C8["bg"] = "#FCDFFF"
        # try:
        #     winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
        # except:pass
        OC_Cell[row][column]["bg"] = "yellow"
        R2C8.set("Processing Buy Order")
        R2C8entry.config(fg="white",bg="green")
        if ARM.ErrorExcptionInfo_scheduled is None: ARM.ErrorExcptionInfo_scheduled = root.after(1000,Clear_ErrorExcptionInfo)
        #if (R0C0.get() == ARM.your_user_id) and (R1C4.get() > 0) and (GUI.R0C5flag == "REAL") and (ARM.APH_SharedPVar[1] == ARM.oc_symbol):
        #if (R0C0.get() == ARM.your_user_id) and (GUI.R0C5flag == "REAL"):#--This is for testing only, comment after testing
        if (R0C0.get() == ARM.your_user_id):
            if (R1C4.get() > 0):
                if (ARM.APH_SharedPVar[1] == ARM.oc_symbol):
                    # value = event.widget.cget("text")
                    # print(f"Clicked on '{value}' at Row {row}, Column {column}")
                    if (column==4):  #-- CE
                        R3C4Pressed()
                        ARM.Buy_CE_PE = "CE"
                        ARM.Buy_AskPrice = ARM.df_opt_final.loc[row-ARM.Display_Row, "CE_Ask"]
                        ARM.Buy_Strike = ARM.df_opt_final.loc[row-ARM.Display_Row, "STRIKE"]
                        ARM.Buy_MIS_NRML = GUI.R0C1flag
                        ARM.Buy_MARKET_LIMIT = GUI.R1C1flag
                        ARM.Buy_Qty = ARM.found_lotsize_opt*GUI.R3C4_Value
                        SYMBOL_TO_TRADE(ARM.Buy_Strike,ARM.Buy_CE_PE)

                        ARM.Buy_Symbol = ARM.NF_OPT_Trade_Symbol

                        existing_trade = any(ARM.Display_NetBuySell_df["Symbol"] == ARM.Buy_Symbol)
                        if (ARM.NetBuySell_len < 5 or existing_trade):
                            Auto_Buy_Status()
                            APH_BUY()
                        else:
                            print("M A X I M U M   T R A D E   L I M I T   I S   :   5")
                    elif (column==6):   #-- PE
                        R3C6Pressed()
                        ARM.Buy_CE_PE = "PE"
                        ARM.Buy_AskPrice = ARM.df_opt_final.loc[row-ARM.Display_Row, "PE_Ask"]
                        ARM.Buy_Strike = ARM.df_opt_final.loc[row-ARM.Display_Row, "STRIKE"]
                        ARM.Buy_MIS_NRML = GUI.R0C1flag
                        ARM.Buy_MARKET_LIMIT = GUI.R1C1flag
                        ARM.Buy_Qty = ARM.found_lotsize_opt*GUI.R3C6_Value
                        SYMBOL_TO_TRADE(ARM.Buy_Strike,ARM.Buy_CE_PE)

                        ARM.Buy_Symbol = ARM.NF_OPT_Trade_Symbol

                        existing_trade = any(ARM.Display_NetBuySell_df["Symbol"] == ARM.Buy_Symbol)
                        if (ARM.NetBuySell_len < 5 or existing_trade):
                            Auto_Buy_Status()
                            APH_BUY()
                        else:
                            print("M A X I M U M   T R A D E   L I M I T   I S   :   5")
                    else:
                        print("Something went worong in On Buy Click")
                    #ARM.Buy_CE_PE = "" #--This is for testing only, comment after testing
                else:
                    R2C8.set("Wait to change Symbol")
                    R2C8entry.config(fg="white",bg="red")
                    if ARM.ErrorExcptionInfo_scheduled is None: ARM.ErrorExcptionInfo_scheduled = root.after(3000,Clear_ErrorExcptionInfo)
            else:
                R2C8.set("Your Live Balance is Zero")
                R2C8entry.config(fg="white",bg="red")
                if ARM.ErrorExcptionInfo_scheduled is None: ARM.ErrorExcptionInfo_scheduled = root.after(3000,Clear_ErrorExcptionInfo)
        else:
            R2C8.set("You have not Logged-in")
            R2C8entry.config(fg="white",bg="red")
            if ARM.ErrorExcptionInfo_scheduled is None: ARM.ErrorExcptionInfo_scheduled = root.after(3000,Clear_ErrorExcptionInfo)
        #OC_Cell[row][column]["bg"] = "#F0F0F0"
    else:
        R2C8.set("Processing Previous Order")
        R2C8entry.config(fg="white",bg="red")
        if ARM.ErrorExcptionInfo_scheduled is None: ARM.ErrorExcptionInfo_scheduled = root.after(3000,Clear_ErrorExcptionInfo)
        print("Wait to Process Previous Order")
        #pass
def on_sell_click(event, row, column):
    if (ARM.order_id_buy is None) and (ARM.order_id_sell is None):
        R26C2["text"] = R26C4["text"] = R26C6["text"] = R26C8["text"] = ""
        R26C2["bg"] = R26C4["bg"] = R26C6["bg"] = R26C8["bg"] = "#FCDFFF"
        # try:
        #     winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
        # except:pass
        OC_Cell[row][column]["bg"] = "yellow"
        R2C8.set("Processing Sell Order")
        R2C8entry.config(fg="white",bg="green")
        if ARM.ErrorExcptionInfo_scheduled is None: ARM.ErrorExcptionInfo_scheduled = root.after(1000,Clear_ErrorExcptionInfo)
        #if (R0C0.get() == ARM.your_user_id) and (R1C4.get() > 0) and (GUI.R0C5flag == "REAL") and (ARM.APH_SharedPVar[1] == ARM.oc_symbol):
        #if (R0C0.get() == ARM.your_user_id) and (GUI.R0C5flag == "REAL"): #--This is for testing only, comment after testing
        if (R0C0.get() == ARM.your_user_id):
            if (R1C4.get() > 0):
                if (ARM.APH_SharedPVar[1] == ARM.oc_symbol):
                    # value = event.widget.cget("text")
                    # print(f"Clicked on '{value}' at Row {row}, Column {column}")
                    if (column==2):      #--CE
                        R3C2Pressed()
                        ARM.Sell_CE_PE = "CE"
                        ARM.Sell_BidPrice = ARM.df_opt_final.loc[row-ARM.Display_Row, "CE_Bid"]
                        ARM.Sell_Strike = ARM.df_opt_final.loc[row-ARM.Display_Row, "STRIKE"]
                        ARM.Sell_MIS_NRML = GUI.R0C1flag
                        ARM.Sell_MARKET_LIMIT = GUI.R1C1flag
                        ARM.Sell_Qty = ARM.found_lotsize_opt*GUI.R3C2_Value
                        SYMBOL_TO_TRADE(ARM.Sell_Strike,ARM.Sell_CE_PE)

                        ARM.Sell_Symbol = ARM.NF_OPT_Trade_Symbol

                        existing_trade = any(ARM.Display_NetBuySell_df["Symbol"] == ARM.Sell_Symbol)
                        if (ARM.NetBuySell_len < 5 or existing_trade):
                            APH_SELL()
                        else:
                            print("M A X I M U M   T R A D E   L I M I T   I S   :   5")
                    elif (column==8):   #--PE
                        R3C8Pressed()
                        ARM.Sell_CE_PE = "PE"
                        ARM.Sell_BidPrice = ARM.df_opt_final.loc[row-ARM.Display_Row, "PE_Bid"]
                        ARM.Sell_Strike = ARM.df_opt_final.loc[row-ARM.Display_Row, "STRIKE"]
                        ARM.Sell_MIS_NRML = GUI.R0C1flag
                        ARM.Sell_MARKET_LIMIT = GUI.R1C1flag
                        ARM.Sell_Qty = ARM.found_lotsize_opt*GUI.R3C8_Value
                        SYMBOL_TO_TRADE(ARM.Sell_Strike,ARM.Sell_CE_PE)

                        ARM.Sell_Symbol = ARM.NF_OPT_Trade_Symbol

                        existing_trade = any(ARM.Display_NetBuySell_df["Symbol"] == ARM.Sell_Symbol)
                        if (ARM.NetBuySell_len < 5 or existing_trade):
                            APH_SELL()
                        else:
                            print("M A X I M U M   T R A D E   L I M I T   I S   :   5")
                    else:
                        print("Something went worong in On Sell Click")
                    #ARM.Sell_CE_PE = "" #--This is for testing only, comment after testing
                else:
                    R2C8.set("Wait to change Symbol")
                    R2C8entry.config(fg="white",bg="red")
                    if ARM.ErrorExcptionInfo_scheduled is None: ARM.ErrorExcptionInfo_scheduled = root.after(3000,Clear_ErrorExcptionInfo)
            else:
                R2C8.set("Your Live Balance is Zero")
                R2C8entry.config(fg="white",bg="red")
                if ARM.ErrorExcptionInfo_scheduled is None: ARM.ErrorExcptionInfo_scheduled = root.after(3000,Clear_ErrorExcptionInfo)
        else:
            R2C8.set("You have not Logged-in")
            R2C8entry.config(fg="white",bg="red")
            if ARM.ErrorExcptionInfo_scheduled is None: ARM.ErrorExcptionInfo_scheduled = root.after(3000,Clear_ErrorExcptionInfo)
        #OC_Cell[row][column]["bg"] = "#F0F0F0"
    else:
        R2C8.set("Processing Previous Order")
        R2C8entry.config(fg="white",bg="red")
        if ARM.ErrorExcptionInfo_scheduled is None: ARM.ErrorExcptionInfo_scheduled = root.after(3000,Clear_ErrorExcptionInfo)
        print("Wait to Process Previous Order")
        #pass
def on_sqoff_enter(event, row, column):
    if (R0C0.get() == ARM.your_user_id):
        if (row==28 and column==10) and ARM.NetBuySell_len >= 1:
            event.widget.config(relief="solid",borderwidth=1)
            event.widget.config(cursor="hand2")
        elif (row==29 and column==10) and ARM.NetBuySell_len >= 2:
            event.widget.config(relief="solid",borderwidth=1)
            event.widget.config(cursor="hand2")
        elif (row==30 and column==10) and ARM.NetBuySell_len >= 3:
            event.widget.config(relief="solid",borderwidth=1)
            event.widget.config(cursor="hand2")
        elif (row==31 and column==10) and ARM.NetBuySell_len >= 4:
            event.widget.config(relief="solid",borderwidth=1)
            event.widget.config(cursor="hand2")
        elif (row==32 and column==10) and ARM.NetBuySell_len >= 5:
            event.widget.config(relief="solid",borderwidth=1)
            event.widget.config(cursor="hand2")
        else:pass
    else:pass
def on_sqoff_leave(event, row, column):
    if (R0C0.get() == ARM.your_user_id):
        if (row==28 and column==10):
            event.widget.config(relief="flat",borderwidth=1)
            event.widget.config(cursor="")
        elif (row==29 and column==10):
            event.widget.config(relief="flat",borderwidth=1)
            event.widget.config(cursor="")
        elif (row==30 and column==10):
            event.widget.config(relief="flat",borderwidth=1)
            event.widget.config(cursor="")
        elif (row==31 and column==10):
            event.widget.config(relief="flat",borderwidth=1)
            event.widget.config(cursor="")
        elif (row==32 and column==10):
            event.widget.config(relief="flat",borderwidth=1)
            event.widget.config(cursor="")
        else:pass
    else:pass
def on_sqoff_click(event, row, column):
    #print("Button Clicked: row {}, column {}".format(row, column))
    if (ARM.order_id_buy is None) and (ARM.order_id_sell is None):
        if (ARM.NetBuySell_len > 0) and not ARM.Target_SL_Hit_Flag:
            if(row==28 and column==10) and ARM.NetBuySell_len >= 1:
                if (ARM.Display_NetBuySell_df.loc[0,'Action'] == 'BUY'):
                    ARM.Sell_Symbol_SqOff = ARM.Display_NetBuySell_df.loc[0,'Symbol']
                    ARM.oc_symbol_SqOff, ARM.Sell_Strike_SqOff, ARM.Sell_CE_PE_SqOff, ARM.Sell_Exp_SqOff = extract_info_from_symbol(ARM.Sell_Symbol_SqOff)
                    ARM.Sell_BidPrice_SqOff = ARM.Display_NetBuySell_df.loc[0,'SqOffPrice']
                    ARM.Sell_MIS_NRML_SqOff = GUI.R0C1flag
                    ARM.Sell_MARKET_LIMIT_SqOff = GUI.R1C1flag
                    ARM.Sell_Qty_SqOff = ARM.Display_NetBuySell_df.loc[0,'Total_Qty']
                    print("#----------------------------------------------------------------------------")
                    print(f"{ARM.Sell_Symbol_SqOff} - Immediate Manual Sell (Squiring Off) Bought Position")
                    print("#----------------------------------------------------------------------------")
                    ARM.Target_SL_Hit_Flag = True
                    #APH_SELL()
                    root.after(1,APH_SELL)
                    # try:
                    #     winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
                    # except:pass
                    R2C8.set("Processing 100% Square-Off")
                    R2C8entry.config(fg="white",bg="green")
                    if ARM.ErrorExcptionInfo_scheduled is None: ARM.ErrorExcptionInfo_scheduled = root.after(1000,Clear_ErrorExcptionInfo)
                    return
                elif (ARM.Display_NetBuySell_df.loc[0,'Action'] == 'SELL'):
                    ARM.Buy_Symbol_SqOff = ARM.Display_NetBuySell_df.loc[0,'Symbol']
                    ARM.oc_symbol_SqOff, ARM.Buy_Strike_SqOff, ARM.Buy_CE_PE_SqOff, ARM.Buy_Exp_SqOff = extract_info_from_symbol(ARM.Buy_Symbol_SqOff)
                    ARM.Buy_AskPrice_SqOff = ARM.Display_NetBuySell_df.loc[0,'SqOffPrice']
                    ARM.Buy_MIS_NRML_SqOff = GUI.R0C1flag
                    ARM.Buy_MARKET_LIMIT_SqOff = GUI.R1C1flag
                    ARM.Buy_Qty_SqOff = ARM.Display_NetBuySell_df.loc[0,'Total_Qty']
                    print("#----------------------------------------------------------------------------")
                    print(f"{ARM.Buy_Symbol_SqOff} - Immediate Manual Buy (Squiring Off) Sold Position")
                    print("#----------------------------------------------------------------------------")
                    ARM.Target_SL_Hit_Flag = True
                    #APH_BUY()
                    root.after(1,APH_BUY)
                    # try:
                    #     winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
                    # except:pass
                    R2C8.set("Processing 100% Square-Off")
                    R2C8entry.config(fg="white",bg="green")
                    if ARM.ErrorExcptionInfo_scheduled is None: ARM.ErrorExcptionInfo_scheduled = root.after(1000,Clear_ErrorExcptionInfo)
                    return
                else:pass
            elif(row==29 and column==10) and ARM.NetBuySell_len >= 2:
                if (ARM.Display_NetBuySell_df.loc[1,'Action'] == 'BUY'):
                    ARM.Sell_Symbol_SqOff = ARM.Display_NetBuySell_df.loc[1,'Symbol']
                    ARM.oc_symbol_SqOff, ARM.Sell_Strike_SqOff, ARM.Sell_CE_PE_SqOff, ARM.Sell_Exp_SqOff = extract_info_from_symbol(ARM.Sell_Symbol_SqOff)
                    ARM.Sell_BidPrice_SqOff = ARM.Display_NetBuySell_df.loc[1,'SqOffPrice']
                    ARM.Sell_MIS_NRML_SqOff = GUI.R0C1flag
                    ARM.Sell_MARKET_LIMIT_SqOff = GUI.R1C1flag
                    ARM.Sell_Qty_SqOff = ARM.Display_NetBuySell_df.loc[1,'Total_Qty']
                    print("#----------------------------------------------------------------------------")
                    print(f"{ARM.Sell_Symbol_SqOff} - Immediate Manual Sell (Squiring Off) Bought Position")
                    print("#----------------------------------------------------------------------------")
                    ARM.Target_SL_Hit_Flag = True
                    #APH_SELL()
                    root.after(1,APH_SELL)
                    # try:
                    #     winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
                    # except:pass
                    R2C8.set("Processing 100% Square-Off")
                    R2C8entry.config(fg="white",bg="green")
                    if ARM.ErrorExcptionInfo_scheduled is None: ARM.ErrorExcptionInfo_scheduled = root.after(1000,Clear_ErrorExcptionInfo)
                    return
                elif (ARM.Display_NetBuySell_df.loc[1,'Action'] == 'SELL'):
                    ARM.Buy_Symbol_SqOff = ARM.Display_NetBuySell_df.loc[1,'Symbol']
                    ARM.oc_symbol_SqOff, ARM.Buy_Strike_SqOff, ARM.Buy_CE_PE_SqOff, ARM.Buy_Exp_SqOff = extract_info_from_symbol(ARM.Buy_Symbol_SqOff)
                    ARM.Buy_AskPrice_SqOff = ARM.Display_NetBuySell_df.loc[1,'SqOffPrice']
                    ARM.Buy_MIS_NRML_SqOff = GUI.R0C1flag
                    ARM.Buy_MARKET_LIMIT_SqOff = GUI.R1C1flag
                    ARM.Buy_Qty_SqOff = ARM.Display_NetBuySell_df.loc[1,'Total_Qty']
                    print("#----------------------------------------------------------------------------")
                    print(f"{ARM.Buy_Symbol_SqOff} - Immediate Manual Buy (Squiring Off) Sold Position")
                    print("#----------------------------------------------------------------------------")
                    ARM.Target_SL_Hit_Flag = True
                    #APH_BUY()
                    root.after(1,APH_BUY)
                    # try:
                    #     winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
                    # except:pass
                    R2C8.set("Processing 100% Square-Off")
                    R2C8entry.config(fg="white",bg="green")
                    if ARM.ErrorExcptionInfo_scheduled is None: ARM.ErrorExcptionInfo_scheduled = root.after(1000,Clear_ErrorExcptionInfo)
                    return
                else:pass
            elif(row==30 and column==10) and ARM.NetBuySell_len >= 3:
                if (ARM.Display_NetBuySell_df.loc[2,'Action'] == 'BUY'):
                    ARM.Sell_Symbol_SqOff = ARM.Display_NetBuySell_df.loc[2,'Symbol']
                    ARM.oc_symbol_SqOff, ARM.Sell_Strike_SqOff, ARM.Sell_CE_PE_SqOff, ARM.Sell_Exp_SqOff = extract_info_from_symbol(ARM.Sell_Symbol_SqOff)
                    ARM.Sell_BidPrice_SqOff = ARM.Display_NetBuySell_df.loc[2,'SqOffPrice']
                    ARM.Sell_MIS_NRML_SqOff = GUI.R0C1flag
                    ARM.Sell_MARKET_LIMIT_SqOff = GUI.R1C1flag
                    ARM.Sell_Qty_SqOff = ARM.Display_NetBuySell_df.loc[2,'Total_Qty']
                    print("#----------------------------------------------------------------------------")
                    print(f"{ARM.Sell_Symbol_SqOff} - Immediate Manual Sell (Squiring Off) Bought Position")
                    print("#----------------------------------------------------------------------------")
                    ARM.Target_SL_Hit_Flag = True
                    #APH_SELL()
                    root.after(1,APH_SELL)
                    # try:
                    #     winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
                    # except:pass
                    R2C8.set("Processing 100% Square-Off")
                    R2C8entry.config(fg="white",bg="green")
                    if ARM.ErrorExcptionInfo_scheduled is None: ARM.ErrorExcptionInfo_scheduled = root.after(1000,Clear_ErrorExcptionInfo)
                    return
                elif (ARM.Display_NetBuySell_df.loc[2,'Action'] == 'SELL'):
                    ARM.Buy_Symbol_SqOff = ARM.Display_NetBuySell_df.loc[2,'Symbol']
                    ARM.oc_symbol_SqOff, ARM.Buy_Strike_SqOff, ARM.Buy_CE_PE_SqOff, ARM.Buy_Exp_SqOff = extract_info_from_symbol(ARM.Buy_Symbol_SqOff)
                    ARM.Buy_AskPrice_SqOff = ARM.Display_NetBuySell_df.loc[2,'SqOffPrice']
                    ARM.Buy_MIS_NRML_SqOff = GUI.R0C1flag
                    ARM.Buy_MARKET_LIMIT_SqOff = GUI.R1C1flag
                    ARM.Buy_Qty_SqOff = ARM.Display_NetBuySell_df.loc[2,'Total_Qty']
                    print("#----------------------------------------------------------------------------")
                    print(f"{ARM.Buy_Symbol_SqOff} - Immediate Manual Buy (Squiring Off) Sold Position")
                    print("#----------------------------------------------------------------------------")
                    ARM.Target_SL_Hit_Flag = True
                    #APH_BUY()
                    root.after(1,APH_BUY)
                    # try:
                    #     winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
                    # except:pass
                    R2C8.set("Processing 100% Square-Off")
                    R2C8entry.config(fg="white",bg="green")
                    if ARM.ErrorExcptionInfo_scheduled is None: ARM.ErrorExcptionInfo_scheduled = root.after(1000,Clear_ErrorExcptionInfo)
                    return
                else:pass
            elif(row==31 and column==10) and ARM.NetBuySell_len >= 4:
                if (ARM.Display_NetBuySell_df.loc[3,'Action'] == 'BUY'):
                    ARM.Sell_Symbol_SqOff = ARM.Display_NetBuySell_df.loc[3,'Symbol']
                    ARM.oc_symbol_SqOff, ARM.Sell_Strike_SqOff, ARM.Sell_CE_PE_SqOff, ARM.Sell_Exp_SqOff = extract_info_from_symbol(ARM.Sell_Symbol_SqOff)
                    ARM.Sell_BidPrice_SqOff = ARM.Display_NetBuySell_df.loc[3,'SqOffPrice']
                    ARM.Sell_MIS_NRML_SqOff = GUI.R0C1flag
                    ARM.Sell_MARKET_LIMIT_SqOff = GUI.R1C1flag
                    ARM.Sell_Qty_SqOff = ARM.Display_NetBuySell_df.loc[3,'Total_Qty']
                    print("#----------------------------------------------------------------------------")
                    print(f"{ARM.Sell_Symbol_SqOff} - Immediate Manual Sell (Squiring Off) Bought Position")
                    print("#----------------------------------------------------------------------------")
                    ARM.Target_SL_Hit_Flag = True
                    #APH_SELL()
                    root.after(1,APH_SELL)
                    # try:
                    #     winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
                    # except:pass
                    R2C8.set("Processing 100% Square-Off")
                    R2C8entry.config(fg="white",bg="green")
                    if ARM.ErrorExcptionInfo_scheduled is None: ARM.ErrorExcptionInfo_scheduled = root.after(1000,Clear_ErrorExcptionInfo)
                    return
                elif (ARM.Display_NetBuySell_df.loc[3,'Action'] == 'SELL'):
                    ARM.Buy_Symbol_SqOff = ARM.Display_NetBuySell_df.loc[3,'Symbol']
                    ARM.oc_symbol_SqOff, ARM.Buy_Strike_SqOff, ARM.Buy_CE_PE_SqOff, ARM.Buy_Exp_SqOff = extract_info_from_symbol(ARM.Buy_Symbol_SqOff)
                    ARM.Buy_AskPrice_SqOff = ARM.Display_NetBuySell_df.loc[3,'SqOffPrice']
                    ARM.Buy_MIS_NRML_SqOff = GUI.R0C1flag
                    ARM.Buy_MARKET_LIMIT_SqOff = GUI.R1C1flag
                    ARM.Buy_Qty_SqOff = ARM.Display_NetBuySell_df.loc[3,'Total_Qty']
                    print("#----------------------------------------------------------------------------")
                    print(f"{ARM.Buy_Symbol_SqOff} - Immediate Manual Buy (Squiring Off) Sold Position")
                    print("#----------------------------------------------------------------------------")
                    ARM.Target_SL_Hit_Flag = True
                    #APH_BUY()
                    root.after(1,APH_BUY)
                    # try:
                    #     winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
                    # except:pass
                    R2C8.set("Processing 100% Square-Off")
                    R2C8entry.config(fg="white",bg="green")
                    if ARM.ErrorExcptionInfo_scheduled is None: ARM.ErrorExcptionInfo_scheduled = root.after(1000,Clear_ErrorExcptionInfo)
                    return
                else:pass
            elif(row==32 and column==10) and ARM.NetBuySell_len >= 5:
                if (ARM.Display_NetBuySell_df.loc[4,'Action'] == 'BUY'):
                    ARM.Sell_Symbol_SqOff = ARM.Display_NetBuySell_df.loc[4,'Symbol']
                    ARM.oc_symbol_SqOff, ARM.Sell_Strike_SqOff, ARM.Sell_CE_PE_SqOff, ARM.Sell_Exp_SqOff = extract_info_from_symbol(ARM.Sell_Symbol_SqOff)
                    ARM.Sell_BidPrice_SqOff = ARM.Display_NetBuySell_df.loc[4,'SqOffPrice']
                    ARM.Sell_MIS_NRML_SqOff = GUI.R0C1flag
                    ARM.Sell_MARKET_LIMIT_SqOff = GUI.R1C1flag
                    ARM.Sell_Qty_SqOff = ARM.Display_NetBuySell_df.loc[4,'Total_Qty']
                    print("#----------------------------------------------------------------------------")
                    print(f"{ARM.Sell_Symbol_SqOff} - Immediate Manual Sell (Squiring Off) Bought Position")
                    print("#----------------------------------------------------------------------------")
                    ARM.Target_SL_Hit_Flag = True
                    #APH_SELL()
                    root.after(1,APH_SELL)
                    # try:
                    #     winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
                    # except:pass
                    R2C8.set("Processing 100% Square-Off")
                    R2C8entry.config(fg="white",bg="green")
                    if ARM.ErrorExcptionInfo_scheduled is None: ARM.ErrorExcptionInfo_scheduled = root.after(1000,Clear_ErrorExcptionInfo)
                    return
                elif (ARM.Display_NetBuySell_df.loc[4,'Action'] == 'SELL'):
                    ARM.Buy_Symbol_SqOff = ARM.Display_NetBuySell_df.loc[4,'Symbol']
                    ARM.oc_symbol_SqOff, ARM.Buy_Strike_SqOff, ARM.Buy_CE_PE_SqOff, ARM.Buy_Exp_SqOff = extract_info_from_symbol(ARM.Buy_Symbol_SqOff)
                    ARM.Buy_AskPrice_SqOff = ARM.Display_NetBuySell_df.loc[4,'SqOffPrice']
                    ARM.Buy_MIS_NRML_SqOff = GUI.R0C1flag
                    ARM.Buy_MARKET_LIMIT_SqOff = GUI.R1C1flag
                    ARM.Buy_Qty_SqOff = ARM.Display_NetBuySell_df.loc[4,'Total_Qty']
                    print("#----------------------------------------------------------------------------")
                    print(f"{ARM.Buy_Symbol_SqOff} - Immediate Manual Buy (Squiring Off) Sold Position")
                    print("#----------------------------------------------------------------------------")
                    ARM.Target_SL_Hit_Flag = True
                    #APH_BUY()
                    root.after(1,APH_BUY)
                    # try:
                    #     winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
                    # except:pass
                    R2C8.set("Processing 100% Square-Off")
                    R2C8entry.config(fg="white",bg="green")
                    if ARM.ErrorExcptionInfo_scheduled is None: ARM.ErrorExcptionInfo_scheduled = root.after(1000,Clear_ErrorExcptionInfo)
                    return
                else:pass
            else:pass
        else:pass
    else:
        R2C8.set("Processing Previous Order")
        R2C8entry.config(fg="white",bg="red")
        if ARM.ErrorExcptionInfo_scheduled is None: ARM.ErrorExcptionInfo_scheduled = root.after(3000,Clear_ErrorExcptionInfo)
        print("Wait to Process Previous Order")

def measure_latency_and_jitter(host, count=10):
    latencies = []
    for _ in range(count):  #--Measure latency multiple times
        latency = ping(host)
        if latency is not None:
            latencies.append(latency * 1000)  #--Convert to milliseconds
        else:
            print(f"Ping to {host} failed.")
            return None, None

    avg_latency = np.mean(latencies) if latencies else None #--Calculate average latency
    jitter = np.std(latencies) if latencies else None   #--Calculate jitter (standard deviation of latency)

    return avg_latency, jitter

def Show_RTT_SDD(event):
    if (not ARM.Internet_NotConnected):
        #------------------------------------------------------
        ARM.ZLatency, ZJitter = measure_latency_and_jitter('www.zerodha.com')
        GLatency, GJtter = measure_latency_and_jitter('www.google.com')

        if ARM.ZLatency is not None and ZJitter is not None:#--Latency(RTT), Standard Deviation of the latency(RTT)
            print("{}, Zerodha RTT : {:.2f} ms, SD RTT: {:.2f} ms".format(time.strftime("%H:%M:%S", time.localtime()),ARM.ZLatency,ZJitter))
        else:
            print("          Failed to measure Zerodha RTT & SD-RTT")

        if GLatency is not None and GJtter is not None:#--Latency(RTT), Standard Deviation of the latency(RTT)
            print(f"          Google  RTT : {GLatency:.2f} ms, SD RTT : {GJtter:.2f} ms")
        else:
            print("          Failed to measure Google RTT & SD-RTT")
        #------------------------------------------------------
    else:pass
def Show_Total_Position(event):
    #print("Showing Total Positions")
    popup = tk.Toplevel(root)                   #--Create a Toplevel window
    popup.title("All Trades Data")
    popup.geometry("800x200")                   #--Fixed size for the popup window

    def export_to_csv():                        #--Function to export DataFrame to CSV
        filename = f"My_AllTrades_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        filepath = os.path.join(os.getcwd(),'Log', filename)
        ARM.My_AllTrades.to_csv(filepath, index=False)
        showinfo("Export Success", f"Data exported to {filename}")

    export_button = tk.Button(popup, text="Export to CSV", command=export_to_csv)
    export_button.pack(pady=5, side="top")

    notebook = ttk.Notebook(popup)              #--Add Notebook to Toplevel window
    notebook.pack(fill='both', expand=True)

    frame = ttk.Frame(notebook)                 #--Add a frame for the DataFrame display
    notebook.add(frame, text="All Trades")

    #--Customize Treeview style to set a smaller font
    style = ttk.Style()
    style.configure("Treeview", font=("Helvetica", 8))  #--Set font size for Treeview items
    style.configure("Treeview.Heading", font=("Helvetica", 8), foreground="blue")  #--Set font size for headers

    columns = list(ARM.My_AllTrades.columns)    #--Set up a Treeview with vertical scrollbar
    tree = ttk.Treeview(frame, columns=columns, show="headings")

    # for col in columns:                         #--Configure columns and headers
    #     tree.heading(col, text=col)

    #     #--Calculate max width for the column based on data and set column width
    #     max_width = max(ARM.My_AllTrades[col].astype(str).apply(len).max(), len(col)) * 8
    #     tree.column(col, anchor="center", width=max_width)

    for col in columns:         #--Configure columns and headers
        tree.heading(col, text=col)
        try:                    #--Handle NaN values and ensure max_width is a valid integer
            max_width = max(
                ARM.My_AllTrades[col].fillna('').astype(str).apply(len).max(),  #--Convert NaN to empty string
                len(col)
            ) * 8
        except ValueError:
            max_width = 80      #--Default width if calculation fails
        if math.isnan(max_width):max_width = 80
        tree.column(col, anchor="center", width=int(max_width))  #--Ensure width is an integer

    vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview) #--Add vertical scrollbar
    tree.configure(yscrollcommand=vsb.set)

    tree.grid(row=0, column=0, sticky="nsew")   #--Pack the treeview and scrollbar
    vsb.grid(row=0, column=1, sticky="ns")

    frame.grid_rowconfigure(0, weight=1)        #--Configure grid weights for proper resizing
    frame.grid_columnconfigure(0, weight=1)

    for _, row in ARM.My_AllTrades.iterrows():  #--Insert data into the Treeview
        tree.insert("", "end", values=list(row))

    popup.transient(root)                       #--Make the popup window modal
    popup.grab_set()                            #--Ensure popup is focused

def open_straddle_dialog():
    if ARM.STRADDLE_DIALOG is None:     # Safe check
        ARM.STRADDLE_DIALOG = StraddleDialog()
    else:
        try:
            ARM.STRADDLE_DIALOG.lift()
            # ARM.CONFIRM_BUYCE_Cur = True #--(Dummy Testing Buy CE/PE)
        except:
            ARM.STRADDLE_DIALOG = StraddleDialog()

def cepe_circle_click(event=None):
    # ce_canvas.itemconfig(circle, fill="red")   # Example: change to red
    # pe_canvas.itemconfig(circle, fill="green")   # Example: change to red
    # print("Circle button clicked")
    open_straddle_dialog()

if __name__ == "__main__":
    freeze_support()
    mgr = Manager()
    ns = mgr.Namespace()
    ARM = APH_ARM(ns)
    GUI = APH_GUI()
    #-----------------------------------
    root = tk.Tk()
    root.title('Option Chain Trader ({})'.format(ARM.ARM_Version))
    root.geometry("1675x900+20+20")   #--("WxH+W+H")  "1675x556+20+20"
    root.minsize(900,575)  #--Set minimum width and height
    root.maxsize(1000,720)  #--Set maximum width and height
    #root.resizable(False,False)
    #root.iconbitmap(cwd + '\\Images\\APH_Icon.ico')
    root.iconbitmap(os.path.join(cwd, 'Images', 'APH_Icon.ico'))
    #---------------------------------------------------
    if not os.path.exists(TxtSettingPath):
        with open(TxtSettingPath, "w") as setting_file:
            setting_file.write("0.0\n")    #--Default Opening_Balance
            setting_file.write("0.0\n")    #--Default Live_Balance
            setting_file.close()
    else:
        with open(TxtSettingPath, 'r') as setting_file:
            lines = setting_file.readlines()
            setting_file.close()
            ARM.Opening_Balance = GUI.R1C3_Value = float(lines[0])
            ARM.Live_Balance = GUI.R1C4_Value = float(lines[1])
            # print(f" Setting Opening_Balance : {ARM.Opening_Balance}")
            # print(f" Setting Live_Balance : {ARM.Live_Balance}")
    #-----------------------------------------------------
    try:                                          #--Check Object creation is required here or not
        kite = KiteApp(enctoken=ARM.enctoken)
    except:
        print("Please Check Internet Connection or Zerodha Login\n")

    #---------------------------------------------------
    try:
        ns.df1 = ARM.instrument_nse
        ns.df2 = ARM.instrument_nfo
        ARM.Exchange_Data_Process = Process(target=Get_Exchange_Data, args=(kite,ns))
        ARM.Exchange_Data_Process.start()
    except:
        print("\n")
        print("Please Check Internet Connection to get Exchange Data\n")
        root.quit()     #--Stop the Tkinter event loop
        root.destroy()  #--Close the GUI cleanly and immediately
        sys.exit()      #--sys.exit(0) Exit the program
    #---------------------------------------------------
    yes_button_image = Image.open(os.path.join(cwd, 'Images', 'Button_Yes.png'))
    no_button_image = Image.open(os.path.join(cwd, 'Images', 'Button_No.png'))
    ok_button_image = Image.open(os.path.join(cwd, 'Images', 'Button_Ok.png'))
    save_button_image = Image.open(os.path.join(cwd, 'Images', 'Button_Save.png'))
    cancel_button_image = Image.open(os.path.join(cwd, 'Images', 'Button_Cancel.png'))
    default_button_image = Image.open(os.path.join(cwd, 'Images', 'Button_Default.png'))
    logout_button_image = Image.open(os.path.join(cwd, 'Images', 'Button_Logout.png'))
    demod_image = Image.open(os.path.join(cwd, 'Images', 'DemoD1.png'))
    demof_image = Image.open(os.path.join(cwd, 'Images', 'DemoF1.png'))
    reald_image = Image.open(os.path.join(cwd, 'Images', 'RealD1.png'))
    realf_image = Image.open(os.path.join(cwd, 'Images', 'RealF1.png'))

    nf_ns_image = Image.open(os.path.join(cwd, 'Images', 'NF_NS.png'))
    nf_s_image = Image.open(os.path.join(cwd, 'Images', 'NF_S.png'))
    nf_tb_image = Image.open(os.path.join(cwd, 'Images', 'NF_TB.png'))
    nf_tr_image = Image.open(os.path.join(cwd, 'Images', 'NF_TR.png'))

    yes_button_image = yes_button_image.resize((80, 40))
    no_button_image = no_button_image.resize((80, 40))
    ok_button_image = ok_button_image.resize((80, 40))
    save_button_image = save_button_image.resize((80, 40))
    cancel_button_image = cancel_button_image.resize((80, 40))
    default_button_image = default_button_image.resize((80, 40))
    logout_button_image = logout_button_image.resize((80, 40))
    demod_image = demod_image.resize((60, 60))
    demof_image = demof_image.resize((60, 60))
    reald_image = reald_image.resize((60, 60))
    realf_image = realf_image.resize((60, 60))
    nf_ns_image = nf_ns_image.resize((80, 30))
    nf_s_image = nf_s_image.resize((80, 30))
    nf_tb_image = nf_tb_image.resize((80, 30))
    nf_tr_image = nf_tr_image.resize((80, 30))

    yes_button_image = ImageTk.PhotoImage(yes_button_image)
    no_button_image = ImageTk.PhotoImage(no_button_image)
    ok_button_image = ImageTk.PhotoImage(ok_button_image)
    save_button_image = ImageTk.PhotoImage(save_button_image)
    cancel_button_image = ImageTk.PhotoImage(cancel_button_image)
    default_button_image = ImageTk.PhotoImage(default_button_image)
    logout_button_image = ImageTk.PhotoImage(logout_button_image)
    demod_image = ImageTk.PhotoImage(demod_image)
    demof_image = ImageTk.PhotoImage(demof_image)
    reald_image = ImageTk.PhotoImage(reald_image)
    realf_image = ImageTk.PhotoImage(realf_image)

    nf_ns_image = ImageTk.PhotoImage(nf_ns_image)
    nf_s_image = ImageTk.PhotoImage(nf_s_image)
    nf_tb_image = ImageTk.PhotoImage(nf_tb_image)
    nf_tr_image = ImageTk.PhotoImage(nf_tr_image)

    #------------------------------------------------------------
    OC_Cell = [[None]*19 for _ in range(40)]
    OC_Cell_Value = [[None]*19 for _ in range(40)]

    for i in range(40):#--Row Counting from 0 to 39
        for j in range(19): #--Column Counting from 0 to 18
            # widget = tk.Label(root, text=f"({i}, {j})")
            if (i in [28,29,30,31,32]) and (j in [8,9]):
                OC_Cell_Value[i][j] = tk.DoubleVar()
                OC_Cell[i][j] = tk.Entry(root,width=8,textvariable=OC_Cell_Value[i][j],justify=CENTER,fg="#F0F0F0",bg="#F0F0F0", borderwidth=0)
                OC_Cell[i][j].grid(row=i,column=j,sticky='nsew')
                # OC_Cell[i][j].bind("<Enter>", lambda event, row=i, column=j:on_SL_Target_enter(event, row, column))
                # OC_Cell[i][j].bind("<Leave>", lambda event, row=i, column=j:on_SL_Target_leave(event, row, column))
                OC_Cell[i][j].bind("<Button-1>", lambda event, row=i, column=j: on_SL_Target_click(event, row, column))
                OC_Cell[i][j].bind('<Return>', lambda event,row=i,column=j:on_SL_Target_Pressed(event, row, column))
            elif(i in [28,29,30,31,32]) and (j in [10]):
                OC_Cell[i][j] = tk.Button(root,width=10,height=1,text="",relief="flat",borderwidth=1)
                OC_Cell[i][j].grid(row=i,column=j,sticky='nesw')
                OC_Cell[i][j].bind("<Enter>", lambda event, row=i, column=j: on_sqoff_enter(event, row, column))
                OC_Cell[i][j].bind("<Leave>", lambda event, row=i, column=j: on_sqoff_leave(event, row, column))
                OC_Cell[i][j].bind("<Button-1>", lambda event, row=i, column=j: on_sqoff_click(event, row, column))
            else:
                if (j == 5):
                    OC_Cell[i][j] = tk.Label(root,width=17,height=1,text="",borderwidth=1)
                    OC_Cell[i][j].grid(row=i,column=j,sticky='nesw')
                else:
                    OC_Cell[i][j] = tk.Label(root,width=10,height=1,text="",borderwidth=1)
                    OC_Cell[i][j].grid(row=i,column=j,sticky='nesw')

                if (26 > i > 4) and (j == 4 or j == 6):
                    OC_Cell[i][j].bind("<Enter>", lambda event, row=i, column=j: on_enter(event, row, column))
                    OC_Cell[i][j].bind("<Leave>", lambda event, row=i, column=j: on_leave(event, row, column))
                    OC_Cell[i][j].bind("<Button-1>", lambda event, row=i, column=j: on_buy_click(event, row, column))
                else:pass
                if (26 > i > 4) and (j == 2 or j == 8):
                    OC_Cell[i][j].bind("<Enter>", lambda event, row=i, column=j: on_enter(event, row, column))
                    OC_Cell[i][j].bind("<Leave>", lambda event, row=i, column=j: on_leave(event, row, column))
                    OC_Cell[i][j].bind("<Button-1>", lambda event, row=i, column=j: on_sell_click(event, row, column))
                else:pass

    #-------------------------------------------------------------------------------------------
    R0C0 = tk.StringVar()
    R0C0options =["Manual Login","Browser Login","Logout"]
    R0C0.set("Login")
    R0C0entry = tk.OptionMenu(root, R0C0,*R0C0options)
    R0C0entry.config(fg="blue",bg="white",width=8,height=2)
    R0C0entry.grid(row=0,column=0,rowspan=2)

    R2C8 = tk.StringVar()
    R2C8entry=tk.Label(root, textvariable=R2C8,justify=CENTER,font='Helvetica 10 bold',fg="black",bg="#F0F0F0")
    R2C8entry.grid(row=2,column=8,columnspan=3,sticky='nsew')

    R0C1 = tk.Button(root,width=8,height=2,text='NRML',command=R0C1Pressed,fg="blue",bg="#00E600") #--'MIS'
    R0C1.grid(row=0,column=1)

    R1C1 = tk.Button(root,width=8,height=2,text='LIMIT',command=R1C1Pressed,fg="blue",bg="#00E600")
    R1C1.grid(row=1,column=1)

    R0C2=tk.Label(root,width=8,height=2,text="NF\nOPT-EXP",bg=GUI.Color1, borderwidth=1)
    R0C2.grid(row=0,column=2,sticky='nsew')

    R1C2 = tk.StringVar()
    R1C2options = []
    R1C2options.append("2023-07-27")
    R1C2.set(R1C2options[0])
    R1C2entry = tk.OptionMenu(root,R1C2,*R1C2options,command=OPT_EXP_R1C2)
    R1C2entry.config(fg="blue",bg="white",width=8,height=1, borderwidth=1)
    R1C2entry.grid(row=1, column=2,sticky='nsew')

    R0C3=tk.Label(root,width=8,text="OP BAL",bg=GUI.Color1, borderwidth=1)
    R0C3.grid(row=0,column=3,sticky='nsew')
    R1C3 = tk.DoubleVar()
    R1C3entry = tk.Entry(root,width=8,textvariable=R1C3,justify=CENTER,fg="blue",bg="white", borderwidth=1)
    R1C3entry.bind('<Return>', lambda event: R1C3Pressed())
    R1C3entry.grid(row=1,column=3,sticky='nsew')

    R0C4=tk.Label(root,width=8,text="LIV BAL",bg=GUI.Color1, borderwidth=1)
    R0C4.grid(row=0,column=4,sticky='nsew')
    R1C4 = tk.DoubleVar()
    R1C4entry = tk.Entry(root,width=8,textvariable=R1C4,state=DISABLED,justify=CENTER)
    R1C4entry.configure(disabledbackground="white",disabledforeground="black",borderwidth=1)
    R1C4entry.grid(row=1,column=4,sticky='nsew')

    R0C5 = tk.Button(root,image=demod_image,command=R0C5Pressed,borderwidth=0)
    R0C5.grid(row=0,column=5,rowspan=2)

    R0C6=tk.Label(root,width=8,text="Loss\nPer Trade",bg=GUI.Color1, borderwidth=1)
    R0C6.grid(row=0,column=6,sticky='nsew')
    R1C6 = tk.IntVar()
    R1C6entry = tk.Entry(root,width=8,textvariable=R1C6,justify=CENTER,fg="blue",bg="white", borderwidth=1)
    R1C6entry.bind('<Return>', lambda event: LPT_Pressed())
    R1C6entry.grid(row=1,column=6,sticky='nsew')

    R0C7=tk.Label(root,width=8,text="Profit\nPer Trade",bg=GUI.Color1, borderwidth=1)
    R0C7.grid(row=0,column=7,sticky='nsew')
    R1C7 = tk.IntVar()
    R1C7entry = tk.Entry(root,width=8,textvariable=R1C7,justify=CENTER,fg="blue",bg="white", borderwidth=1)
    R1C7entry.bind('<Return>', lambda event: PPT_Pressed())
    R1C7entry.grid(row=1,column=7,sticky='nsew')

    R0C8=tk.Label(root,width=8,text="Deadband",bg=GUI.Color1, borderwidth=1)
    R0C8.grid(row=0,column=8,sticky='nsew')
    R1C8 = tk.IntVar()
    R1C8entry = tk.Entry(root,width=8,textvariable=R1C8,justify=CENTER,fg="blue",bg="white", borderwidth=1)
    R1C8entry.bind('<Return>', lambda event: Deadband_Pressed())
    R1C8entry.grid(row=1,column=8,sticky='nsew')
    R1C8.set(ARM.UDS_DEADBAND)

    R0C9=tk.Label(root,width=8,text="Min Diff",bg=GUI.Color1, borderwidth=1)
    R0C9.grid(row=0,column=9,sticky='nsew')
    R1C9 = tk.DoubleVar()
    R1C9entry = tk.Entry(root,width=8,textvariable=R1C9,justify=CENTER,fg="blue",bg="white", borderwidth=1)
    R1C9entry.bind('<Return>', lambda event: Min_Diff_Pressed())
    R1C9entry.grid(row=1,column=9,sticky='nsew')
    R1C9.set(ARM.STD_STR_MIN_diff)

    R0C10=tk.Label(root,width=8,text="Min Bars",bg=GUI.Color1, borderwidth=1)
    R0C10.grid(row=0,column=10,sticky='nsew')
    R1C10 = tk.IntVar()
    R1C10entry = tk.Entry(root,width=8,textvariable=R1C10,justify=CENTER,fg="blue",bg="white", borderwidth=1)
    R1C10entry.bind('<Return>', lambda event: Min_Bars_Pressed())
    R1C10entry.grid(row=1,column=10,sticky='nsew')
    R1C10.set(ARM.CONSECUTIVE_BAR_value)

    R2C4 = tk.Button(root,text='Refresh Trade',command=R2C4Pressed,fg="blue",bg="#FCDFFF")
    R2C4.grid(row=2,column=4,columnspan=3,sticky='nsew')

    R26C0 = tk.Label(root,text='',bg="#FCDFFF")
    R26C0.grid(row=26,column=0,columnspan=11,sticky='nsew')
    #-----------------------------------------------
    OC_Cell[27][0].config(text="Time_In",font='Helvetica 10 bold',bg=GUI.Color1)

    OC_Cell[27][1].config(text="Action",font='Helvetica 10 bold',bg=GUI.Color1)
    # OC_Cell[27][1]=tk.Label(root,text="Symbol",font='Helvetica 10 bold',bg=GUI.Color1)
    # OC_Cell[27][1].grid(row=27,column=1,columnspan=2,sticky='nsew')

    OC_Cell[27][2].config(text="Avg Price",font='Helvetica 10 bold',bg=GUI.Color1)
    OC_Cell[27][3].config(text="Qty",font='Helvetica 10 bold',bg=GUI.Color1)
    OC_Cell[27][4].config(text="SqOff Price",font='Helvetica 10 bold',bg=GUI.Color1)
    OC_Cell[27][5].config(text="Symbol",font='Helvetica 10 bold',bg=GUI.Color1)
    OC_Cell[27][6].config(text="P/L(P)",font='Helvetica 10 bold',bg=GUI.Color1)
    OC_Cell[27][7].config(text="P/L(Rs)",font='Helvetica 10 bold',bg=GUI.Color1)
    OC_Cell[27][8].config(text="SL(P)",font='Helvetica 10 bold',bg=GUI.Color1)
    OC_Cell[27][9].config(text="Target(P)",font='Helvetica 10 bold',bg=GUI.Color1)
    OC_Cell[27][10].config(text="SQ(Off)",font='Helvetica 10 bold',bg=GUI.Color1)

    OC_Cell[28][1]=tk.Label(root,text="")
    OC_Cell[28][1].grid(row=28,column=1,sticky='nsew')
    OC_Cell[29][1]=tk.Label(root,text="")
    OC_Cell[29][1].grid(row=29,column=1,sticky='nsew')
    OC_Cell[30][1]=tk.Label(root,text="")
    OC_Cell[30][1].grid(row=30,column=1,sticky='nsew')
    OC_Cell[31][1]=tk.Label(root,text="")
    OC_Cell[31][1].grid(row=31,column=1,sticky='nsew')
    OC_Cell[32][1]=tk.Label(root,text="")
    OC_Cell[32][1].grid(row=32,column=1,sticky='nsew')
    #-------------------------------------------------
    R3C2 = tk.DoubleVar()
    R3C2entry = tk.Entry(root,width=8,textvariable=R3C2,justify=CENTER,fg="blue",bg="white", borderwidth=1)
    R3C2entry.bind('<Return>', lambda event: R3C2Pressed())
    R3C2entry.grid(row=3,column=2,sticky='nsew')
    R3C3=tk.Label(root,width=8,text="<== Qty ==>",font='Helvetica 10 bold',bg="#4CE84A", borderwidth=1)
    R3C3.grid(row=3,column=3,sticky='nsew')
    R3C4 = tk.DoubleVar()
    R3C4entry = tk.Entry(root,width=8,textvariable=R3C4,justify=CENTER,fg="blue",bg="white", borderwidth=1)
    R3C4entry.bind('<Return>', lambda event: R3C4Pressed())
    R3C4entry.grid(row=3,column=4,sticky='nsew')
    R3C6 = tk.DoubleVar()
    R3C6entry = tk.Entry(root,width=8,textvariable=R3C6,justify=CENTER,fg="blue",bg="white", borderwidth=1)
    R3C6entry.bind('<Return>', lambda event: R3C6Pressed())
    R3C6entry.grid(row=3,column=6,sticky='nsew')
    R3C7=tk.Label(root,width=8,text="<== Qty ==>",font='Helvetica 10 bold',bg="#FF7188", borderwidth=1)
    R3C7.grid(row=3,column=7,sticky='nsew')
    R3C8 = tk.DoubleVar()
    R3C8entry = tk.Entry(root,width=8,textvariable=R3C8,justify=CENTER,fg="blue",bg="white", borderwidth=1)
    R3C8entry.bind('<Return>', lambda event: R3C8Pressed())
    R3C8entry.grid(row=3,column=8,sticky='nsew')

    #-------------------------------------------------------------------------------------------
    CE_Side = tk.Label(root,width=10,height=1,text="CE",font='Helvetica 10 bold',bg="#4CE84A",borderwidth=1)
    CE_Side.grid(row=3,column=0,columnspan=2,sticky='nesw')

    OC_Cell[3][5] = tk.Label(root,width=10,height=1,text=str(ARM.Spot_Value),font='Helvetica 10 bold',bg="yellow",borderwidth=1)
    OC_Cell[3][5].grid(row=3,column=5,rowspan=2,sticky='nesw')

    PE_Side = tk.Label(root,width=10,height=1,text="PE",font='Helvetica 10 bold',bg="#FF7188",borderwidth=1)
    PE_Side.grid(row=3,column=9,columnspan=2,sticky='nesw')

    #-------------------------------------------------------------------------------------------
    OC_Cell[4][0].config(text="CLTP",font='Helvetica 10 bold',bg="#4CE84A")
    OC_Cell[4][1].config(text="LTP",font='Helvetica 10 bold',bg="#4CE84A")
    OC_Cell[4][2].config(text="SELL (Bid)",font='Helvetica 10 bold',bg="#4CE84A")
    OC_Cell[4][3].config(text="ATP",font='Helvetica 10 bold',bg="#4CE84A")
    OC_Cell[4][4].config(text="BUY (Ask)",font='Helvetica 10 bold',bg="#4CE84A")

    #OC_Cell[4][5].config(text="STRIKE",font='Helvetica 10 bold',bg="yellow")

    OC_Cell[4][6].config(text="BUY (Ask)",font='Helvetica 10 bold',bg="#FF7188")
    OC_Cell[4][7].config(text="ATP",font='Helvetica 10 bold',bg="#FF7188")
    OC_Cell[4][8].config(text="SELL (Bid)",font='Helvetica 10 bold',bg="#FF7188")
    OC_Cell[4][9].config(text="LTP",font='Helvetica 10 bold',bg="#FF7188")
    OC_Cell[4][10].config(text="CLTP",font='Helvetica 10 bold',bg="#FF7188")

    OC_Cell[15][5].config(relief="solid",borderwidth=2)
    # for i in range(5,26):      #--Row
    #     for j in range(2):     #--Column
    #         OC_Cell[i][j].config(bg="white")
    #         if (10 < i < 15 ):OC_Cell[i][j].config(bg="#F0F0F0")
    #         if  (i==15) :OC_Cell[i][j].config(bg="yellow")
    #         if (20 > i > 15 ):OC_Cell[i][j].config(bg="#F0F0F0")

    R26C2 = tk.Label(root,text="",fg="black",bg="#FCDFFF")
    R26C2.grid(row=26,column=2,sticky='nesw')

    R26C4 = tk.Label(root,text="",fg="black",bg="#FCDFFF")
    R26C4.grid(row=26,column=4,sticky='nesw')

    R26C6 = tk.Label(root,text="",fg="black",bg="#FCDFFF")
    R26C6.grid(row=26,column=6,sticky='nesw')

    R26C8 = tk.Label(root,text="",fg="black",bg="#FCDFFF")
    R26C8.grid(row=26,column=8,sticky='nesw')
    #-------------------------------------------
    KWS_frame = tk.Frame(root,bg="#FCDFFF")
    KWS_frame.grid(row=26, column=5)
    kws_canvas = tk.Canvas(KWS_frame, width=20, height=20, bg="#FCDFFF")
    kws_canvas.grid(row=0, column=1, padx=10)
    square = kws_canvas.create_rectangle(0, 0, 20, 20, fill='white')
    kws_canvas.tag_bind(square, '<Button-3>', Show_RTT_SDD)
    kws_canvas.tag_bind(square, '<Button-1>', Show_Total_Position)
    #-------------------------------------------
    uds_ce_canvas = tk.Canvas(KWS_frame, width=20, height=20, highlightthickness=0, bg="#FCDFFF")
    uds_ce_canvas.grid(row=0, column=0, padx=10)
    circle = uds_ce_canvas.create_oval(3, 3, 17, 17, fill="yellow", outline="black")
    uds_ce_canvas.tag_bind(circle, "<Button-1>", cepe_circle_click)

    uds_pe_canvas = tk.Canvas(KWS_frame, width=20, height=20, highlightthickness=0, bg="#FCDFFF")
    uds_pe_canvas.grid(row=0, column=2, padx=10)
    circle = uds_pe_canvas.create_oval(3, 3, 17, 17, fill="yellow", outline="black")
    uds_pe_canvas.tag_bind(circle, "<Button-1>", cepe_circle_click)
    #-------------------------------------------
    ce_frame = tk.Frame(root, bg="#FCDFFF")
    ce_frame.grid(row=26, column=0, columnspan=2, sticky="nsew")
    ce_frame.grid_rowconfigure(0, weight=1)
    ce_frame.grid_columnconfigure(0, weight=1)
    ce_canvas = tk.Canvas(ce_frame, width=20, height=20, highlightthickness=0, bg="#FCDFFF")
    ce_canvas.grid(row=0, column=0)
    circle = ce_canvas.create_oval(3, 3, 17, 17, fill="yellow", outline="black")
    ce_canvas.tag_bind(circle, "<Button-1>", cepe_circle_click)

    pe_frame = tk.Frame(root, bg="#FCDFFF")
    pe_frame.grid(row=26, column=9, columnspan=2, sticky="nsew")
    pe_frame.grid_rowconfigure(0, weight=1)
    pe_frame.grid_columnconfigure(0, weight=1)
    pe_canvas = tk.Canvas(pe_frame, width=20, height=20, highlightthickness=0, bg="#FCDFFF")
    pe_canvas.grid(row=0, column=0)
    circle = pe_canvas.create_oval(3, 3, 17, 17, fill="yellow", outline="black")
    pe_canvas.tag_bind(circle, "<Button-1>", cepe_circle_click)
    #-------------------------------------------
    #--All Button Declairation End---------------------------------------------------

    #============ (END) TKinter Code ===========================================

    sys.stdout = Logger() #--Start Logger

    print("-------------------------------------------------------------------------------------")
    print('NIFTY Index Options Trader ({0})                {1}'.format(ARM.ARM_Version,datetime.now().strftime("%H:%M:%S, %d %b %Y, %a")))
    print('One-Click Index Options Trading by A2Soft.')
    print("-------------------------------------------------------------------------------------")
    #===========================================================================
    ARM.ZLatency, ZJitter = measure_latency_and_jitter('www.zerodha.com')
    GLatency, GJtter = measure_latency_and_jitter('www.google.com')

    if ARM.ZLatency is not None and ZJitter is not None:#--Latency(RTT), Standard Deviation of the latency(RTT)
        print(f"Zerodha RTT : {ARM.ZLatency:.2f} ms, SD RTT: {ZJitter:.2f} ms")
    else:
        ARM.ZLatency = 50
        print("Failed to measure Zerodha RTT & SD-RTT")

    if GLatency is not None and GJtter is not None:#--Latency(RTT), Standard Deviation of the latency(RTT)
        print(f"Google  RTT : {GLatency:.2f} ms, SD RTT : {GJtter:.2f} ms")
    else:
        print("Failed to measure Google RTT & SD-RTT")
    #===========================================================================
    ARM.Buy_CE_PE = ""          #--CE/PE
    ARM.Buy_Strike = 0          #--Strike
    ARM.Sell_CE_PE = ""
    ARM.Sell_Strike = 0

    ARM.oc_symbol = "NIFTY"                 #--NIFTY Default Symbol at Start
    ARM.APH_SharedPVar[0] = ARM.oc_symbol   #--NIFTY Default Symbol at Start
    ARM.APH_SharedPVar[1] = ARM.oc_symbol   #--NIFTY Default Symbol at Start

    R3C2.set("{:.0f}".format(GUI.R3C2_Value))
    R3C4.set("{:.0f}".format(GUI.R3C4_Value))
    R3C6.set("{:.0f}".format(GUI.R3C6_Value))
    R3C8.set("{:.0f}".format(GUI.R3C8_Value))

    R1C6.set(GUI.LPT_Value)
    R1C7.set(GUI.PPT_Value)
    GUI.R1C4_Value = GUI.R1C3_Value = 100000.00
    ARM.Opening_Balance = GUI.R1C3_Value
    ARM.Live_Balance = GUI.R1C4_Value
    R1C3.set("{:.2f}".format(GUI.R1C3_Value))
    R1C4.set("{:.2f}".format(GUI.R1C4_Value))

    TT_Login()
    #=============================================
    if (os.path.exists('Etoken.txt')):
        with open('Etoken.txt', 'r') as efile:
            lines = efile.readlines()
        efile.close()
        if (len(lines) == 2):
            ARM.your_user_id = lines[0][0:6]
        else:
            ARM.your_user_id = "User-ID"
    else:
        ARM.your_user_id = "User-ID"
    #=============================================
    if os.path.exists(Initial_Strikes_Path):
        try:
            df = pd.read_csv(Initial_Strikes_Path)
            if (
                    list(df.columns) == ["STRIKE", "CE_LTP", "PE_LTP"] and
                    len(df) == 21
                ):
                Use_Data = POPUP_YESNO("Old Data Found", "You want to use old Straddle Data?\n\nIf No,\n Old Straddle Data file will get Deleted....")
                if Use_Data:
                    ARM.Initial_Strikes_df = df.copy()
                    print("Using old Initial_Strikes.csv")
                    ARM.Use_Initial_Strikes = True
                else:
                    os.remove(Initial_Strikes_Path)
                    ARM.Initial_Strikes_df = pd.DataFrame()
                    print("Old Initial_Strikes.csv deleted")
                    ARM.Use_Initial_Strikes = False
            else:
                os.remove(Initial_Strikes_Path)
                ARM.Initial_Strikes_df = pd.DataFrame()
                ARM.Use_Initial_Strikes = False
                print("Initial_Strikes.csv exists but structure/rows invalid, deleting old file.....")
        except Exception as e:
            print(f"Error reading Initial_Strikes.csv: {e}")
    #=============================================
    #--Show_Total_Position(event=True)
    root.protocol("WM_DELETE_WINDOW", Close_OTM_Direct)
    root.mainloop()