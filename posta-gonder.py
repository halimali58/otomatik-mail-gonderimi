import yfinance as yf
import pandas as pd
import numpy as np
from datetime import datetime
import pytz
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import schedule
import time
import os
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import colors
import base64
import sys

# Türkiye saat dilimi
turkey_tz = pytz.timezone('Europe/Istanbul')

# E-posta ayarları
EMAIL_ADDRESS = "alijak5818@gmail.com"
EMAIL_PASSWORD = "xfbc fuvy fonx kbxi"  # Gmail için uygulama özel şifresi
RECIPIENT_EMAIL = "halimali58@hotmail.com"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

# Sembol listesi
symbols = [
    'A1CAP.IS', 'A1YEN.IS', 'ACSEL.IS', 'ADEL.IS', 'ADESE.IS', 'ADGYO.IS', 'AEFES.IS', 'AFYON.IS', 'AGESA.IS', 'AGHOL.IS',
]

# Zaman dilimleri ve Türkçe karşılıkları
timeframes = {
    '1h': '60d',
    '2h': '3mo',
    '4h': '120d',
    '1d': '1y',
    '1wk': '3y',
    '1mo': '10y'
}

timeframes_tr = {
    '1h': 'Saatlik',
    '2h': '2 Saatlik',
    '4h': '4 Saatlik',
    '1d': 'Günlük',
    '1wk': 'Haftalık',
    '1mo': 'Aylık'
}

# Sekme renkleri
tab_colors = {
    'Saatlik': '87CEEB',
    '2 Saatlik': '98FB98',
    '4 Saatlik': 'FFFFE0',
    'Günlük': 'FF9800',
    'Haftalık': 'E6E6FA',
    'Aylık': 'FFB6C1'
}

# Sinyal doğrulama için çubuk sayıları
min_confirm_bars = 2
max_confirm_bars = 5

# Supertrend indikatörünü hesaplama
def compute_supertrend(df, atr_period=10, factor=3.0, atrline=1.5):
    df = df.copy()
    df['TR'] = pd.concat([
        df['High'] - df['Low'],
        (df['High'] - df['Close'].shift()).abs(),
        (df['Low'] - df['Close'].shift()).abs()
    ], axis=1).max(axis=1)
    df['ATR'] = df['TR'].rolling(window=atr_period).mean()
    df['hl2'] = (df['High'] + df['Low']) / 2
    df['upperband'] = df['hl2'] + factor * df['ATR']
    df['lowerband'] = df['hl2'] - factor * df['ATR']

    n = len(df)
    direction = np.zeros(n, dtype=int)
    supertrend = np.zeros(n)

    close = df['Close'].values
    upperband = df['upperband'].values
    lowerband = df['lowerband'].values

    if n > 0 and not np.isnan(lowerband[0]):
        supertrend[0] = lowerband[0]
        direction[0] = 1

    for i in range(1, n):
        close_i = close[i]
        supertrend_prev = supertrend[i - 1]
        direction_prev = direction[i - 1]
        upperband_i = upperband[i]
        lowerband_i = lowerband[i]

        if close_i > supertrend_prev:
            direction[i] = 1
        elif close_i < supertrend_prev:
            direction[i] = -1
        else:
            direction[i] = direction_prev

        if direction[i] == 1:
            supertrend[i] = max(lowerband_i, supertrend_prev)
        else:
            supertrend[i] = min(upperband_i, supertrend_prev)

    df['supertrend'] = supertrend
    df['direction'] = direction
    df['upatrline'] = df['supertrend'] + atrline * df['ATR']
    df['dnatrline'] = df['supertrend'] - atrline * df['ATR']
    return df

# 2 saatlik veri çekme
def get_2h_data(symbol, period="3mo"):
    try:
        df_1h = yf.download(symbol, period=period, interval="60m", progress=False, auto_adjust=False, timeout=30)
        if df_1h.empty:
            print(f"[UYARI] {symbol} için 1 saatlik veri boş.")
            return None
        if isinstance(df_1h.columns, pd.MultiIndex):
            df_1h.columns = df_1h.columns.get_level_values(0)
        df_1h.index = pd.to_datetime(df_1h.index, utc=True).tz_convert('Europe/Istanbul')
        df_2h = df_1h.resample('2h').agg({
            'Open': 'first',
            'High': 'max',
            'Low': 'min',
            'Close': 'last',
            'Volume': 'sum'
        }).dropna()
        if df_2h.empty:
            print(f"[UYARI] {symbol} için 2 saatlik resample sonrası veri boş.")
            return None
        if len(df_2h) < 10:
            print(f"[UYARI] {symbol} için 2 saatlik veri yetersiz (uzunluk: {len(df_2h)}).")
            return None
        return df_2h
    except Exception as e:
        print(f"[HATA] {symbol} 2 saatlik veri alınırken hata: {e}")
        return None

# Son mumdan önceki AL ve SAT fiyatlarını bulma
def get_previous_signals(df, minConfirmBars=2, maxConfirmBars=5):
    direction = df['direction'].values
    close = df['Close'].values
    supertrend = df['supertrend'].values
    high = df['High'].values
    low = df['Low'].values
    bar_index = np.arange(len(df))

    turnGreen = np.insert((direction[1:] < direction[:-1]), 0, False)
    turnRed = np.insert((direction[1:] > direction[:-1]), 0, False)

    turn_green_indices = bar_index[turnGreen]
    turn_red_indices = bar_index[turnRed]

    last_turn_green_bar = turn_green_indices[-2] if len(turn_green_indices) >= 2 else np.nan
    last_turn_red_bar = turn_red_indices[-2] if len(turn_red_indices) >= 2 else np.nan

    bar_since_turn_green = (len(df) - 2 - last_turn_green_bar) if not np.isnan(last_turn_green_bar) else np.nan
    bar_since_turn_red = (len(df) - 2 - last_turn_red_bar) if not np.isnan(last_turn_red_bar) else np.nan

    barsg = bar_since_turn_green if not np.isnan(bar_since_turn_green) else 1
    barsr = bar_since_turn_red if not np.isnan(bar_since_turn_red) else 1

    hh1 = np.max(high[int(last_turn_red_bar):int(len(df)-1)]) if not np.isnan(last_turn_red_bar) and len(high) >= barsr and barsr > 0 else np.nan
    ll2 = np.min(low[int(last_turn_green_bar):int(len(df)-1)]) if not np.isnan(last_turn_green_bar) and len(low) >= barsg and barsg > 0 else np.nan

    prev_al_price = np.nan
    if not np.isnan(last_turn_green_bar) and minConfirmBars <= bar_since_turn_green <= maxConfirmBars:
        confirmAL = True
        for i in range(minConfirmBars):
            idx = int(len(df) - 2 - i)
            if idx < 0 or close[idx] <= supertrend[idx]:
                confirmAL = False
                break
        if confirmAL and turnGreen[int(last_turn_green_bar)]:
            prev_al_price = ll2 if not np.isnan(ll2) else np.nan

    prev_sat_price = np.nan
    if not np.isnan(last_turn_red_bar) and minConfirmBars <= bar_since_turn_red <= maxConfirmBars:
        confirmSAT = True
        for i in range(minConfirmBars):
            idx = int(len(df) - 2 - i)
            if idx < 0 or close[idx] >= supertrend[idx]:
                confirmSAT = False
                break
        if confirmSAT and turnRed[int(last_turn_red_bar)]:
            prev_sat_price = hh1 if not np.isnan(hh1) else np.nan

    return prev_al_price, prev_sat_price

# Son barın sinyal yönünü belirleme
def get_last_signal(df, minConfirmBars=2, maxConfirmBars=5):
    direction = df['direction'].values
    close = df['Close'].values
    supertrend = df['supertrend'].values
    bar_index = np.arange(len(df))

    turnGreen = np.insert((direction[1:] < direction[:-1]), 0, False)
    turnRed = np.insert((direction[1:] > direction[:-1]), 0, False)

    last_turn_green_bar = bar_index[turnGreen][-1] if np.any(turnGreen) else np.nan
    last_turn_red_bar = bar_index[turnRed][-1] if np.any(turnRed) else np.nan

    bar_since_turn_green = len(df) - 1 - last_turn_green_bar if not np.isnan(last_turn_green_bar) else np.nan
    bar_since_turn_red = len(df) - 1 - last_turn_red_bar if not np.isnan(last_turn_red_bar) else np.nan

    last_signal = None
    if not np.isnan(last_turn_green_bar) and minConfirmBars <= bar_since_turn_green <= maxConfirmBars:
        confirmAL = True
        for i in range(minConfirmBars):
            idx = int(len(df) - 1 - i)
            if idx < 0 or close[idx] <= supertrend[idx]:
                confirmAL = False
                break
        if confirmAL and turnGreen[int(last_turn_green_bar)]:
            last_signal = "AL"

    if not np.isnan(last_turn_red_bar) and minConfirmBars <= bar_since_turn_red <= maxConfirmBars:
        confirmSAT = True
        for i in range(minConfirmBars):
            idx = int(len(df) - 1 - i)
            if idx < 0 or close[idx] >= supertrend[idx]:
                confirmSAT = False
                break
        if confirmSAT and turnRed[int(last_turn_red_bar)]:
            last_signal = "SAT"

    return last_signal

# TradingView'deki getSignal fonksiyonuna uyarlanmış sinyal oluşturma
def get_signals(df, minConfirmBars=2, maxConfirmBars=5, prev_al_price=None, prev_sat_price=None, proximity_threshold=0.02):
    last_signal = get_last_signal(df, minConfirmBars, maxConfirmBars)
    last_buy_signal = None
    last_sell_signal = None
    last_buy_row = None
    last_sell_row = None

    direction = df['direction'].values
    close = df['Close'].values
    supertrend = df['supertrend'].values
    high = df['High'].values
    low = df['Low'].values
    bar_index = np.arange(len(df))

    turnGreen = np.insert((direction[1:] < direction[:-1]), 0, False)
    turnRed = np.insert((direction[1:] > direction[:-1]), 0, False)

    turn_green_indices = bar_index[turnGreen]
    turn_red_indices = bar_index[turnRed]

    last_turn_green_bar = turn_green_indices[-1] if np.any(turnGreen) else np.nan
    last_turn_red_bar = turn_red_indices[-1] if np.any(turnRed) else np.nan

    bar_since_turn_green = len(df) - 1 - last_turn_green_bar if not np.isnan(last_turn_green_bar) else np.nan
    bar_since_turn_red = len(df) - 1 - last_turn_red_bar if not np.isnan(last_turn_red_bar) else np.nan

    barsg = bar_since_turn_green if not np.isnan(bar_since_turn_green) else 1
    barsr = bar_since_turn_red if not np.isnan(bar_since_turn_red) else 1

    ll2 = np.min(low[int(last_turn_green_bar):int(len(df)-1)]) if not np.isnan(last_turn_green_bar) and len(low) >= barsg and barsg > 0 else np.nan
    hh1 = np.max(high[int(last_turn_red_bar):int(len(df)-1)]) if not np.isnan(last_turn_red_bar) and len(high) >= barsr and barsr > 0 else np.nan

    lookback_period = min(75, len(df))
    hh1_75 = np.max(high[-lookback_period:]) if len(high) >= lookback_period else np.nan
    ll2_75 = np.min(low[-lookback_period:]) if len(low) >= lookback_period else np.nan

    validAL = False
    alPrice = np.nan
    if last_signal == "AL":
        if not np.isnan(last_turn_green_bar) and minConfirmBars <= bar_since_turn_green <= maxConfirmBars:
            confirmAL = True
            for i in range(minConfirmBars):
                idx = int(len(df) - 1 - i)
                if idx < 0 or close[idx] <= supertrend[idx]:
                    confirmAL = False
                    break
            validAL = confirmAL and turnGreen[int(last_turn_green_bar)]
            alPrice = ll2 if validAL and not np.isnan(ll2) else np.nan

    validSAT = False
    satPrice = np.nan
    if last_signal == "SAT":
        if not np.isnan(last_turn_red_bar) and minConfirmBars <= bar_since_turn_red <= maxConfirmBars:
            confirmSAT = True
            for i in range(minConfirmBars):
                idx = int(len(df) - 1 - i)
                if idx < 0 or close[idx] >= supertrend[idx]:
                    confirmSAT = False
                    break
            validSAT = confirmSAT and turnRed[int(last_turn_red_bar)]
            satPrice = hh1 if validSAT and not np.isnan(hh1) else np.nan

    currClose = close[-1] if len(close) > 0 else np.nan
    symbol = df['Symbol'].iloc[0]

    def to_scalar(value):
        if isinstance(value, (np.ndarray, pd.Series)):
            return value.item() if value.size == 1 else np.nan
        return value if not np.isnan(value) else np.nan

    alPrice = to_scalar(alPrice)
    satPrice = to_scalar(satPrice)
    currClose = to_scalar(currClose)
    hh1_75 = to_scalar(hh1_75)
    ll2_75 = to_scalar(ll2_75)

    # Yakınlık kontrolü ve alarm için renk bilgisi
    alarm_color = None
    if validAL and not np.isnan(alPrice) and not np.isnan(currClose):
        if currClose <= alPrice * (1 + proximity_threshold):
            price_str = f"{alPrice:.2f} ({ll2_75:.2f})".replace('.', ',') if not np.isnan(ll2_75) else f"{alSupertrend indikatörünü kullanarak hisse senedi fiyat verilerini analiz eden ve alım/satım sinyalleri üreten bu kod, Borsa İstanbul (BIST) hisseleri için çeşitli zaman dilimlerinde (saatlik, 2 saatlik, 4 saatlik, günlük, haftalık, aylık) tarama yapar. Sonuçları bir Excel dosyasına kaydeder ve belirtilen e-posta adresine gönderir. Ayrıca, haftanın belirli günlerinde (Pazartesi-Cuma) saat 19:00'da otomatik olarak çalışacak şekilde zamanlayıcı içerir.
