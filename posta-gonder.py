# Gerekli kütüphaneleri yükleme ve sabit tanımlamalar
!pip install schedule yfinance pandas numpy openpyxl

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
from IPython.display import HTML, display
import sys

# Türkiye saat dilimi
turkey_tz = pytz.timezone('Europe/Istanbul')

# E-posta ayarları
EMAIL_ADDRESS = "alijak5818@gmail.com"
EMAIL_PASSWORD = "vhzl ezjr dgyx fqto"  # Gmail için uygulama özel şifresi
RECIPIENT_EMAIL = "halimali58@hotmail.com"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

# Sembol listesi
symbols = [
    'A1CAP.IS', 'A1YEN.IS', 'ACSEL.IS', 'ADEL.IS', 'ADESE.IS', 'ADGYO.IS', 'AEFES.IS', 'AFYON.IS', 'AGESA.IS', 'AGHOL.IS',
    'AGROT.IS', 'AGYO.IS', 'AHGAZ.IS', 'AHSGY.IS', 'AKBNK.IS', 'AKCNS.IS', 'AKENR.IS', 'AKFGY.IS', 'AKFIS.IS', 'AKFYE.IS',
    'AKGRT.IS', 'AKMGY.IS', 'AKSA.IS', 'AKSEN.IS', 'AKSGY.IS', 'AKSUE.IS', 'AKYHO.IS', 'ALARK.IS', 'ALBRK.IS', 'ALCAR.IS',
    'ALCTL.IS', 'ALFAS.IS', 'ALGYO.IS', 'ALKA.IS', 'ALKIM.IS', 'ALKLC.IS', 'ALTNY.IS', 'ALVES.IS', 'ANELE.IS', 'ANGEN.IS',
    'ANHYT.IS', 'ANSGR.IS', 'ARASE.IS', 'ARCLK.IS', 'ARDYZ.IS', 'ARENA.IS', 'ARMGD.IS', 'ARSAN.IS', 'ARTMS.IS', 'ARZUM.IS',
    'ASELS.IS', 'ASGYO.IS', 'ASTOR.IS', 'ASUZU.IS', 'ATAGY.IS', 'ATAKP.IS', 'ATATP.IS', 'AVGYO.IS', 'AVHOL.IS', 'AVOD.IS',
    'AVPGY.IS', 'AVTUR.IS', 'AYCES.IS', 'AYDEM.IS', 'AYEN.IS', 'AYES.IS', 'AYGAZ.IS', 'AZTEK.IS', 'BAGFS.IS', 'BAHKM.IS',
    'BAKAB.IS', 'BALAT.IS', 'BALSU.IS', 'BANVT.IS', 'BARMA.IS', 'BASCM.IS', 'BASGZ.IS', 'BAYRK.IS', 'BEGYO.IS', 'BERA.IS',
    'BEYAZ.IS', 'BFREN.IS', 'BIENY.IS', 'BIGCH.IS', 'BIGEN.IS', 'BIMAS.IS', 'BINBN.IS', 'BINHO.IS', 'BIOEN.IS', 'BIZIM.IS',
    'BJKAS.IS', 'BLCYT.IS', 'BMSCH.IS', 'BMSTL.IS', 'BNTAS.IS', 'BOBET.IS', 'BORLS.IS', 'BORSK.IS', 'BOSSA.IS', 'BRISA.IS',
    'BRKSN.IS', 'BRKVY.IS', 'BRLSM.IS', 'BRSAN.IS', 'BRYAT.IS', 'BSOKE.IS', 'BTCIM.IS', 'BUCIM.IS', 'BULGS.IS', 'BURCE.IS',
    'BURVA.IS', 'BVSAN.IS', 'BYDNR.IS', 'CANTE.IS', 'CATES.IS', 'CCOLA.IS', 'CELHA.IS', 'CEMAS.IS', 'CEMTS.IS', 'CEMZY.IS',
    'CEOEM.IS', 'CGCAM.IS', 'CIMSA.IS', 'CLEBI.IS', 'CMBTN.IS', 'CMENT.IS', 'CONSE.IS', 'COSMO.IS', 'CRDFA.IS', 'CRFSA.IS',
    'CUSAN.IS', 'CVKMD.IS', 'CWENE.IS', 'DAGHL.IS', 'DAGI.IS', 'DAPGM.IS', 'DARDL.IS', 'DCTTR.IS', 'DENGE.IS', 'DERHL.IS',
    'DERIM.IS', 'DESA.IS', 'DESPC.IS', 'DEVA.IS', 'DGATE.IS', 'DGGYO.IS', 'DGNMO.IS', 'DITAS.IS', 'DMRGD.IS', 'DMSAS.IS',
    'DNISI.IS', 'DOAS.IS', 'DOBUR.IS', 'DOCO.IS', 'DOFER.IS', 'DOGUB.IS', 'DOHOL.IS', 'DOKTA.IS', 'DSTKF.IS', 'DURDO.IS',
    'DURKN.IS', 'DYOBY.IS', 'DZGYO.IS', 'EBEBK.IS', 'ECILC.IS', 'ECZYT.IS', 'EDATA.IS', 'EDIP.IS', 'EFORC.IS', 'EGEEN.IS',
    'EGEGY.IS', 'EGEPO.IS', 'EGGUB.IS', 'EGPRO.IS', 'EGSER.IS', 'EKGYO.IS', 'EKOS.IS', 'EKSUN.IS', 'ELITE.IS', 'EMKEL.IS',
    'ENDAE.IS', 'ENERY.IS', 'ENJSA.IS', 'ENKAI.IS', 'ENSRI.IS', 'ENTRA.IS', 'EPLAS.IS', 'ERBOS.IS', 'ERCB.IS', 'EREGL.IS',
    'ERSU.IS', 'ESCAR.IS', 'ESCOM.IS', 'ESEN.IS', 'ETILR.IS', 'EUPWR.IS', 'EUREN.IS', 'EYGYO.IS', 'FADE.IS', 'FENER.IS',
    'FLAP.IS', 'FMIZP.IS', 'FONET.IS', 'FORMT.IS', 'FORTE.IS', 'FRIGO.IS', 'FROTO.IS', 'FZLGY.IS', 'GARAN.IS', 'GARFA.IS',
    'GEDIK.IS', 'GEDZA.IS', 'GENIL.IS', 'GENTS.IS', 'GEREL.IS', 'GESAN.IS', 'GIPTA.IS', 'GLBMD.IS', 'GLCVY.IS', 'GLRMK.IS',
    'GLRYH.IS', 'GLYHO.IS', 'GMTAS.IS', 'GOKNR.IS', 'GOLTS.IS', 'GOODY.IS', 'GOZDE.IS', 'GRSEL.IS', 'GRTHO.IS', 'GSDDE.IS',
    'GSDHO.IS', 'GSRAY.IS', 'GUBRF.IS', 'GUNDG.IS', 'GWIND.IS', 'GZNMI.IS', 'HALKB.IS', 'HATEK.IS', 'HATSN.IS', 'HDFGS.IS',
    'HEDEF.IS', 'HEKTS.IS', 'HKTM.IS', 'HLGYO.IS', 'HOROZ.IS', 'HRKET.IS', 'HTTBT.IS', 'HUBVC.IS', 'HUNER.IS', 'HURGZ.IS',
    'ICBCT.IS', 'ICUGS.IS', 'IDGYO.IS', 'IEYHO.IS', 'IHAAS.IS', 'IHEVA.IS', 'IHGZT.IS', 'IHLAS.IS', 'IHLGM.IS', 'IHYAY.IS',
    'IMASM.IS', 'INDES.IS', 'INFO.IS', 'INGRM.IS', 'INTEK.IS', 'INTEM.IS', 'INVEO.IS', 'INVES.IS', 'IPEKE.IS', 'ISATR.IS',
    'ISBIR.IS', 'ISBTR.IS', 'ISCTR.IS', 'ISDMR.IS', 'ISFIN.IS', 'ISGSY.IS', 'ISGYO.IS', 'ISKPL.IS', 'ISMEN.IS', 'ISSEN.IS',
    'IZENR.IS', 'IZFAS.IS', 'IZINV.IS', 'IZMDC.IS', 'JANTS.IS', 'KAPLM.IS', 'KAREL.IS', 'KARSN.IS', 'KARTN.IS', 'KATMR.IS',
    'KAYSE.IS', 'KBORU.IS', 'KCAER.IS', 'KCHOL.IS', 'KENT.IS', 'KERVT.IS', 'KFEIN.IS', 'KGYO.IS', 'KIMMR.IS', 'KLGYO.IS',
    'KLKIM.IS', 'KLMSN.IS', 'KLNMA.IS', 'KLRHO.IS', 'KLSER.IS', 'KLSYN.IS', 'KLYPV.IS', 'KMPUR.IS', 'KNFRT.IS', 'KOCMT.IS',
    'KONKA.IS', 'KONTR.IS', 'KONYA.IS', 'KOPOL.IS', 'KORDS.IS', 'KOTON.IS', 'KOZAA.IS', 'KOZAL.IS', 'KRDMA.IS', 'KRDMB.IS',
    'KRDMD.IS', 'KRGYO.IS', 'KRONT.IS', 'KRPLS.IS', 'KRSTL.IS', 'KRTEK.IS', 'KRVGD.IS', 'KSTUR.IS', 'KTLEV.IS', 'KTSKR.IS',
    'KUTPO.IS', 'KUYAS.IS', 'KZBGY.IS', 'KZGYO.IS', 'LIDER.IS', 'LIDFA.IS', 'LILAK.IS', 'LINK.IS', 'LKMNH.IS', 'LMKDC.IS',
    'LOGO.IS', 'LRSHO.IS', 'LUKSK.IS', 'LYDHO.IS', 'LYDYE.IS', 'MAALT.IS', 'MACKO.IS', 'MAGEN.IS', 'MAKIM.IS', 'MAKTK.IS',
    'MANAS.IS', 'MARBL.IS', 'MARKA.IS', 'MARTI.IS', 'MAVI.IS', 'MEDTR.IS', 'MEGMT.IS', 'MEKAG.IS', 'MEPET.IS', 'MERCN.IS',
    'MERIT.IS', 'MERKO.IS', 'METRO.IS', 'METUR.IS', 'MGROS.IS', 'MHRGY.IS', 'MIATK.IS', 'MNDRS.IS', 'MNDTR.IS', 'MOBTL.IS',
    'MOGAN.IS', 'MOPAS.IS', 'MPARK.IS', 'MRGYO.IS', 'MRSHL.IS', 'MSGYO.IS', 'MTRKS.IS', 'MZHLD.IS', 'NATEN.IS', 'NETAS.IS',
    'NIBAS.IS', 'NTGAZ.IS', 'NTHOL.IS', 'NUGYO.IS', 'NUHCM.IS', 'OBAMS.IS', 'OBASE.IS', 'ODAS.IS', 'ODINE.IS', 'OFSYM.IS',
    'ONCSM.IS', 'ONRYT.IS', 'ORCAY.IS', 'ORGE.IS', 'ORMA.IS', 'OSMEN.IS', 'OSTIM.IS', 'OTKAR.IS', 'OTTO.IS', 'OYAKC.IS',
    'OYLUM.IS', 'OYYAT.IS', 'OZATD.IS', 'OZGYO.IS', 'OZKGY.IS', 'OZRDN.IS', 'OZSUB.IS', 'OZYSR.IS', 'PAGYO.IS', 'PAMEL.IS',
    'PAPIL.IS', 'PARSN.IS', 'PASEU.IS', 'PATEK.IS', 'PCILT.IS', 'PEHOL.IS', 'PEKGY.IS', 'PENGD.IS', 'PENTA.IS', 'PETKM.IS',
    'PETUN.IS', 'PGSUS.IS', 'PINSU.IS', 'PKART.IS', 'PKENT.IS', 'PLTUR.IS', 'PNLSN.IS', 'PNSUT.IS', 'POLHO.IS', 'POLTK.IS',
    'PRDGS.IS', 'PRKAB.IS', 'PRKME.IS', 'PRZMA.IS', 'PSDTC.IS', 'PSGYO.IS', 'QNBFK.IS', 'QNBTR.IS', 'QUAGR.IS', 'RALYH.IS',
    'RAYSG.IS', 'REEDR.IS', 'RGYAS.IS', 'RNPOL.IS', 'RODRG.IS', 'RTALB.IS', 'RUBNS.IS', 'RUZYE.IS', 'RYGYO.IS', 'RYSAS.IS',
    'SAFKR.IS', 'SAHOL.IS', 'SAMAT.IS', 'SANEL.IS', 'SANFM.IS', 'SANKO.IS', 'SARKY.IS', 'SASA.IS', 'SAYAS.IS', 'SDTTR.IS',
    'SEGMN.IS', 'SEGYO.IS', 'SEKFK.IS', 'SEKUR.IS', 'SELEC.IS', 'SELGD.IS', 'SELVA.IS', 'SERNT.IS', 'SEYKM.IS', 'SILVR.IS',
    'SISE.IS', 'SKBNK.IS', 'SKTAS.IS', 'SKYLP.IS', 'SKYMD.IS', 'SMART.IS', 'SMRTG.IS', 'SMRVA.IS', 'SNGYO.IS', 'SNICA.IS',
    'SNPAM.IS', 'SODSN.IS', 'SOKE.IS', 'SOKM.IS', 'SONME.IS', 'SRVGY.IS', 'SUMAS.IS', 'SUNTK.IS', 'SURGY.IS', 'SUWEN.IS',
    'TABGD.IS', 'TARKM.IS', 'TATEN.IS', 'TATGD.IS', 'TAVHL.IS', 'TBORG.IS', 'TCELL.IS', 'TCKRC.IS', 'TDGYO.IS', 'TEKTU.IS',
    'TERA.IS', 'TEZOL.IS', 'TGSAS.IS', 'THYAO.IS', 'TKFEN.IS', 'TKNSA.IS', 'TLMAN.IS', 'TMPOL.IS', 'TMSN.IS', 'TNZTP.IS',
    'TOASO.IS', 'TRCAS.IS', 'TRGYO.IS', 'TRILC.IS', 'TSGYO.IS', 'TSKB.IS', 'TSPOR.IS', 'TTKOM.IS', 'TTRAK.IS', 'TUCLK.IS',
    'TUKAS.IS', 'TUPRS.IS', 'TUREX.IS', 'TURGG.IS', 'TURSG.IS', 'UFUK.IS', 'ULAS.IS', 'ULKER.IS', 'ULUFA.IS', 'ULUSE.IS',
    'ULUUN.IS', 'UNLU.IS', 'USAK.IS', 'VAKBN.IS', 'VAKFN.IS', 'VAKKO.IS', 'VANGD.IS', 'VBTYZ.IS', 'VERTU.IS', 'VERUS.IS',
    'VESBE.IS', 'VESTL.IS', 'VKGYO.IS', 'VKING.IS', 'VRGYO.IS', 'VSNMD.IS', 'YAPRK.IS', 'YATAS.IS', 'YAYLA.IS', 'YBTAS.IS',
    'YEOTK.IS', 'YESIL.IS', 'YGGYO.IS', 'YIGIT.IS', 'YKBNK.IS', 'YKSLN.IS', 'YONGA.IS', 'YUNSA.IS', 'YYAPI.IS', 'YYLGD.IS',
    'ZEDUR.IS', 'ZOREN.IS', 'ZRGYO.IS'
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
            price_str = f"{alPrice:.2f} ({ll2_75:.2f})".replace('.', ',') if not np.isnan(ll2_75) else f"{alPrice:.2f}".replace('.', ',')
            last_buy_signal = f"{symbol} - AL => {price_str} - Son: {currClose:.2f}".replace('.', ',')
            last_buy_row = [symbol, "AL", price_str, f"{currClose:.2f}".replace('.', ','), 'green']
            alarm_color = 'green'
        else:
            price_str = f"{alPrice:.2f} ({ll2_75:.2f})".replace('.', ',') if not np.isnan(ll2_75) else f"{alPrice:.2f}".replace('.', ',')
            last_buy_signal = f"{symbol} - AL => {price_str} - Son: {currClose:.2f}".replace('.', ',')
            last_buy_row = [symbol, "AL", price_str, f"{currClose:.2f}".replace('.', ','), None]

    if validSAT and not np.isnan(satPrice) and not np.isnan(currClose):
        if currClose >= satPrice * (1 - proximity_threshold):
            price_str = f"{satPrice:.2f} ({hh1_75:.2f})".replace('.', ',') if not np.isnan(hh1_75) else f"{satPrice:.2f}".replace('.', ',')
            last_sell_signal = f"{symbol} - SAT => {price_str} - Son: {currClose:.2f}".replace('.', ',')
            last_sell_row = [symbol, "SAT", price_str, f"{currClose:.2f}".replace('.', ','), 'red']
            alarm_color = 'red'
        else:
            price_str = f"{satPrice:.2f} ({hh1_75:.2f})".replace('.', ',') if not np.isnan(hh1_75) else f"{satPrice:.2f}".replace('.', ',')
            last_sell_signal = f"{symbol} - SAT => {price_str} - Son: {currClose:.2f}".replace('.', ',')
            last_sell_row = [symbol, "SAT", price_str, f"{currClose:.2f}".replace('.', ','), None]

    return last_buy_signal, last_sell_signal, last_buy_row, last_sell_row, alarm_color

# E-posta gönderme fonksiyonu
def send_email(excel_file_name):
    try:
        print(f"E-posta gönderiliyor: {excel_file_name} -> {RECIPIENT_EMAIL}")  # Log ekle
        msg = MIMEMultipart()
        msg['From'] = EMAIL_ADDRESS
        msg['To'] = RECIPIENT_EMAIL
        msg['Subject'] = f"Dip Tepe Tarama Sonuçları - {datetime.now(turkey_tz).strftime('%d-%m-%Y %H:%M')}"

        body = "Merhaba,\n\nEkli dosyada dip ve tepe tarama sonuçları bulunmaktadır.\n\nİyi günler,\nOtomatik Tarama Sistemi"
        msg.attach(MIMEText(body, 'plain'))

        with open(excel_file_name, 'rb') as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())

        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename= {excel_file_name}')
        msg.attach(part)

        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        print("SMTP bağlantısı kuruldu, login deneniyor...")  # Log ekle
        server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        print("SMTP login başarılı, e-posta gönderiliyor...")  # Log ekle
        text = msg.as_string()
        server.sendmail(EMAIL_ADDRESS, RECIPIENT_EMAIL, text)
        server.quit()
        print(f"✅ {excel_file_name} dosyası {RECIPIENT_EMAIL} adresine başarıyla gönderildi.")
    except Exception as e:
        print(f"⚠️ E-posta gönderilirken hata: {e}")
        raise  # Hatanın loglara yazılmasını sağla

# Excel dosyası indirme bağlantısı
def provide_download_link(excel_file_name):
    try:
        if 'google.colab' in sys.modules:
            with open(excel_file_name, 'rb') as f:
                veri = f.read()
                b64 = base64.b64encode(veri).decode()
                href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{excel_file_name}">📥 Excel dosyasını indirmek için tıklayın</a>'
            display(HTML(href))
        else:
            print(f"✅ Excel dosyası oluşturuldu: {excel_file_name}. Lütfen dosya sisteminizden indirin.")
    except Exception as e:
        print(f"⚠️ Excel dosyası için indirme bağlantısı oluşturulurken hata: {e}")

# Excel dosyası oluşturma ve biçimlendirme
def run_analysis():
    now_file = datetime.now(turkey_tz).strftime("%d-%m-%Y_%H.%M")
    excel_file_name = f"Dip_Tepe_Tarama_Tum_Zamanlar_{now_file}.xlsx"
    now = datetime.now(turkey_tz).strftime("%d-%m-%Y %H:%M")

    if datetime.now(turkey_tz).weekday() >= 5:
        print(f"[UYARI] Bugün ({now}) hafta sonu. Borsalar kapalı olabilir, veri çekimi sınırlı olabilir.")

    any_signals = False
    try:
        with pd.ExcelWriter(excel_file_name, engine='openpyxl') as writer:
            for tf, period in timeframes.items():
                print(f"\n📈 {timeframes_tr[tf]} Zaman Dilimi - SİNYALLER ({now})\n")
                buy_rows = []
                sell_rows = []

                for sym in symbols:
                    print(f"Veri çekiliyor: {sym} ({timeframes_tr[tf]})")
                    try:
                        if tf == '2h':
                            df = get_2h_data(sym, period=period)
                        else:
                            df = yf.download(sym, period=period, interval=tf, progress=False, auto_adjust=False, timeout=30)

                        if df is None or df.empty or len(df) < 60:
                            print(f"[UYARI] {sym} için yeterli veri yok (uzunluk: {len(df) if df is not None else 0}).")
                            continue
                        df['Symbol'] = sym.replace('.IS', '')
                        df.index = pd.to_datetime(df.index, utc=True).tz_convert('Europe/Istanbul')

                        df = compute_supertrend(df, atr_period=10, factor=3.0, atrline=1.5)

                        prev_al_price, prev_sat_price = get_previous_signals(df, minConfirmBars=min_confirm_bars, maxConfirmBars=max_confirm_bars)

                        buy_signal, sell_signal, buy_row, sell_row, alarm_color = get_signals(
                            df, minConfirmBars=min_confirm_bars, maxConfirmBars=max_confirm_bars,
                            prev_al_price=prev_al_price, prev_sat_price=prev_sat_price, proximity_threshold=0.02)

                        if buy_row:
                            buy_rows.append(buy_row)
                            print(f"📈 AL Sinyali: {buy_signal}")
                            any_signals = True
                        if sell_row:
                            sell_rows.append(sell_row)
                            print(f"📉 SAT Sinyali: {sell_signal}")
                            any_signals = True

                    except Exception as e:
                        print(f"[HATA] {sym} {tf} veri işlenirken hata: {e}")

                if buy_rows or sell_rows:
                    columns_buy = ["Sembol", "Sinyal", "Fiyat (Dip)", "Son Fiyat"]
                    columns_sell = ["Sembol", "Sinyal", "Fiyat (Tepe)", "Son Fiyat"]
                    df_buy = pd.DataFrame(buy_rows, columns=columns_buy + ["AlarmColor"]) if buy_rows else pd.DataFrame(columns=columns_buy + ["AlarmColor"])
                    df_sell = pd.DataFrame(sell_rows, columns=columns_sell + ["AlarmColor"]) if sell_rows else pd.DataFrame(columns=columns_sell + ["AlarmColor"])

                    max_rows = max(len(df_buy), len(df_sell)) if (df_buy.empty or df_sell.empty) else max(len(df_buy), len(df_sell))
                    combined_rows = []

                    combined_rows.append([f"📈 AL Sinyali ({now})", "", "", "", "", f"📉 SAT Sinyali ({now})", "", "", "", ""])
                    combined_rows.append(columns_buy + [""] + columns_sell + [""])

                    for i in range(max_rows):
                        buy_row = df_buy.iloc[i].tolist() if i < len(df_buy) else ["", "", "", "", ""]
                        sell_row = df_sell.iloc[i].tolist() if i < len(df_sell) else ["", "", "", "", ""]
                        combined_rows.append(buy_row[:4] + [""] + sell_row[:4] + [""])

                    df_combined = pd.DataFrame(combined_rows, columns=columns_buy + ["Boş"] + columns_sell + ["Boş"])
                    df_combined.to_excel(writer, sheet_name=f"{timeframes_tr[tf]}", index=False, header=False)

                    worksheet = writer.sheets[f"{timeframes_tr[tf]}"]
                    worksheet.sheet_properties.tabColor = tab_colors.get(timeframes_tr[tf], 'FFFFFF')
                    bold_font = Font(bold=True)
                    center_alignment = Alignment(horizontal='center', vertical='center')
                    light_green_fill = PatternFill(start_color='90EE90', end_color='90EE90', fill_type='solid')
                    light_red_fill = PatternFill(start_color='FF9999', end_color='FF9999', fill_type='solid')
                    yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
                    hyperlink_font = Font(color="0000FF", underline="single")

                    worksheet.merge_cells(start_row=1, start_column=11, end_row=1, end_column=14)
                    cell_kl = worksheet.cell(row=1, column=11)
                    cell_kl.value = ""
                    cell_kl.alignment = center_alignment
                    cell_kl.fill = yellow_fill

                    worksheet.cell(row=2, column=11).value = "Sembol"
                    worksheet.cell(row=2, column=12).value = "Sinyal"
                    worksheet.cell(row=2, column=13).value = '=IF(L3="SAT","Fiyat (TEPE)",IF(L3="AL","Fiyat (DIP)",IF(L3="","Fiyat","")))'
                    worksheet.cell(row=2, column=13).alignment = center_alignment
                    worksheet.cell(row=2, column=14).value = "Son Fiyat"

                    for col_idx in range(11, 15):
                        cell = worksheet.cell(row=2, column=col_idx)
                        cell.font = bold_font
                        cell.alignment = center_alignment

                    worksheet.cell(row=3, column=11).value = '=IF(K1="","",HYPERLINK("https://tr.tradingview.com/chart/?symbol=BIST:"&K1,K1))'
                    worksheet.cell(row=3, column=11).font = hyperlink_font
                    worksheet.cell(row=3, column=11).alignment = center_alignment
                    worksheet.cell(row=3, column=12).value = '=IF(K1="","",IFERROR(INDEX(B:B,MATCH(K1,A:A,0)),IFERROR(INDEX(G:G,MATCH(K1,F:F,0)),"")))'
                    worksheet.cell(row=3, column=12).alignment = center_alignment
                    worksheet.cell(row=3, column=13).value = '=IF(K1="","",IFERROR(INDEX(C:C,MATCH(K1,A:A,0)),IFERROR(INDEX(H:H,MATCH(K1,F:F,0)),"")))'
                    worksheet.cell(row=3, column=13).alignment = center_alignment
                    worksheet.cell(row=3, column=14).value = '=IF(K1="","",IFERROR(INDEX(D:D,MATCH(K1,A:A,0)),IFERROR(INDEX(I:I,MATCH(K1,F:F,0)),"")))'
                    worksheet.cell(row=3, column=14).alignment = center_alignment

                    # L3 hücresi için koşullu biçimlendirme
                    green_rule = FormulaRule(formula=['L3="AL"'], fill=light_green_fill)
                    red_rule = FormulaRule(formula=['L3="SAT"'], fill=light_red_fill)
                    worksheet.conditional_formatting.add('L3', green_rule)
                    worksheet.conditional_formatting.add('L3', red_rule)

                    # N3 hücresi için koşullu biçimlendirme
                    green_rule_n3 = FormulaRule(
                        formula=[
                            'AND(L3="AL", K1<>"", N3<>"", VALUE(SUBSTITUTE(N3,",","."))<=VALUE(LEFT(SUBSTITUTE(M3," (",""),FIND(",",SUBSTITUTE(M3," (",""))-1))*1.02)'
                        ],
                        fill=light_green_fill
                    )
                    red_rule_n3 = FormulaRule(
                        formula=[
                            'AND(L3="SAT", K1<>"", N3<>"", VALUE(SUBSTITUTE(N3,",","."))>=VALUE(LEFT(SUBSTITUTE(M3," (",""),FIND(",",SUBSTITUTE(M3," (",""))-1))*0.98)'
                        ],
                        fill=light_red_fill
                    )
                    worksheet.conditional_formatting.add('N3', green_rule_n3)
                    worksheet.conditional_formatting.add('N3', red_rule_n3)

                    row_idx = 1
                    for i, row in enumerate(combined_rows):
                        if row[0].startswith("📈") or row[5].startswith("📉"):
                            worksheet.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=4)
                            worksheet.merge_cells(start_row=row_idx, start_column=6, end_row=row_idx, end_column=9)
                            cell_buy = worksheet.cell(row=row_idx, column=1)
                            cell_sell = worksheet.cell(row=row_idx, column=6)
                            cell_buy.font = bold_font
                            cell_sell.font = bold_font
                            cell_buy.alignment = center_alignment
                            cell_sell.alignment = center_alignment
                            cell_buy.fill = light_green_fill
                            cell_sell.fill = light_red_fill
                        elif row[:4] == columns_buy and row[5:9] == columns_sell:
                            for col_idx in [1, 2, 3, 4, 6, 7, 8, 9]:
                                cell = worksheet.cell(row=row_idx, column=col_idx)
                                cell.font = bold_font
                                cell.alignment = center_alignment
                        else:
                            alarm_color = None
                            if i-2 < len(df_buy) and df_buy.iloc[i-2]['AlarmColor'] == 'green':
                                alarm_color = light_green_fill
                            elif i-2 < len(df_sell) and df_sell.iloc[i-2]['AlarmColor'] == 'red':
                                alarm_color = light_red_fill
                            for col_idx in [1, 2, 3, 4, 6, 7, 8, 9]:
                                cell = worksheet.cell(row=row_idx, column=col_idx)
                                cell.alignment = center_alignment
                                if col_idx == 2 and cell.value == "AL":
                                    cell.fill = light_green_fill
                                elif col_idx == 7 and cell.value == "SAT":
                                    cell.fill = light_red_fill
                                elif col_idx in [4, 9] and alarm_color:
                                    if (col_idx == 4 and alarm_color == light_green_fill) or (col_idx == 9 and alarm_color == light_red_fill):
                                        cell.fill = alarm_color
                                if col_idx in [1, 6] and cell.value and cell.value != "":
                                    tradingview_url = f"https://tr.tradingview.com/chart/?symbol=BIST:{cell.value}"
                                    cell.hyperlink = tradingview_url
                                    cell.font = hyperlink_font
                                    cell.value = cell.value
                        row_idx += 1

                    for col_idx in range(1, 15):
                        column_letter = get_column_letter(col_idx)
                        if col_idx in [3, 8, 13]:
                            worksheet.column_dimensions[column_letter].width = 14.50
                        elif col_idx in [5, 10]:
                            worksheet.column_dimensions[column_letter].width = 2
                        elif col_idx in [1, 2, 4, 6, 7, 9, 11, 12, 14]:
                            worksheet.column_dimensions[column_letter].width = 10
                        else:
                            worksheet.column_dimensions[column_letter].width = 15
                else:
                    print(f"{timeframes_tr[tf]} için sinyal bulunamadı.")

            if not any_signals:
                print("[UYARI] Hiçbir zaman diliminde sinyal bulunamadı. Boş bir sayfa oluşturuluyor.")
                columns = ["Bilgi"]
                df_empty = pd.DataFrame([["Hiçbir sinyal bulunamadı"]], columns=columns)
                df_empty.to_excel(writer, sheet_name="Bilgi", index=False)

        print("✅ Excel dosyası başarıyla oluşturuldu.")
        send_email(excel_file_name)
        provide_download_link(excel_file_name)
    except Exception as e:
        print(f"⚠️ Excel dosyası oluşturulurken hata: {e}")

# Hemen test etmek için
run_analysis()
