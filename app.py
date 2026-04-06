#!/usr/bin/env python3
"""
Event Study Dashboard  v2
=========================
Dashboard interactivo para comparar el comportamiento histórico de activos
financieros alrededor de fechas de eventos clave.

Fuentes de datos: CSV / Excel  |  Bloomberg Terminal (blpapi)

Uso:
    streamlit run app.py
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from datetime import date, timedelta
from dateutil.relativedelta import relativedelta
import io

# ── Bloomberg (opcional) ──────────────────────────────────────────────────────
try:
    import blpapi
    BLOOMBERG_AVAILABLE = True
except ImportError:
    BLOOMBERG_AVAILABLE = False

# ── Kaleido para export PNG (opcional) ────────────────────────────────────────
try:
    import kaleido  # noqa: F401
    KALEIDO_AVAILABLE = True
except ImportError:
    KALEIDO_AVAILABLE = False


# ══════════════════════════════════════════════════════════════════════════════
# CONSTANTES
# ══════════════════════════════════════════════════════════════════════════════

ASSET_TYPES    = ["equity", "rate", "fx", "commodity", "other"]
FIELD_MODES    = ["cumulative_return", "absolute_change", "pct_change", "price"]
BASELINE_MODES = ["base100", "base0", "none"]
FREQUENCIES    = ["daily", "weekly", "monthly", "annual"]

FREQ_LABEL = {"daily": "D", "weekly": "W", "monthly": "M", "annual": "Y"}

FIELD_MODE_LABELS = {
    "price":             "Precio Raw (nivel)",
    "cumulative_return": "Retorno Acumulado (%)",
    "pct_change":        "Cambio Porcentual (%)",
    "absolute_change":   "Cambio Absoluto",
}

BASELINE_LABELS = {
    "base100": "Base 100 (indexado al evento)",
    "base0":   "Base 0 (cambio desde el evento)",
    "none":    "Sin normalizar (nivel real)",
}

ASSET_DEFAULTS = {
    "equity":    {"field_mode": "cumulative_return", "baseline_mode": "base100"},
    "rate":      {"field_mode": "absolute_change",   "baseline_mode": "base0"},
    "fx":        {"field_mode": "cumulative_return", "baseline_mode": "base100"},
    "commodity": {"field_mode": "cumulative_return", "baseline_mode": "base100"},
    "other":     {"field_mode": "price",             "baseline_mode": "none"},
}

EVENT_COLORS = [
    "#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd",
    "#8c564b", "#e377c2", "#7f7f7f", "#bcbd22", "#17becf",
    "#aec7e8", "#ffbb78", "#98df8a", "#ff9896", "#c5b0d5",
    "#ff6b6b", "#4ecdc4", "#45b7d1", "#96ceb4", "#ffeaa7",
    "#636efa", "#ef553b", "#00cc96", "#ab63fa", "#ffa15a",
]

# ══════════════════════════════════════════════════════════════════════════════
# CATALOGO DE EVENTOS DE REFERENCIA (24 eventos)
# ══════════════════════════════════════════════════════════════════════════════

REFERENCE_EVENTS = [
    {"label": "9/11 Attacks",                "date": "2001-09-11"},
    {"label": "Iraq War Start",              "date": "2003-03-20"},
    {"label": "Lehman Brothers Collapse",    "date": "2008-09-15"},
    {"label": "Flash Crash 2010",            "date": "2010-05-06"},
    {"label": "US Credit Downgrade (S&P)",   "date": "2011-08-05"},
    {"label": "Taper Tantrum",               "date": "2013-05-22"},
    {"label": "China Devaluation",           "date": "2015-08-11"},
    {"label": "Oil Crash 2016",              "date": "2016-01-20"},
    {"label": "Brexit Referendum",           "date": "2016-06-24"},
    {"label": "Trump Election 2016",         "date": "2016-11-08"},
    {"label": "Volmageddon (VIX blow-up)",   "date": "2018-02-05"},
    {"label": "Fed Pivot Dec 2018",          "date": "2018-12-24"},
    {"label": "COVID Crash",                 "date": "2020-03-16"},
    {"label": "US Election 2020",            "date": "2020-11-03"},
    {"label": "Meme Stocks (GME)",           "date": "2021-01-27"},
    {"label": "Russia-Ukraine Invasion",     "date": "2022-02-24"},
    {"label": "Fed 75 bps Hike (Jun 2022)",  "date": "2022-06-15"},
    {"label": "UK Gilt Crisis (LDI)",        "date": "2022-09-26"},
    {"label": "SVB Collapse",                "date": "2023-03-10"},
    {"label": "Israel-Hamas Conflict",       "date": "2023-10-07"},
    {"label": "Japan Rate Hike 2024",        "date": "2024-07-31"},
    {"label": "Yen Carry Trade Unwind",      "date": "2024-08-05"},
    {"label": "Trump Election 2024",         "date": "2024-11-05"},
    {"label": "DeepSeek AI Shock",           "date": "2025-01-27"},
    {"label": "Trump Tariffs (Liberation Day)", "date": "2025-04-02"},
]


# ══════════════════════════════════════════════════════════════════════════════
# CATALOGO DE TICKERS BLOOMBERG  (clasificación por tipo de activo)
# ══════════════════════════════════════════════════════════════════════════════
# El catálogo se usa para auto-clasificar columnas del CSV y tickers de
# Bloomberg.  Si un ticker no está aquí, se intenta por patrón de sufijo;
# si tampoco matchea, default = equity.

TICKER_CATALOG: dict[str, dict] = {
    # ── Equity Indices ────────────────────────────────────────────────────────
    # US
    "SPX Index":      {"asset_type": "equity",    "display_name": "S&P 500"},
    "INDU Index":     {"asset_type": "equity",    "display_name": "Dow Jones Industrial"},
    "CCMP Index":     {"asset_type": "equity",    "display_name": "NASDAQ Composite"},
    "NDX Index":      {"asset_type": "equity",    "display_name": "NASDAQ 100"},
    "RTY Index":      {"asset_type": "equity",    "display_name": "Russell 2000"},
    "RAY Index":      {"asset_type": "equity",    "display_name": "Russell 3000"},
    "VIX Index":      {"asset_type": "equity",    "display_name": "VIX (Volatilidad)"},
    "MOVE Index":     {"asset_type": "rate",      "display_name": "MOVE (Vol Rates)"},
    # Europe
    "SX5E Index":     {"asset_type": "equity",    "display_name": "Euro Stoxx 50"},
    "SXXP Index":     {"asset_type": "equity",    "display_name": "Stoxx Europe 600"},
    "UKX Index":      {"asset_type": "equity",    "display_name": "FTSE 100"},
    "DAX Index":      {"asset_type": "equity",    "display_name": "DAX"},
    "CAC Index":      {"asset_type": "equity",    "display_name": "CAC 40"},
    "IBEX Index":     {"asset_type": "equity",    "display_name": "IBEX 35"},
    "FTSEMIB Index":  {"asset_type": "equity",    "display_name": "FTSE MIB"},
    "SMI Index":      {"asset_type": "equity",    "display_name": "Swiss Market"},
    # Asia-Pacific
    "NKY Index":      {"asset_type": "equity",    "display_name": "Nikkei 225"},
    "TPX Index":      {"asset_type": "equity",    "display_name": "TOPIX"},
    "HSI Index":      {"asset_type": "equity",    "display_name": "Hang Seng"},
    "HSCEI Index":    {"asset_type": "equity",    "display_name": "Hang Seng China Ent"},
    "SHCOMP Index":   {"asset_type": "equity",    "display_name": "Shanghai Composite"},
    "SHSZ300 Index":  {"asset_type": "equity",    "display_name": "CSI 300"},
    "KOSPI Index":    {"asset_type": "equity",    "display_name": "KOSPI"},
    "TWSE Index":     {"asset_type": "equity",    "display_name": "TAIEX"},
    "AS51 Index":     {"asset_type": "equity",    "display_name": "ASX 200"},
    "STI Index":      {"asset_type": "equity",    "display_name": "Straits Times"},
    "SENSEX Index":   {"asset_type": "equity",    "display_name": "BSE Sensex"},
    "NIFTY Index":    {"asset_type": "equity",    "display_name": "Nifty 50"},
    # Latin America
    "MEXBOL Index":   {"asset_type": "equity",    "display_name": "IPC Mexico"},
    "IBOV Index":     {"asset_type": "equity",    "display_name": "Bovespa"},
    "IPSA Index":     {"asset_type": "equity",    "display_name": "IPSA Chile"},
    "COLCAP Index":   {"asset_type": "equity",    "display_name": "COLCAP Colombia"},
    # Other EM
    "JALSH Index":    {"asset_type": "equity",    "display_name": "JSE Top 40"},
    "XU100 Index":    {"asset_type": "equity",    "display_name": "BIST 100"},
    "TA-125 Index":   {"asset_type": "equity",    "display_name": "TA-125 Israel"},
    # MSCI
    "MXWO Index":     {"asset_type": "equity",    "display_name": "MSCI World"},
    "MXWD Index":     {"asset_type": "equity",    "display_name": "MSCI ACWI"},
    "MXEF Index":     {"asset_type": "equity",    "display_name": "MSCI EM"},
    "MXEA Index":     {"asset_type": "equity",    "display_name": "MSCI EAFE"},
    "MXLA Index":     {"asset_type": "equity",    "display_name": "MSCI LatAm"},

    # ── Government Bond Yields (Rates) ───────────────────────────────────────
    # US Treasuries
    "USGG3M Index":   {"asset_type": "rate", "display_name": "US 3M Yield"},
    "USGG6M Index":   {"asset_type": "rate", "display_name": "US 6M Yield"},
    "USGG2YR Index":  {"asset_type": "rate", "display_name": "US 2Y Yield"},
    "USGG3YR Index":  {"asset_type": "rate", "display_name": "US 3Y Yield"},
    "USGG5YR Index":  {"asset_type": "rate", "display_name": "US 5Y Yield"},
    "USGG7YR Index":  {"asset_type": "rate", "display_name": "US 7Y Yield"},
    "USGG10YR Index": {"asset_type": "rate", "display_name": "US 10Y Yield"},
    "USGG30YR Index": {"asset_type": "rate", "display_name": "US 30Y Yield"},
    "FDTR Index":     {"asset_type": "rate", "display_name": "Fed Funds Target Rate"},
    "FDTRMID Index":  {"asset_type": "rate", "display_name": "Fed Funds Mid"},
    "US0001M Index":  {"asset_type": "rate", "display_name": "USD SOFR 1M"},
    "US0003M Index":  {"asset_type": "rate", "display_name": "USD SOFR 3M"},
    "SOFRRATE Index": {"asset_type": "rate", "display_name": "SOFR Rate"},
    "USSW2 Curncy":   {"asset_type": "rate", "display_name": "USD 2Y Swap"},
    "USSW5 Curncy":   {"asset_type": "rate", "display_name": "USD 5Y Swap"},
    "USSW10 Curncy":  {"asset_type": "rate", "display_name": "USD 10Y Swap"},
    "USSW30 Curncy":  {"asset_type": "rate", "display_name": "USD 30Y Swap"},
    "USYC2Y10 Index": {"asset_type": "rate", "display_name": "US 2s10s Slope"},
    "USYC3M10 Index": {"asset_type": "rate", "display_name": "US 3M10Y Slope"},
    # Germany
    "GDBR2 Index":    {"asset_type": "rate", "display_name": "Bund 2Y Yield"},
    "GDBR5 Index":    {"asset_type": "rate", "display_name": "Bund 5Y Yield"},
    "GDBR10 Index":   {"asset_type": "rate", "display_name": "Bund 10Y Yield"},
    "GDBR30 Index":   {"asset_type": "rate", "display_name": "Bund 30Y Yield"},
    # UK
    "GUKG2 Index":    {"asset_type": "rate", "display_name": "Gilt 2Y Yield"},
    "GUKG10 Index":   {"asset_type": "rate", "display_name": "Gilt 10Y Yield"},
    "GUKG30 Index":   {"asset_type": "rate", "display_name": "Gilt 30Y Yield"},
    # France
    "GFRN10 Index":   {"asset_type": "rate", "display_name": "OAT 10Y Yield"},
    # Italy
    "GBTPGR10 Index": {"asset_type": "rate", "display_name": "BTP 10Y Yield"},
    # Japan
    "GJGB2 Index":    {"asset_type": "rate", "display_name": "JGB 2Y Yield"},
    "GJGB10 Index":   {"asset_type": "rate", "display_name": "JGB 10Y Yield"},
    # Mexico
    "GMXN02YR Index": {"asset_type": "rate", "display_name": "MBONO 2Y Yield"},
    "GMXN05YR Index": {"asset_type": "rate", "display_name": "MBONO 5Y Yield"},
    "GMXN10YR Index": {"asset_type": "rate", "display_name": "MBONO 10Y Yield"},
    "GMXN30YR Index": {"asset_type": "rate", "display_name": "MBONO 30Y Yield"},
    "MXIBTIIE Index": {"asset_type": "rate", "display_name": "TIIE 28D"},
    "MPSWF Index":    {"asset_type": "rate", "display_name": "TIIE Swap 1Y"},
    # Brazil
    "GEBR10Y Index":  {"asset_type": "rate", "display_name": "Brazil 10Y Yield"},
    # Spreads / Credit
    "CDX IG CDSI GEN 5Y Corp":  {"asset_type": "rate", "display_name": "CDX IG 5Y"},
    "CDX HY CDSI GEN 5Y Corp":  {"asset_type": "rate", "display_name": "CDX HY 5Y"},
    "ITRX EUR CDSI GEN 5Y Corp":{"asset_type": "rate", "display_name": "iTraxx Main 5Y"},
    "ITRX XOVER CDSI GEN 5Y Corp":{"asset_type": "rate", "display_name": "iTraxx Xover 5Y"},
    "LF98OAS Index":  {"asset_type": "rate", "display_name": "US Agg OAS"},
    "LF98TRUU Index": {"asset_type": "rate", "display_name": "US Agg Total Return"},
    # Breakevens / TIPS
    "USGGBE05 Index": {"asset_type": "rate", "display_name": "US 5Y Breakeven"},
    "USGGBE10 Index": {"asset_type": "rate", "display_name": "US 10Y Breakeven"},

    # ── FX ────────────────────────────────────────────────────────────────────
    # Major
    "EURUSD Curncy":  {"asset_type": "fx", "display_name": "EUR/USD"},
    "USDJPY Curncy":  {"asset_type": "fx", "display_name": "USD/JPY"},
    "GBPUSD Curncy":  {"asset_type": "fx", "display_name": "GBP/USD"},
    "USDCHF Curncy":  {"asset_type": "fx", "display_name": "USD/CHF"},
    "AUDUSD Curncy":  {"asset_type": "fx", "display_name": "AUD/USD"},
    "NZDUSD Curncy":  {"asset_type": "fx", "display_name": "NZD/USD"},
    "USDCAD Curncy":  {"asset_type": "fx", "display_name": "USD/CAD"},
    "USDNOK Curncy":  {"asset_type": "fx", "display_name": "USD/NOK"},
    "USDSEK Curncy":  {"asset_type": "fx", "display_name": "USD/SEK"},
    # EM FX
    "USDMXN Curncy":  {"asset_type": "fx", "display_name": "USD/MXN"},
    "USDBRL Curncy":  {"asset_type": "fx", "display_name": "USD/BRL"},
    "USDCLP Curncy":  {"asset_type": "fx", "display_name": "USD/CLP"},
    "USDCOP Curncy":  {"asset_type": "fx", "display_name": "USD/COP"},
    "USDARS Curncy":  {"asset_type": "fx", "display_name": "USD/ARS"},
    "USDCNH Curncy":  {"asset_type": "fx", "display_name": "USD/CNH"},
    "USDCNY Curncy":  {"asset_type": "fx", "display_name": "USD/CNY"},
    "USDKRW Curncy":  {"asset_type": "fx", "display_name": "USD/KRW"},
    "USDINR Curncy":  {"asset_type": "fx", "display_name": "USD/INR"},
    "USDIDR Curncy":  {"asset_type": "fx", "display_name": "USD/IDR"},
    "USDTWD Curncy":  {"asset_type": "fx", "display_name": "USD/TWD"},
    "USDZAR Curncy":  {"asset_type": "fx", "display_name": "USD/ZAR"},
    "USDTRY Curncy":  {"asset_type": "fx", "display_name": "USD/TRY"},
    "USDRUB Curncy":  {"asset_type": "fx", "display_name": "USD/RUB"},
    "USDPLN Curncy":  {"asset_type": "fx", "display_name": "USD/PLN"},
    "USDHUF Curncy":  {"asset_type": "fx", "display_name": "USD/HUF"},
    "USDCZK Curncy":  {"asset_type": "fx", "display_name": "USD/CZK"},
    "USDTHB Curncy":  {"asset_type": "fx", "display_name": "USD/THB"},
    "USDPHP Curncy":  {"asset_type": "fx", "display_name": "USD/PHP"},
    # Crosses
    "EURGBP Curncy":  {"asset_type": "fx", "display_name": "EUR/GBP"},
    "EURJPY Curncy":  {"asset_type": "fx", "display_name": "EUR/JPY"},
    "EURCHF Curncy":  {"asset_type": "fx", "display_name": "EUR/CHF"},
    "GBPJPY Curncy":  {"asset_type": "fx", "display_name": "GBP/JPY"},
    # DXY
    "DXY Index":      {"asset_type": "fx", "display_name": "Dollar Index (DXY)"},
    "BBDXY Index":    {"asset_type": "fx", "display_name": "Bloomberg Dollar"},

    # ── Commodities ───────────────────────────────────────────────────────────
    # Energy
    "CL1 Comdty":     {"asset_type": "commodity", "display_name": "WTI Crude Oil"},
    "CO1 Comdty":     {"asset_type": "commodity", "display_name": "Brent Crude Oil"},
    "NG1 Comdty":     {"asset_type": "commodity", "display_name": "Natural Gas (HH)"},
    "HO1 Comdty":     {"asset_type": "commodity", "display_name": "Heating Oil"},
    "XB1 Comdty":     {"asset_type": "commodity", "display_name": "RBOB Gasoline"},
    # Precious Metals
    "GC1 Comdty":     {"asset_type": "commodity", "display_name": "Gold Futures"},
    "SI1 Comdty":     {"asset_type": "commodity", "display_name": "Silver Futures"},
    "PL1 Comdty":     {"asset_type": "commodity", "display_name": "Platinum Futures"},
    "PA1 Comdty":     {"asset_type": "commodity", "display_name": "Palladium Futures"},
    "XAU Curncy":     {"asset_type": "commodity", "display_name": "Gold Spot (XAU)"},
    "XAG Curncy":     {"asset_type": "commodity", "display_name": "Silver Spot (XAG)"},
    # Base Metals
    "HG1 Comdty":     {"asset_type": "commodity", "display_name": "Copper Futures"},
    "LMAHDS03 Comdty":{"asset_type": "commodity", "display_name": "LME Aluminum 3M"},
    "LMZSDS03 Comdty":{"asset_type": "commodity", "display_name": "LME Zinc 3M"},
    "LMNIDS03 Comdty":{"asset_type": "commodity", "display_name": "LME Nickel 3M"},
    # Agriculture
    "W 1 Comdty":     {"asset_type": "commodity", "display_name": "Wheat"},
    "C 1 Comdty":     {"asset_type": "commodity", "display_name": "Corn"},
    "S 1 Comdty":     {"asset_type": "commodity", "display_name": "Soybeans"},
    "SB1 Comdty":     {"asset_type": "commodity", "display_name": "Sugar #11"},
    "KC1 Comdty":     {"asset_type": "commodity", "display_name": "Coffee Arabica"},
    "CC1 Comdty":     {"asset_type": "commodity", "display_name": "Cocoa"},
    "CT1 Comdty":     {"asset_type": "commodity", "display_name": "Cotton"},
    "LC1 Comdty":     {"asset_type": "commodity", "display_name": "Live Cattle"},
}


# ══════════════════════════════════════════════════════════════════════════════
# CLASIFICACION AUTOMATICA DE TICKERS
# ══════════════════════════════════════════════════════════════════════════════

def classify_ticker(name: str) -> dict:
    """
    Clasifica un ticker (nombre de columna) por tipo de activo.
    1. Busca en el catálogo exacto.
    2. Aplica heurísticas por sufijo / keywords.
    3. Default = equity.
    """
    name_stripped = name.strip()

    # 1) Exact match
    if name_stripped in TICKER_CATALOG:
        info = TICKER_CATALOG[name_stripped]
        at = info["asset_type"]
        defs = ASSET_DEFAULTS.get(at, ASSET_DEFAULTS["equity"])
        return {
            "asset_type": at,
            "display_name": info.get("display_name", name_stripped),
            "field_mode": defs["field_mode"],
            "baseline_mode": defs["baseline_mode"],
        }

    # 2) Heuristics
    upper = name_stripped.upper()

    # Rate keywords
    rate_keywords = [
        "YIELD", "YLD", "RATE", "SWAP", "USGG", "GDBR", "GUKG", "GFRN",
        "GJGB", "GMXN", "GEBR", "TIIE", "SOFR", "OAS", "BREAKEVEN",
        "SLOPE", "CDX", "ITRX", "CDSI", "MOVE",
    ]
    for kw in rate_keywords:
        if kw in upper:
            defs = ASSET_DEFAULTS["rate"]
            return {"asset_type": "rate", "display_name": name_stripped,
                    "field_mode": defs["field_mode"], "baseline_mode": defs["baseline_mode"]}

    # Suffix-based
    if upper.endswith(" CURNCY") or upper.endswith("CURNCY"):
        # Check if it's a metal spot (XAU, XAG) or a swap rate
        if any(m in upper for m in ["XAU", "XAG", "XPT", "XPD"]):
            defs = ASSET_DEFAULTS["commodity"]
            return {"asset_type": "commodity", "display_name": name_stripped,
                    "field_mode": defs["field_mode"], "baseline_mode": defs["baseline_mode"]}
        elif any(sw in upper for sw in ["USSW", "EUSA", "BPSW"]):
            defs = ASSET_DEFAULTS["rate"]
            return {"asset_type": "rate", "display_name": name_stripped,
                    "field_mode": defs["field_mode"], "baseline_mode": defs["baseline_mode"]}
        else:
            defs = ASSET_DEFAULTS["fx"]
            return {"asset_type": "fx", "display_name": name_stripped,
                    "field_mode": defs["field_mode"], "baseline_mode": defs["baseline_mode"]}

    if upper.endswith(" COMDTY") or upper.endswith("COMDTY"):
        defs = ASSET_DEFAULTS["commodity"]
        return {"asset_type": "commodity", "display_name": name_stripped,
                "field_mode": defs["field_mode"], "baseline_mode": defs["baseline_mode"]}

    # FX patterns (e.g., EURUSD, USD/MXN)
    fx_patterns = ["USD", "EUR", "GBP", "JPY", "CHF", "AUD", "NZD", "CAD",
                   "MXN", "BRL", "CNH", "CNY", "KRW", "DXY"]
    if ("/" in name_stripped and len(name_stripped) <= 10) or \
       any(name_stripped.upper().startswith(p) and len(name_stripped) <= 12 for p in fx_patterns):
        count = sum(1 for p in fx_patterns if p in upper)
        if count >= 2:
            defs = ASSET_DEFAULTS["fx"]
            return {"asset_type": "fx", "display_name": name_stripped,
                    "field_mode": defs["field_mode"], "baseline_mode": defs["baseline_mode"]}

    # Commodity keywords
    commodity_keywords = ["OIL", "CRUDE", "GOLD", "SILVER", "COPPER", "WHEAT",
                          "CORN", "SOY", "SUGAR", "COFFEE", "GAS", "BRENT"]
    for kw in commodity_keywords:
        if kw in upper:
            defs = ASSET_DEFAULTS["commodity"]
            return {"asset_type": "commodity", "display_name": name_stripped,
                    "field_mode": defs["field_mode"], "baseline_mode": defs["baseline_mode"]}

    # 3) Default → equity
    defs = ASSET_DEFAULTS["equity"]
    return {"asset_type": "equity", "display_name": name_stripped,
            "field_mode": defs["field_mode"], "baseline_mode": defs["baseline_mode"]}


def auto_detect_tickers(data_df: pd.DataFrame) -> list[dict]:
    """Genera lista de tickers a partir de las columnas del DataFrame."""
    tickers = []
    for col in data_df.columns:
        info = classify_ticker(col)
        tickers.append({
            "ticker": col,
            "display_name": info["display_name"],
            "csv_column": col,
            "bloomberg_field": "PX_LAST",
            "asset_type": info["asset_type"],
            "field_mode": info["field_mode"],
            "baseline_mode": info["baseline_mode"],
        })
    return tickers


# ══════════════════════════════════════════════════════════════════════════════
# CARGA DE DATOS
# ══════════════════════════════════════════════════════════════════════════════

def load_from_file(uploaded_file):
    """Carga datos de precios desde CSV o Excel subido por el usuario."""
    try:
        name = uploaded_file.name.lower()
        if name.endswith((".xlsx", ".xls")):
            df = pd.read_excel(uploaded_file, index_col=0, parse_dates=True)
        else:
            content = uploaded_file.read()
            uploaded_file.seek(0)
            try:
                df = pd.read_csv(io.BytesIO(content), index_col=0, parse_dates=True)
            except Exception:
                df = pd.read_csv(io.BytesIO(content), index_col=0, parse_dates=True, sep=";")
        df.index = pd.to_datetime(df.index, errors="coerce")
        df = df[df.index.notna()].sort_index()
        df = df[~df.index.duplicated(keep="last")]
        df = df.apply(pd.to_numeric, errors="coerce")
        return df, None
    except Exception as e:
        return None, str(e)


def load_from_bloomberg(fields_map: dict, start_date, end_date):
    """Descarga datos históricos desde Bloomberg API."""
    if not BLOOMBERG_AVAILABLE:
        return None, (
            "blpapi no esta instalado.\n"
            "Instalalo con:\n"
            "  pip install --index-url=https://bcms.bloomberg.com/pip/simple/ blpapi\n"
            "Luego reinicia la aplicacion."
        )
    try:
        opts = blpapi.SessionOptions()
        opts.setServerHost("localhost")
        opts.setServerPort(8194)
        session = blpapi.Session(opts)
        if not session.start():
            return None, "No se pudo iniciar sesion Bloomberg. Verifica que el Terminal este abierto."
        if not session.openService("//blp/refdata"):
            session.stop()
            return None, "No se pudo abrir el servicio refdata."
        service = session.getService("//blp/refdata")

        by_field: dict[str, list[str]] = {}
        for ticker, field in fields_map.items():
            by_field.setdefault(field, []).append(ticker)

        all_series: dict[str, pd.Series] = {}

        for field, field_tickers in by_field.items():
            req = service.createRequest("HistoricalDataRequest")
            for t in field_tickers:
                req.getElement("securities").appendValue(t)
            req.getElement("fields").appendValue(field)
            req.set("startDt", start_date.strftime("%Y%m%d"))
            req.set("endDt",   end_date.strftime("%Y%m%d"))
            req.set("periodicitySelection", "DAILY")
            req.set("nonTradingDayFillOption", "ACTIVE_DAYS_ONLY")
            session.sendRequest(req)

            while True:
                ev = session.nextEvent(3000)
                for msg in ev:
                    if msg.messageType() == blpapi.Name("HistoricalDataResponse"):
                        sec_data = msg.getElement("securityData")
                        sec_name = sec_data.getElementAsString("security")
                        if sec_data.hasElement("securityError"):
                            err_txt = sec_data.getElement("securityError").getElementAsString("message")
                            st.warning(f"Bloomberg: {sec_name} -> {err_txt}")
                            continue
                        fd = sec_data.getElement("fieldData")
                        dates_l, vals_l = [], []
                        for j in range(fd.numValues()):
                            pt = fd.getValue(j)
                            try:
                                d = pt.getElementAsDatetime("date")
                                v = pt.getElementAsFloat(field)
                                dates_l.append(pd.Timestamp(d.year, d.month, d.day))
                                vals_l.append(v)
                            except Exception:
                                pass
                        if dates_l:
                            all_series[sec_name] = pd.Series(vals_l, index=dates_l, name=sec_name)
                if ev.eventType() == blpapi.Event.RESPONSE:
                    break
        session.stop()
        if not all_series:
            return None, "Bloomberg no devolvio datos. Revisa tickers y rango de fechas."
        df = pd.DataFrame(all_series).sort_index()
        return df, None
    except Exception as e:
        return None, f"Error Bloomberg ({type(e).__name__}): {e}"


# ══════════════════════════════════════════════════════════════════════════════
# MOTOR DE VENTANAS
# ══════════════════════════════════════════════════════════════════════════════

def build_target_dates(event_date, lookback, lookforward, frequency):
    event_ts = pd.Timestamp(event_date)
    result = []
    for offset in range(-lookback, lookforward + 1):
        if frequency == "daily":
            target = event_ts + pd.Timedelta(days=offset)
        elif frequency == "weekly":
            target = event_ts + pd.Timedelta(weeks=offset)
        elif frequency == "monthly":
            target = event_ts + relativedelta(months=offset)
        elif frequency == "annual":
            target = event_ts + relativedelta(years=offset)
        else:
            target = event_ts + pd.Timedelta(days=offset)
        result.append((offset, target))
    return result


def get_prior_observation(target_ts, series_index):
    prior = series_index[series_index <= target_ts]
    return prior[-1] if len(prior) > 0 else None


def extract_aligned_values(series, event_date, lookback, lookforward, frequency):
    target_dates = build_target_dates(event_date, lookback, lookforward, frequency)
    result = {}
    for period, target_ts in target_dates:
        mapped = get_prior_observation(target_ts, series.index)
        result[period] = series.loc[mapped] if mapped is not None else np.nan
    return result


# ══════════════════════════════════════════════════════════════════════════════
# TRANSFORMACIONES
# ══════════════════════════════════════════════════════════════════════════════

def apply_transformation(raw_dict: dict, field_mode: str, baseline_mode: str) -> dict:
    """
    Transforma valores alineados.
    - price + none  => nivel real (sin tocar)
    - price + base100 => indexado 100
    - price + base0 => cambio absoluto
    - cumulative_return / pct_change => (v/anchor - 1)*100
    - absolute_change => v - anchor
    """
    anchor = raw_dict.get(0, np.nan)
    if pd.isna(anchor):
        return {p: np.nan for p in raw_dict}

    out = {}
    for p, v in raw_dict.items():
        if pd.isna(v):
            out[p] = np.nan
            continue

        if field_mode in ("cumulative_return", "pct_change"):
            if anchor == 0:
                out[p] = np.nan
            else:
                out[p] = (v / anchor - 1) * 100

        elif field_mode == "absolute_change":
            out[p] = v - anchor

        elif field_mode == "price":
            # *** CORREGIDO: "none" devuelve el nivel real ***
            if baseline_mode == "base100":
                out[p] = (v / anchor * 100) if anchor != 0 else np.nan
            elif baseline_mode == "base0":
                out[p] = v - anchor
            else:  # none → valor raw tal cual
                out[p] = v
        else:
            out[p] = v
    return out


# ══════════════════════════════════════════════════════════════════════════════
# PLOTLY
# ══════════════════════════════════════════════════════════════════════════════

def build_period_labels(periods, freq):
    fl = FREQ_LABEL.get(freq, "T")
    return [
        "Evento (T0)" if p == 0 else (f"{fl}+{p}" if p > 0 else f"{fl}{p}")
        for p in periods
    ]


def create_event_chart(
    ticker_id, display_name, events, aligned_data,
    field_mode, baseline_mode, frequency, show_avg=True
):
    """Crea figura Plotly para un ticker con una linea por evento + promedio."""
    fig = go.Figure()

    y_title = FIELD_MODE_LABELS.get(field_mode, field_mode)
    all_periods = sorted({p for ev_data in aligned_data.values() for p in ev_data})
    period_labels = build_period_labels(all_periods, frequency)
    label_map = dict(zip(all_periods, period_labels))

    # ── Lineas por evento ─────────────────────────────────────────────────────
    for i, ev in enumerate(events):
        ev_label = ev.get("label", f"Evento {i+1}")
        ev_date  = ev.get("date", "")
        color    = EVENT_COLORS[i % len(EVENT_COLORS)]
        if ev_label not in aligned_data:
            continue
        ev_data = aligned_data[ev_label]
        y_vals  = [ev_data.get(p, np.nan) for p in all_periods]

        hover = []
        for j, p in enumerate(all_periods):
            v = ev_data.get(p, np.nan)
            vs = f"{v:,.4f}" if not pd.isna(v) else "Sin dato"
            hover.append(
                f"<b>{ev_label}</b><br>Fecha evento: {ev_date}<br>"
                f"Periodo: {label_map[p]}<br>{y_title}: {vs}"
            )
        fig.add_trace(go.Scatter(
            x=all_periods, y=y_vals, mode="lines+markers",
            name=f"{ev_label} ({ev_date})",
            line=dict(color=color, width=2.5),
            marker=dict(size=5, color=color),
            hovertext=hover, hoverinfo="text", connectgaps=False,
        ))

    # ── Linea promedio ────────────────────────────────────────────────────────
    if show_avg and len(aligned_data) > 1:
        mean_y = []
        for p in all_periods:
            vals = [d.get(p, np.nan) for d in aligned_data.values()]
            clean = [v for v in vals if not pd.isna(v)]
            mean_y.append(np.mean(clean) if clean else np.nan)

        fig.add_trace(go.Scatter(
            x=all_periods, y=mean_y, mode="lines",
            name="Promedio",
            line=dict(color="black", width=3.5, dash="dot"),
            opacity=0.65,
            hovertemplate="<b>Promedio</b><br>Periodo: %{x}<br>Valor: %{y:.4f}<extra></extra>",
        ))

    # ── Linea vertical T=0 ───────────────────────────────────────────────────
    fig.add_vline(
        x=0, line_dash="dash", line_color="rgba(80,80,80,0.55)", line_width=1.8,
        annotation_text="  T=0", annotation_position="top",
        annotation_font=dict(size=11, color="#666"),
    )

    # ── Linea horizontal de referencia ────────────────────────────────────────
    if field_mode in ("cumulative_return", "pct_change", "absolute_change"):
        fig.add_hline(y=0, line_color="rgba(150,150,150,0.35)", line_width=1)
    elif field_mode == "price" and baseline_mode == "base100":
        fig.add_hline(y=100, line_color="rgba(150,150,150,0.35)", line_width=1)

    # ── Ticks ─────────────────────────────────────────────────────────────────
    step = max(1, len(all_periods) // 14)
    tick_vals = [p for i, p in enumerate(all_periods) if i % step == 0 or p == 0]
    tick_text = [label_map[p] for p in tick_vals]

    fig.update_layout(
        title=dict(text=f"<b>{display_name or ticker_id}</b>",
                   font=dict(size=15, color="#0d1b2a"), x=0),
        xaxis=dict(title="Periodo relativo al evento",
                   tickvals=tick_vals, ticktext=tick_text,
                   gridcolor="#ebebeb", zeroline=False, tickfont=dict(size=11)),
        yaxis=dict(title=y_title, gridcolor="#ebebeb", tickfont=dict(size=11)),
        plot_bgcolor="white", paper_bgcolor="white", hovermode="closest",
        legend=dict(orientation="h", yanchor="bottom", y=-0.42,
                    xanchor="center", x=0.5, font=dict(size=11),
                    bordercolor="#ddd", borderwidth=1),
        height=490, margin=dict(l=70, r=30, t=55, b=140),
    )
    return fig


# ══════════════════════════════════════════════════════════════════════════════
# CSV DE MUESTRA
# ══════════════════════════════════════════════════════════════════════════════

def generate_sample_csv() -> str:
    np.random.seed(42)
    dates = pd.bdate_range("2015-01-02", "2024-12-31")
    n, ds = len(dates), [d.strftime("%Y-%m-%d") for d in dates]

    def idx(d):
        return ds.index(d) if d in ds else None

    r = np.random.normal(0.0004, 0.010, n)
    for sd, sv in [("2020-03-16",-0.12),("2020-03-17",-0.06),
                   ("2022-02-24",-0.030),("2023-03-10",-0.025)]:
        i = idx(sd)
        if i is not None: r[i] = sv
    spx = 2000.0 * np.exp(np.cumsum(r))

    ch = np.random.normal(0.00005, 0.035, n)
    for sd, sv in [("2020-03-16",-0.25),("2022-02-24",0.08),("2022-06-15",0.15)]:
        i = idx(sd)
        if i is not None: ch[i] = sv
    us10y = np.maximum(0.05, 2.0 + np.cumsum(ch))

    fx = np.random.normal(0.00008, 0.007, n)
    for sd, sv in [("2020-03-16",0.085),("2022-02-24",0.012)]:
        i = idx(sd)
        if i is not None: fx[i] = sv
    usdmxn = 14.5 * np.exp(np.cumsum(fx))

    gr = np.random.normal(0.0003, 0.009, n)
    for sd, sv in [("2020-03-16",-0.04),("2022-02-24",0.03)]:
        i = idx(sd)
        if i is not None: gr[i] = sv
    gold = 1180.0 * np.exp(np.cumsum(gr))

    df = pd.DataFrame({
        "SPX Index":      np.round(spx, 2),
        "USGG10YR Index": np.round(us10y, 3),
        "USDMXN Curncy":  np.round(usdmxn, 4),
        "XAU Curncy":     np.round(gold, 2),
    }, index=dates)
    df.index.name = "Date"
    return df.to_csv()


# ══════════════════════════════════════════════════════════════════════════════
# APP PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════════

def main():
    st.set_page_config(
        page_title="Event Study Dashboard", page_icon="📊",
        layout="wide", initial_sidebar_state="expanded",
    )
    st.markdown("""
    <style>
    .main-hdr{font-size:1.9rem;font-weight:700;color:#0d1b2a;line-height:1.2}
    .sub-hdr{color:#555;font-size:.92rem;margin-top:4px;margin-bottom:1.4rem}
    .sec-title{font-size:.78rem;font-weight:600;text-transform:uppercase;
               letter-spacing:.06em;color:#888;margin-bottom:6px;margin-top:2px}
    .status-ok{background:#e8f5e9;border-left:3px solid #43a047;padding:7px 11px;
               border-radius:4px;color:#1b5e20;font-size:.84rem;margin-bottom:6px}
    .status-warn{background:#fffde7;border-left:3px solid #fbc02d;padding:7px 11px;
                 border-radius:4px;color:#6d4c00;font-size:.84rem;margin-bottom:6px}
    div[data-testid="stExpander"]{border:1px solid #e8e8e8;border-radius:6px}
    </style>
    """, unsafe_allow_html=True)

    st.markdown('<div class="main-hdr">📊 Event Study Dashboard</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="sub-hdr">Compara el comportamiento historico de activos financieros '
        'alrededor de fechas de eventos clave — Bloomberg y CSV/Excel.</div>',
        unsafe_allow_html=True,
    )

    # ── Init session state ────────────────────────────────────────────────────
    if "events" not in st.session_state:
        st.session_state.events = [
            {"label": "COVID Crash",         "date": "2020-03-16"},
            {"label": "Russia-Ukraine",      "date": "2022-02-24"},
            {"label": "SVB Crisis",          "date": "2023-03-10"},
        ]
    if "tickers" not in st.session_state:
        st.session_state.tickers = []  # se auto-pobla al subir archivo
    if "last_file_key" not in st.session_state:
        st.session_state.last_file_key = None

    # ══════════════════════════════════════════════════════════════════════════
    # SIDEBAR
    # ══════════════════════════════════════════════════════════════════════════
    with st.sidebar:
        st.markdown("### ⚙️ Configuracion")

        # ── Fuente ────────────────────────────────────────────────────────────
        st.markdown('<div class="sec-title">Fuente de datos</div>', unsafe_allow_html=True)
        data_source = st.radio("", ["📂 CSV / Excel", "🔵 Bloomberg"], label_visibility="collapsed")
        use_bloomberg = "Bloomberg" in data_source
        data_df = None

        if not use_bloomberg:
            st.caption("Sube tu archivo de precios historicos.")
            uploaded = st.file_uploader("", type=["csv","xlsx","xls"], label_visibility="collapsed")
            if uploaded:
                data_df, err = load_from_file(uploaded)
                if err:
                    st.error(f"Error: {err}")
                else:
                    st.markdown(
                        f'<div class="status-ok">✅ {len(data_df.columns)} series · '
                        f'{len(data_df):,} fechas<br>'
                        f'{data_df.index[0].date()} → {data_df.index[-1].date()}</div>',
                        unsafe_allow_html=True,
                    )
                    # *** AUTO-DETECTAR tickers al subir archivo nuevo ***
                    file_key = uploaded.name + str(uploaded.size)
                    if st.session_state.last_file_key != file_key:
                        st.session_state.last_file_key = file_key
                        st.session_state.tickers = auto_detect_tickers(data_df)
                        st.rerun()

                    with st.expander("Vista previa"):
                        st.dataframe(data_df.tail(5), use_container_width=True)
        else:
            if BLOOMBERG_AVAILABLE:
                st.markdown('<div class="status-ok">✅ blpapi detectado</div>', unsafe_allow_html=True)
            else:
                st.markdown('<div class="status-warn">⚠️ blpapi no encontrado</div>', unsafe_allow_html=True)
                with st.expander("Instrucciones"):
                    st.code("pip install --index-url=https://bcms.bloomberg.com/pip/simple/ blpapi", language="bash")
            bbg_start = st.date_input("Desde", value=date(2010,1,1), key="bbg_s")
            bbg_end   = st.date_input("Hasta", value=date.today(),   key="bbg_e")

        st.divider()

        # ── Eventos ───────────────────────────────────────────────────────────
        st.markdown('<div class="sec-title">📅 Eventos</div>', unsafe_allow_html=True)

        remove_ev = None
        for i, ev in enumerate(st.session_state.events):
            with st.expander(f"🔴 {ev['label']}", expanded=(i == 0)):
                c1, c2 = st.columns([3, 2])
                with c1:
                    new_label = st.text_input("Nombre", value=ev["label"], key=f"el_{i}")
                with c2:
                    new_date = st.text_input("Fecha", value=ev["date"], key=f"ed_{i}", help="YYYY-MM-DD")
                st.session_state.events[i] = {"label": new_label, "date": new_date}
                if len(st.session_state.events) > 1:
                    if st.button("🗑 Eliminar", key=f"re_{i}", use_container_width=True):
                        remove_ev = i
        if remove_ev is not None:
            st.session_state.events.pop(remove_ev)
            st.rerun()

        col_add, col_ref = st.columns(2)
        with col_add:
            if st.button("➕ Agregar", use_container_width=True):
                n = len(st.session_state.events) + 1
                st.session_state.events.append({"label": f"Evento {n}", "date": "2024-01-01"})
                st.rerun()
        with col_ref:
            if st.button("📋 Catalogo", use_container_width=True, help="Ver eventos de referencia"):
                st.session_state.show_event_catalog = True

        st.divider()

        # ── Ventana ───────────────────────────────────────────────────────────
        st.markdown('<div class="sec-title">🪟 Ventana de analisis</div>', unsafe_allow_html=True)
        frequency = st.selectbox(
            "Frecuencia", FREQUENCIES,
            format_func=lambda x: {"daily":"Diaria","weekly":"Semanal",
                                   "monthly":"Mensual","annual":"Anual"}[x],
        )
        c1, c2 = st.columns(2)
        with c1:
            lookback    = st.number_input("Lookback", min_value=1, max_value=500, value=30)
        with c2:
            lookforward = st.number_input("Lookforward", min_value=1, max_value=500, value=30)

        st.divider()

        # ── Tickers ───────────────────────────────────────────────────────────
        st.markdown('<div class="sec-title">📈 Tickers</div>', unsafe_allow_html=True)

        if st.session_state.tickers:
            st.caption(f"{len(st.session_state.tickers)} tickers configurados")

        remove_tk = None
        for i, tk in enumerate(st.session_state.tickers):
            lbl = tk.get("display_name") or tk["ticker"]
            at_emoji = {"equity":"📈","rate":"📊","fx":"💱","commodity":"🛢️"}.get(tk.get("asset_type",""),"📈")
            with st.expander(f"{at_emoji} {lbl}", expanded=False):
                tk["ticker"]       = st.text_input("Ticker ID", value=tk["ticker"], key=f"tt_{i}")
                tk["display_name"] = st.text_input("Nombre",    value=tk.get("display_name",""), key=f"tn_{i}")

                if not use_bloomberg:
                    if data_df is not None and len(data_df.columns) > 0:
                        cols = list(data_df.columns)
                        def_col = tk.get("csv_column", cols[0])
                        if def_col not in cols:
                            def_col = cols[0]
                        tk["csv_column"] = st.selectbox("Columna", cols,
                            index=cols.index(def_col), key=f"tc_{i}")
                    else:
                        tk["csv_column"] = st.text_input("Columna", value=tk.get("csv_column",""), key=f"tc_{i}")
                else:
                    tk["bloomberg_field"] = st.text_input("Campo BBG",
                        value=tk.get("bloomberg_field","PX_LAST"), key=f"tbf_{i}",
                        help="PX_LAST | YLD_YTM_MID | PX_BID")

                tk["asset_type"] = st.selectbox("Tipo activo", ASSET_TYPES,
                    index=ASSET_TYPES.index(tk.get("asset_type","equity")),
                    format_func=str.capitalize, key=f"ta_{i}")

                defs = ASSET_DEFAULTS.get(tk["asset_type"], ASSET_DEFAULTS["equity"])
                tk["field_mode"] = st.selectbox("Transformacion", FIELD_MODES,
                    index=FIELD_MODES.index(tk.get("field_mode", defs["field_mode"])),
                    format_func=lambda x: FIELD_MODE_LABELS.get(x,x), key=f"tfm_{i}")
                tk["baseline_mode"] = st.selectbox("Baseline", BASELINE_MODES,
                    index=BASELINE_MODES.index(tk.get("baseline_mode", defs["baseline_mode"])),
                    format_func=lambda x: BASELINE_LABELS.get(x,x), key=f"tbm_{i}")

                st.session_state.tickers[i] = tk
                if st.button("🗑 Eliminar", key=f"rt_{i}", use_container_width=True):
                    remove_tk = i

        if remove_tk is not None:
            st.session_state.tickers.pop(remove_tk)
            st.rerun()

        if st.button("➕ Agregar ticker", use_container_width=True):
            n = len(st.session_state.tickers) + 1
            st.session_state.tickers.append({
                "ticker": f"TICKER_{n}", "display_name": f"Activo {n}",
                "csv_column": "", "bloomberg_field": "PX_LAST",
                "asset_type": "equity", "field_mode": "cumulative_return", "baseline_mode": "base100",
            })
            st.rerun()

        st.divider()

        # ── Opciones de grafica ───────────────────────────────────────────────
        st.markdown('<div class="sec-title">Opciones de grafica</div>', unsafe_allow_html=True)
        show_avg = st.checkbox("Mostrar linea promedio", value=True)

        st.divider()
        run = st.button("🚀 Ejecutar analisis", type="primary", use_container_width=True)

    # ══════════════════════════════════════════════════════════════════════════
    # CONTENIDO PRINCIPAL
    # ══════════════════════════════════════════════════════════════════════════

    # ── Catalogo de eventos (popup) ───────────────────────────────────────────
    if st.session_state.get("show_event_catalog", False):
        st.session_state.show_event_catalog = False
        st.markdown("### 📋 Catalogo de Eventos de Referencia")
        st.caption("Selecciona los eventos que quieras agregar a tu analisis.")
        ev_df = pd.DataFrame(REFERENCE_EVENTS)
        ev_df.index = range(1, len(ev_df) + 1)
        ev_df.columns = ["Evento", "Fecha"]
        st.dataframe(ev_df, use_container_width=True)

        selected = st.multiselect(
            "Selecciona eventos para agregar:",
            options=[f"{e['label']} ({e['date']})" for e in REFERENCE_EVENTS],
        )
        if selected and st.button("✅ Agregar seleccionados"):
            for sel in selected:
                for ref_ev in REFERENCE_EVENTS:
                    tag = f"{ref_ev['label']} ({ref_ev['date']})"
                    if tag == sel:
                        # No duplicar
                        existing = [e["date"] for e in st.session_state.events]
                        if ref_ev["date"] not in existing:
                            st.session_state.events.append(
                                {"label": ref_ev["label"], "date": ref_ev["date"]}
                            )
            st.rerun()
        st.divider()

    # ── Catalogo de tickers Bloomberg ─────────────────────────────────────────
    with st.expander("📚 Catalogo de Tickers Bloomberg (referencia)"):
        cat_rows = []
        for tk_id, info in TICKER_CATALOG.items():
            cat_rows.append({
                "Ticker Bloomberg": tk_id,
                "Nombre": info.get("display_name", tk_id),
                "Tipo": info["asset_type"].capitalize(),
            })
        cat_df = pd.DataFrame(cat_rows)

        # Filtros
        filter_type = st.selectbox("Filtrar por tipo:", ["Todos"] + [a.capitalize() for a in ASSET_TYPES], key="cat_filter")
        if filter_type != "Todos":
            cat_df = cat_df[cat_df["Tipo"] == filter_type]

        search_q = st.text_input("Buscar:", key="cat_search", placeholder="Ej: Mexico, Gold, SPX...")
        if search_q:
            mask = cat_df.apply(lambda row: search_q.upper() in " ".join(row.values).upper(), axis=1)
            cat_df = cat_df[mask]

        st.caption(f"{len(cat_df)} tickers en catalogo")
        st.dataframe(cat_df, use_container_width=True, height=300)

    # ── Pantalla inicial ──────────────────────────────────────────────────────
    if not run:
        col_a, col_b = st.columns([3, 2])
        with col_a:
            st.info("👈 Configura en el panel izquierdo y presiona **🚀 Ejecutar analisis**.\n\n"
                    "Al subir un CSV/Excel, los tickers se detectan automaticamente.")
            with st.expander("📋 Formato esperado del CSV/Excel"):
                st.markdown(
                    "**Primera columna**: Fechas. **Columnas siguientes**: una serie "
                    "por columna con el nombre del ticker como encabezado."
                )
                sp = pd.DataFrame({
                    "SPX Index":[3257.85,3265.35,2480.64,2304.92],
                    "USGG10YR Index":[1.88,1.90,0.73,0.76],
                    "USDMXN Curncy":[18.87,18.90,24.51,23.04],
                    "XAU Curncy":[1520.0,1547.8,1477.2,1680.0],
                }, index=pd.to_datetime(["2020-01-02","2020-01-03","2020-03-16","2020-04-01"]))
                sp.index.name = "Date"
                st.dataframe(sp, use_container_width=True)
                st.download_button("⬇️ CSV de muestra (2015-2024)", generate_sample_csv(),
                                   "sample_event_study.csv", "text/csv", use_container_width=True)
        with col_b:
            st.markdown("#### 💡 Eventos de referencia")
            for ev in REFERENCE_EVENTS[:12]:
                st.markdown(f"- **{ev['label']}** — `{ev['date']}`")
            st.caption(f"...y {len(REFERENCE_EVENTS)-12} mas. Usa el boton 📋 Catalogo en la sidebar.")
        return

    # ══════════════════════════════════════════════════════════════════════════
    # VALIDACION
    # ══════════════════════════════════════════════════════════════════════════
    events  = st.session_state.events
    tickers = st.session_state.tickers

    valid_events = []
    for ev in events:
        try:
            pd.Timestamp(ev["date"])
            valid_events.append(ev)
        except Exception:
            st.warning(f"⚠️ Fecha invalida: '{ev['label']}' -> '{ev['date']}'")
    if not valid_events:
        st.error("❌ Ningun evento tiene fecha valida.")
        return
    if not tickers:
        st.error("❌ Agrega al menos un ticker (sube un archivo o agrega manualmente).")
        return

    # ══════════════════════════════════════════════════════════════════════════
    # CARGA BLOOMBERG
    # ══════════════════════════════════════════════════════════════════════════
    if use_bloomberg:
        fields_map = {tk["ticker"]: tk.get("bloomberg_field","PX_LAST") for tk in tickers}
        all_ev_ts = [pd.Timestamp(ev["date"]) for ev in valid_events]
        buf_map = {
            "daily":   (timedelta(days=lookback+10),      timedelta(days=lookforward+10)),
            "weekly":  (timedelta(weeks=lookback+2),       timedelta(weeks=lookforward+2)),
            "monthly": (relativedelta(months=lookback+1),  relativedelta(months=lookforward+1)),
            "annual":  (relativedelta(years=lookback+1),   relativedelta(years=lookforward+1)),
        }
        bb, bf = buf_map.get(frequency, (timedelta(days=90), timedelta(days=90)))
        ns = (min(all_ev_ts) - bb).date()
        ne = (max(all_ev_ts) + bf).date()
        actual_s = min(bbg_start, ns)
        actual_e = max(bbg_end, ne)
        with st.spinner("🔄 Descargando de Bloomberg..."):
            data_df, err = load_from_bloomberg(fields_map, actual_s, actual_e)
        if err:
            st.error(f"❌ {err}")
            return
        if data_df is None or data_df.empty:
            st.error("❌ Bloomberg no devolvio datos.")
            return
        st.success(f"✅ Bloomberg: {len(data_df.columns)} series · {len(data_df):,} obs")

    elif data_df is None:
        st.error("❌ Sube un archivo CSV o Excel primero.")
        return

    # ══════════════════════════════════════════════════════════════════════════
    # PROCESAMIENTO
    # ══════════════════════════════════════════════════════════════════════════
    st.markdown("---")
    st.markdown("## 📊 Resultados")

    for tk_idx, tk in enumerate(tickers):
        ticker_id     = tk["ticker"]
        display_name  = tk.get("display_name") or ticker_id
        field_mode    = tk.get("field_mode",    "cumulative_return")
        baseline_mode = tk.get("baseline_mode", "base100")

        # Obtener serie
        series = None
        if not use_bloomberg:
            col = tk.get("csv_column", "")
            if not col and data_df is not None and len(data_df.columns) > 0:
                col = data_df.columns[0]
            if col and data_df is not None and col in data_df.columns:
                series = data_df[col].dropna()
            else:
                st.error(f"❌ Columna **'{col}'** no encontrada para **'{display_name}'**.")
                continue
        else:
            if data_df is not None and ticker_id in data_df.columns:
                series = data_df[ticker_id].dropna()
            else:
                st.error(f"❌ Sin datos Bloomberg para **'{ticker_id}'**.")
                continue
        if series is None or len(series) == 0:
            st.warning(f"⚠️ Serie vacia para '{display_name}'.")
            continue

        # Ventanas alineadas
        aligned_data: dict[str, dict] = {}
        skipped: list[str] = []

        for ev in valid_events:
            ev_label = ev["label"]
            try:
                raw = extract_aligned_values(series, ev["date"], lookback, lookforward, frequency)
                if pd.isna(raw.get(0, np.nan)):
                    skipped.append(ev_label)
                    continue
                aligned_data[ev_label] = apply_transformation(raw, field_mode, baseline_mode)
            except Exception as e:
                st.warning(f"⚠️ '{ev_label}' — '{display_name}': {e}")

        if skipped:
            st.caption(f"⚠️ Sin datos en anchor para '{display_name}': {', '.join(skipped)}")
        if not aligned_data:
            st.error(f"❌ Sin datos procesados para **'{display_name}'**.")
            continue

        # ── Grafica ───────────────────────────────────────────────────────────
        fig = create_event_chart(
            ticker_id, display_name, valid_events,
            aligned_data, field_mode, baseline_mode, frequency, show_avg=show_avg,
        )
        st.plotly_chart(fig, use_container_width=True)

        # ── Botones de export ─────────────────────────────────────────────────
        exp_c1, exp_c2, exp_c3 = st.columns(3)
        safe_name = ticker_id.replace(" ", "_").replace("/", "_")

        with exp_c1:
            html_bytes = fig.to_html(include_plotlyjs="cdn").encode("utf-8")
            st.download_button(
                "📥 Descargar HTML interactivo",
                html_bytes, f"{safe_name}_chart.html", "text/html",
                key=f"html_{tk_idx}",
                use_container_width=True,
            )
        with exp_c2:
            if KALEIDO_AVAILABLE:
                png_bytes = fig.to_image(format="png", width=1400, height=550, scale=2)
                st.download_button(
                    "📥 Descargar PNG",
                    png_bytes, f"{safe_name}_chart.png", "image/png",
                    key=f"png_{tk_idx}",
                    use_container_width=True,
                )
            else:
                st.caption("Para PNG: `pip install kaleido`")
        with exp_c3:
            # CSV de datos
            all_periods   = sorted({p for d in aligned_data.values() for p in d})
            period_labels = build_period_labels(all_periods, frequency)
            rows = {el: [ed.get(p, np.nan) for p in all_periods]
                    for el, ed in aligned_data.items()}
            result_df = pd.DataFrame(rows, index=period_labels)
            result_df.index.name = "Periodo"
            st.download_button(
                "📥 Descargar datos CSV",
                result_df.to_csv(), f"{safe_name}_data.csv", "text/csv",
                key=f"csv_{tk_idx}",
                use_container_width=True,
            )

        # ── Tabla expandible ──────────────────────────────────────────────────
        with st.expander(f"📋 Datos — {display_name}"):
            all_periods   = sorted({p for d in aligned_data.values() for p in d})
            period_labels = build_period_labels(all_periods, frequency)
            rows = {el: [ed.get(p, np.nan) for p in all_periods]
                    for el, ed in aligned_data.items()}
            result_df = pd.DataFrame(rows, index=period_labels)
            result_df.index.name = "Periodo"
            st.dataframe(
                result_df.style.format("{:.4f}", na_rep="—"),
                use_container_width=True,
            )

    st.success("✅ Analisis completado.")


if __name__ == "__main__":
    main()
