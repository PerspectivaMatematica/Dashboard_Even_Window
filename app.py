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
from pathlib import Path
from io import BytesIO
import io
import os

# ── Bloomberg (opcional) ──────────────────────────────────────────────────────
try:
    import blpapi
    BLOOMBERG_AVAILABLE = True
except ImportError:
    BLOOMBERG_AVAILABLE = False

# ── Kaleido para export PNG (opcional) ────────────────────────────────────────
KALEIDO_AVAILABLE = False
try:
    import kaleido  # noqa: F401
    # Verificar que realmente puede renderizar (necesita Chrome en el sistema)
    import plotly.io as pio
    _test_fig = go.Figure(data=[go.Scatter(x=[0], y=[0])])
    _test_fig.to_image(format="png", width=100, height=100)
    KALEIDO_AVAILABLE = True
    del _test_fig
except Exception:
    KALEIDO_AVAILABLE = False


# ══════════════════════════════════════════════════════════════════════════════
# CONSTANTES
# ══════════════════════════════════════════════════════════════════════════════

ASSET_TYPES    = ["equity", "rate", "fx", "commodity", "other"]
FIELD_MODES    = ["cumulative_return", "absolute_change", "pct_change", "price_indexed", "price"]
BASELINE_MODES = ["base100", "base0", "none"]
FREQUENCIES    = ["daily", "weekly", "monthly", "annual"]
AGG_METHODS    = ["last", "average"]

FREQ_LABEL = {"daily": "D", "weekly": "W", "monthly": "M", "annual": "Y"}

FIELD_MODE_LABELS = {
    "price":             "Precio Raw (nivel real)",
    "price_indexed":     "Precio Indexado (Base 100)",
    "cumulative_return": "Retorno Acumulado (%)",
    "pct_change":        "Cambio Porcentual (%)",
    "absolute_change":   "Cambio Absoluto",
}

BASELINE_LABELS = {
    "base100": "Base 100 (indexado al evento)",
    "base0":   "Base 0 (cambio desde el evento)",
    "none":    "Sin normalizar (nivel real)",
}

AGG_METHOD_LABELS = {
    "last":    "Ultimo dato disponible",
    "average": "Promedio del periodo",
}

ASSET_DEFAULTS = {
    "equity":    {"field_mode": "price_indexed", "baseline_mode": "base100"},
    "rate":      {"field_mode": "absolute_change",   "baseline_mode": "base0"},
    "fx":        {"field_mode": "price_indexed", "baseline_mode": "base100"},
    "commodity": {"field_mode": "price_indexed", "baseline_mode": "base100"},
    "other":     {"field_mode": "price",             "baseline_mode": "none"},
}

# Paleta corporativa (basada en colors_afore)
PALETTE = {
    "rojo":               "#FF1B44",
    "azul":               "#003746",
    "azul_digital":       "#009CC6",
    "azul_digital_oscuro":"#001E22",
    "azul_digital_claro": "#005162",
    "granate":            "#601636",
    "granate_claro":      "#B77493",
    "violeta_oscuro":     "#B18DFB",
    "naranja_oscuro":     "#FF5F00",
    "naranja":            "#FA8D5A",
    "verde_oscuro":       "#00AD59",
    "verde":              "#2DDC8E",
    "azul_grisaceo":      "#A0D6E2",
    "crema":              "#FDE8E0",
}

# Colores para lineas de eventos (orden pensado para buen contraste)
EVENT_COLORS = [
    "#003746",   # Azul
    "#FF1B44",   # Rojo
    "#00AD59",   # Verde Oscuro
    "#FF5F00",   # Naranja Oscuro
    "#B18DFB",   # Violeta Oscuro
    "#009CC6",   # Azul Digital
    "#601636",   # Granate
    "#FA8D5A",   # Naranja
    "#2DDC8E",   # Verde
    "#B77493",   # Granate Claro
    "#005162",   # Azul Digital Claro
    "#A0D6E2",   # Azul Grisaceo
]

# ══════════════════════════════════════════════════════════════════════════════
# TABS PRESETS — Configuraciones pre-armadas para analisis rapido via Bloomberg
# ══════════════════════════════════════════════════════════════════════════════

_EVENTS_GLOBAL = [
    {"label": "COVID Crash",              "date": "2020-03-16"},
    {"label": "GFC (Lehman)",             "date": "2008-09-15"},
    {"label": "9/11 Attacks",             "date": "2001-09-11"},
    {"label": "Russia-Ukraine",           "date": "2022-02-24"},
    {"label": "SVB Collapse",             "date": "2023-03-10"},
    {"label": "Israel-Hamas",             "date": "2023-10-07"},
    {"label": "Trump Tariffs 2025",       "date": "2025-04-02"},
]

TAB_PRESETS = {
    "General": {
        "icon": "🌎",
        "description": "Panorama macro: equity, tasas, FX y commodities",
        "tickers": [
            # Equity
            {"ticker": "SPX Index",      "display_name": "S&P 500",       "bloomberg_field": "PX_LAST", "asset_type": "equity",    "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "NDX Index",      "display_name": "Nasdaq 100",    "bloomberg_field": "PX_LAST", "asset_type": "equity",    "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "VIX Index",      "display_name": "VIX",           "bloomberg_field": "PX_LAST", "asset_type": "other",     "field_mode": "price",         "baseline_mode": "none"},
            {"ticker": "SX5E Index",     "display_name": "Euro Stoxx 50", "bloomberg_field": "PX_LAST", "asset_type": "equity",    "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "MEXBOL Index",   "display_name": "IPC Mexico",    "bloomberg_field": "PX_LAST", "asset_type": "equity",    "field_mode": "price_indexed", "baseline_mode": "base100"},
            # Tasas nominales
            {"ticker": "GT2 Govt",        "display_name": "US 2Y",          "bloomberg_field": "PX_LAST", "asset_type": "rate",      "field_mode": "absolute_change","baseline_mode": "base0"},
            {"ticker": "GT10 Govt",       "display_name": "US 10Y",         "bloomberg_field": "PX_LAST", "asset_type": "rate",      "field_mode": "absolute_change","baseline_mode": "base0"},
            {"ticker": "GTMXN2Y Govt",    "display_name": "MX 2Y",          "bloomberg_field": "PX_LAST", "asset_type": "rate",      "field_mode": "absolute_change","baseline_mode": "base0"},
            {"ticker": "GTMXN10Y Govt",   "display_name": "Mbono 10Y",      "bloomberg_field": "PX_LAST", "asset_type": "rate",      "field_mode": "absolute_change","baseline_mode": "base0"},
            {"ticker": "GDBR10 Index",    "display_name": "Bund 10Y",       "bloomberg_field": "PX_LAST", "asset_type": "rate",      "field_mode": "absolute_change","baseline_mode": "base0"},
            # Tasas reales & breakeven
            {"ticker": "GTII10 Govt",     "display_name": "US TIPS 10Y",    "bloomberg_field": "PX_LAST", "asset_type": "rate",      "field_mode": "absolute_change","baseline_mode": "base0"},
            {"ticker": "USGGBE10 Index",  "display_name": "US BE 10Y",      "bloomberg_field": "PX_LAST", "asset_type": "rate",      "field_mode": "absolute_change","baseline_mode": "base0"},
            {"ticker": "GTMXNII10Y Govt", "display_name": "MX Real 10Y",    "bloomberg_field": "PX_LAST", "asset_type": "rate",      "field_mode": "absolute_change","baseline_mode": "base0"},
            # FX & Commodities
            {"ticker": "USDMXN Curncy",  "display_name": "USD/MXN",       "bloomberg_field": "PX_LAST", "asset_type": "fx",        "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "DXY Index",      "display_name": "DXY",           "bloomberg_field": "PX_LAST", "asset_type": "fx",        "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "EURUSD Curncy",  "display_name": "EUR/USD",       "bloomberg_field": "PX_LAST", "asset_type": "fx",        "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "XAU Curncy",     "display_name": "Oro",           "bloomberg_field": "PX_LAST", "asset_type": "commodity", "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "XAG Curncy",     "display_name": "Plata",         "bloomberg_field": "PX_LAST", "asset_type": "commodity", "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "CL1 Comdty",     "display_name": "WTI Crudo",     "bloomberg_field": "PX_LAST", "asset_type": "commodity", "field_mode": "price_indexed", "baseline_mode": "base100"},
        ],
        "events": _EVENTS_GLOBAL,
    },
    "Sectores S&P": {
        "icon": "📊",
        "description": "11 sectores GICS del S&P 500",
        "tickers": [
            {"ticker": "S5INFT Index", "display_name": "Tecnologia",          "bloomberg_field": "PX_LAST", "asset_type": "equity", "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "S5FINL Index", "display_name": "Financieros",         "bloomberg_field": "PX_LAST", "asset_type": "equity", "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "S5HLTH Index", "display_name": "Salud",               "bloomberg_field": "PX_LAST", "asset_type": "equity", "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "S5COND Index", "display_name": "Consumo Discrecional","bloomberg_field": "PX_LAST", "asset_type": "equity", "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "S5CONS Index", "display_name": "Consumo Basico",      "bloomberg_field": "PX_LAST", "asset_type": "equity", "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "S5ENRS Index", "display_name": "Energia",             "bloomberg_field": "PX_LAST", "asset_type": "equity", "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "S5INDU Index", "display_name": "Industriales",        "bloomberg_field": "PX_LAST", "asset_type": "equity", "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "S5MATR Index", "display_name": "Materiales",          "bloomberg_field": "PX_LAST", "asset_type": "equity", "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "S5TELS Index", "display_name": "Comunicaciones",      "bloomberg_field": "PX_LAST", "asset_type": "equity", "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "S5UTIL Index", "display_name": "Utilities",           "bloomberg_field": "PX_LAST", "asset_type": "equity", "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "S5RLST Index", "display_name": "Real Estate",         "bloomberg_field": "PX_LAST", "asset_type": "equity", "field_mode": "price_indexed", "baseline_mode": "base100"},
        ],
        "events": _EVENTS_GLOBAL,
    },
    "Tasas US & MX": {
        "icon": "📈",
        "description": "Nominales, reales (TIPS/UDIBONOS) y breakevens US & MX",
        "tickers": [
            # ── US Nominales ──
            {"ticker": "GT2 Govt",        "display_name": "US Nom 2Y",      "bloomberg_field": "PX_LAST", "asset_type": "rate", "field_mode": "absolute_change", "baseline_mode": "base0"},
            {"ticker": "GT5 Govt",        "display_name": "US Nom 5Y",      "bloomberg_field": "PX_LAST", "asset_type": "rate", "field_mode": "absolute_change", "baseline_mode": "base0"},
            {"ticker": "GT10 Govt",       "display_name": "US Nom 10Y",     "bloomberg_field": "PX_LAST", "asset_type": "rate", "field_mode": "absolute_change", "baseline_mode": "base0"},
            {"ticker": "GT20 Govt",       "display_name": "US Nom 20Y",     "bloomberg_field": "PX_LAST", "asset_type": "rate", "field_mode": "absolute_change", "baseline_mode": "base0"},
            {"ticker": "GT30 Govt",       "display_name": "US Nom 30Y",     "bloomberg_field": "PX_LAST", "asset_type": "rate", "field_mode": "absolute_change", "baseline_mode": "base0"},
            # ── US Reales (TIPS) ──
            {"ticker": "USGGT02Y Index",  "display_name": "US TIPS 2Y",    "bloomberg_field": "PX_LAST", "asset_type": "rate", "field_mode": "absolute_change", "baseline_mode": "base0"},
            {"ticker": "GTII5 Govt",      "display_name": "US TIPS 5Y",    "bloomberg_field": "PX_LAST", "asset_type": "rate", "field_mode": "absolute_change", "baseline_mode": "base0"},
            {"ticker": "GTII10 Govt",     "display_name": "US TIPS 10Y",   "bloomberg_field": "PX_LAST", "asset_type": "rate", "field_mode": "absolute_change", "baseline_mode": "base0"},
            {"ticker": "GTII20 Govt",     "display_name": "US TIPS 20Y",   "bloomberg_field": "PX_LAST", "asset_type": "rate", "field_mode": "absolute_change", "baseline_mode": "base0"},
            {"ticker": "GTII30 Govt",     "display_name": "US TIPS 30Y",   "bloomberg_field": "PX_LAST", "asset_type": "rate", "field_mode": "absolute_change", "baseline_mode": "base0"},
            # ── US Breakevens ──
            {"ticker": "USGGBE02 Index",  "display_name": "US BE 2Y",      "bloomberg_field": "PX_LAST", "asset_type": "rate", "field_mode": "absolute_change", "baseline_mode": "base0"},
            {"ticker": "USGGBE05 Index",  "display_name": "US BE 5Y",      "bloomberg_field": "PX_LAST", "asset_type": "rate", "field_mode": "absolute_change", "baseline_mode": "base0"},
            {"ticker": "USGGBE10 Index",  "display_name": "US BE 10Y",     "bloomberg_field": "PX_LAST", "asset_type": "rate", "field_mode": "absolute_change", "baseline_mode": "base0"},
            # ── MX Nominales ──
            {"ticker": "MXIBTIIE Index",  "display_name": "TIIE 28d",      "bloomberg_field": "PX_LAST", "asset_type": "rate", "field_mode": "absolute_change", "baseline_mode": "base0"},
            {"ticker": "GTMXN2Y Govt",    "display_name": "MX Nom 2Y",    "bloomberg_field": "PX_LAST", "asset_type": "rate", "field_mode": "absolute_change", "baseline_mode": "base0"},
            {"ticker": "GTMXN5Y Govt",    "display_name": "MX Nom 5Y",    "bloomberg_field": "PX_LAST", "asset_type": "rate", "field_mode": "absolute_change", "baseline_mode": "base0"},
            {"ticker": "GTMXN10Y Govt",   "display_name": "MX Nom 10Y",   "bloomberg_field": "PX_LAST", "asset_type": "rate", "field_mode": "absolute_change", "baseline_mode": "base0"},
            {"ticker": "GTMXN20Y Govt",   "display_name": "MX Nom 20Y",   "bloomberg_field": "PX_LAST", "asset_type": "rate", "field_mode": "absolute_change", "baseline_mode": "base0"},
            {"ticker": "GTMXN30Y Govt",   "display_name": "MX Nom 30Y",   "bloomberg_field": "PX_LAST", "asset_type": "rate", "field_mode": "absolute_change", "baseline_mode": "base0"},
            # ── MX Reales (UDIBONOS) ──
            {"ticker": "GTMXNII5Y Govt",  "display_name": "MX Real 5Y",   "bloomberg_field": "PX_LAST", "asset_type": "rate", "field_mode": "absolute_change", "baseline_mode": "base0"},
            {"ticker": "GTMXNII10Y Govt", "display_name": "MX Real 10Y",  "bloomberg_field": "PX_LAST", "asset_type": "rate", "field_mode": "absolute_change", "baseline_mode": "base0"},
            {"ticker": "GTMXNII20Y Govt", "display_name": "MX Real 20Y",  "bloomberg_field": "PX_LAST", "asset_type": "rate", "field_mode": "absolute_change", "baseline_mode": "base0"},
            {"ticker": "GTMXNII30Y Govt", "display_name": "MX Real 30Y",  "bloomberg_field": "PX_LAST", "asset_type": "rate", "field_mode": "absolute_change", "baseline_mode": "base0"},
        ],
        "events": _EVENTS_GLOBAL + [
            {"label": "Taper Tantrum",            "date": "2013-05-22"},
            {"label": "Fed Pivot Dic 2023",       "date": "2023-12-13"},
        ],
    },
    "FX": {
        "icon": "💱",
        "description": "Pares principales y emergentes",
        "tickers": [
            {"ticker": "USDMXN Curncy", "display_name": "USD/MXN", "bloomberg_field": "PX_LAST", "asset_type": "fx", "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "EURUSD Curncy", "display_name": "EUR/USD", "bloomberg_field": "PX_LAST", "asset_type": "fx", "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "USDJPY Curncy", "display_name": "USD/JPY", "bloomberg_field": "PX_LAST", "asset_type": "fx", "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "GBPUSD Curncy", "display_name": "GBP/USD", "bloomberg_field": "PX_LAST", "asset_type": "fx", "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "DXY Index",     "display_name": "DXY",     "bloomberg_field": "PX_LAST", "asset_type": "fx", "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "USDBRL Curncy", "display_name": "USD/BRL", "bloomberg_field": "PX_LAST", "asset_type": "fx", "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "USDCNH Curncy", "display_name": "USD/CNH", "bloomberg_field": "PX_LAST", "asset_type": "fx", "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "USDZAR Curncy", "display_name": "USD/ZAR", "bloomberg_field": "PX_LAST", "asset_type": "fx", "field_mode": "price_indexed", "baseline_mode": "base100"},
        ],
        "events": _EVENTS_GLOBAL + [
            {"label": "Brexit Referendum",        "date": "2016-06-24"},
            {"label": "Eleccion MX 2024",         "date": "2024-06-02"},
        ],
    },
    "Commodities": {
        "icon": "🛢️",
        "description": "Energeticos, metales preciosos, industriales y agricolas",
        "tickers": [
            {"ticker": "CL1 Comdty",  "display_name": "WTI Crudo",    "bloomberg_field": "PX_LAST", "asset_type": "commodity", "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "CO1 Comdty",  "display_name": "Brent",        "bloomberg_field": "PX_LAST", "asset_type": "commodity", "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "NG1 Comdty",  "display_name": "Gas Natural",  "bloomberg_field": "PX_LAST", "asset_type": "commodity", "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "XAU Curncy",  "display_name": "Oro",          "bloomberg_field": "PX_LAST", "asset_type": "commodity", "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "XAG Curncy",  "display_name": "Plata",        "bloomberg_field": "PX_LAST", "asset_type": "commodity", "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "HG1 Comdty",  "display_name": "Cobre",        "bloomberg_field": "PX_LAST", "asset_type": "commodity", "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "LA1 Comdty",  "display_name": "Aluminio",     "bloomberg_field": "PX_LAST", "asset_type": "commodity", "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "LIT US Equity","display_name": "Litio (ETF)", "bloomberg_field": "PX_LAST", "asset_type": "equity",    "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "C 1 Comdty",  "display_name": "Maiz",         "bloomberg_field": "PX_LAST", "asset_type": "commodity", "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "W 1 Comdty",  "display_name": "Trigo",        "bloomberg_field": "PX_LAST", "asset_type": "commodity", "field_mode": "price_indexed", "baseline_mode": "base100"},
        ],
        "events": _EVENTS_GLOBAL + [
            {"label": "OPEC+ Recorte",            "date": "2020-04-12"},
            {"label": "Guerra del Golfo",         "date": "1990-08-02"},
        ],
    },
    "Europa & EM": {
        "icon": "🌍",
        "description": "Europa, mercados emergentes y riesgo soberano",
        "tickers": [
            {"ticker": "SX5E Index",      "display_name": "Euro Stoxx 50",  "bloomberg_field": "PX_LAST", "asset_type": "equity", "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "DAX Index",       "display_name": "DAX",            "bloomberg_field": "PX_LAST", "asset_type": "equity", "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "UKX Index",       "display_name": "FTSE 100",       "bloomberg_field": "PX_LAST", "asset_type": "equity", "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "GDBR10 Index",    "display_name": "Bund 10Y",       "bloomberg_field": "PX_LAST", "asset_type": "rate",   "field_mode": "absolute_change","baseline_mode": "base0"},
            {"ticker": "MEXBOL Index",    "display_name": "IPC Mexico",     "bloomberg_field": "PX_LAST", "asset_type": "equity", "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "IBOV Index",      "display_name": "Bovespa",        "bloomberg_field": "PX_LAST", "asset_type": "equity", "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "SHCOMP Index",    "display_name": "Shanghai Comp",  "bloomberg_field": "PX_LAST", "asset_type": "equity", "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "MXEF Index",      "display_name": "MSCI EM",        "bloomberg_field": "PX_LAST", "asset_type": "equity", "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "NIFTY Index",     "display_name": "Nifty 50",       "bloomberg_field": "PX_LAST", "asset_type": "equity", "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "HSI Index",       "display_name": "Hang Seng",      "bloomberg_field": "PX_LAST", "asset_type": "equity", "field_mode": "price_indexed",  "baseline_mode": "base100"},
        ],
        "events": _EVENTS_GLOBAL + [
            {"label": "Crisis Deuda EU",          "date": "2011-08-05"},
            {"label": "Devaluacion CNY",          "date": "2015-08-11"},
        ],
    },
}

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

def get_excel_sheet_names(uploaded_file) -> list[str]:
    """Devuelve la lista de hojas de un Excel. Lista vacia si es CSV."""
    name = uploaded_file.name.lower()
    if not name.endswith((".xlsx", ".xls")):
        return []
    try:
        uploaded_file.seek(0)
        xls = pd.ExcelFile(uploaded_file)
        sheets = xls.sheet_names
        uploaded_file.seek(0)
        return sheets
    except Exception:
        uploaded_file.seek(0)
        return []


def load_from_file(uploaded_file, sheet_name=None):
    """Carga datos de precios desde CSV o Excel subido por el usuario."""
    try:
        name = uploaded_file.name.lower()
        if name.endswith((".xlsx", ".xls")):
            uploaded_file.seek(0)
            df = pd.read_excel(uploaded_file, index_col=0, parse_dates=True,
                               sheet_name=sheet_name or 0)
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
            req.set("startDate", start_date.strftime("%Y%m%d"))
            req.set("endDate",   end_date.strftime("%Y%m%d"))
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
# CACHE MAESTRO BLOOMBERG
# ══════════════════════════════════════════════════════════════════════════════

# Directorio donde vive app.py para guardar el cache junto al proyecto
_APP_DIR = Path(__file__).resolve().parent
_CACHE_DIR = _APP_DIR / ".bbg_cache"

# ── Carpeta data/ con archivos Excel default del repo ────────────────────────
# Estos archivos son los que se versionan en GitHub y sirven como fuente de
# datos cuando el usuario NO tiene Bloomberg Terminal disponible.
DATA_DIR = _APP_DIR / "data"
# Archivo default que se sobrescribe al correr la app localmente con Bloomberg.
# Cambia este nombre si prefieres otro archivo como "fuente de verdad".
# (Si lo cambias, asegurate de que el archivo exista en data/ o que sea creado al correr local)
DEFAULT_DATA_FILENAME = "Data_historica.xlsx"
DEFAULT_DATA_PATH = DATA_DIR / DEFAULT_DATA_FILENAME


def list_default_data_files() -> list[Path]:
    """Lista los archivos .xlsx/.csv dentro de data/ (ordenados alfabeticamente)."""
    if not DATA_DIR.exists():
        return []
    files = []
    for ext in ("*.xlsx", "*.xls", "*.csv"):
        files.extend(DATA_DIR.glob(ext))
    return sorted(files, key=lambda p: p.name.lower())


def load_default_data_file(path: Path, sheet_name=None):
    """Carga un archivo (CSV o Excel) desde la carpeta data/ del repo.

    Auto-normaliza headers: si las columnas son display_names (ej. "S&P 500"),
    los convierte a tickers Bloomberg (ej. "SPX Index") usando el mapeo de los
    presets. Esto permite que el analisis por categoria funcione aunque el
    archivo se haya guardado con nombres legibles.
    """
    try:
        name = str(path).lower()
        if name.endswith((".xlsx", ".xls")):
            df = pd.read_excel(path, index_col=0, parse_dates=True,
                               sheet_name=sheet_name or 0)
        else:
            try:
                df = pd.read_csv(path, index_col=0, parse_dates=True)
            except Exception:
                df = pd.read_csv(path, index_col=0, parse_dates=True, sep=";")
        df.index = pd.to_datetime(df.index, errors="coerce")
        df = df[df.index.notna()].sort_index()
        df = df[~df.index.duplicated(keep="last")]
        df = df.apply(pd.to_numeric, errors="coerce")

        # ── Normalizacion: display_name → ticker ──
        # Si el archivo se guardo con nombres legibles (bug previo o intencional),
        # convertimos las columnas a tickers Bloomberg para que los presets matcheen.
        try:
            ticker_to_name = _get_all_preset_names()  # {ticker: display_name}
            name_to_ticker = {v: k for k, v in ticker_to_name.items()}
            rename_back = {c: name_to_ticker[c] for c in df.columns
                           if c in name_to_ticker and name_to_ticker[c] not in df.columns}
            if rename_back:
                df = df.rename(columns=rename_back)
        except Exception:
            pass  # Si falla la normalizacion, dejamos el df como esta.

        return df, None
    except Exception as e:
        return None, str(e)


def get_local_default_sheets(path: Path) -> list[str]:
    """Hojas disponibles en un Excel local. Lista vacia si es CSV."""
    name = str(path).lower()
    if not name.endswith((".xlsx", ".xls")):
        return []
    try:
        return pd.ExcelFile(path).sheet_names
    except Exception:
        return []


def save_data_to_default_excel(df: pd.DataFrame, target_path: Path = None,
                                rename_map: dict = None) -> tuple[bool, str]:
    """Guarda el DataFrame de Bloomberg al archivo Excel default del repo.

    IMPORTANTE: por defecto NO renombra columnas — mantiene los tickers
    Bloomberg (ej. "SPX Index", "GT10 Govt") como headers para que el analisis
    por categoria pueda hacer lookup por ticker. Solo pasa rename_map si quieres
    nombres legibles (NO recomendado para el archivo default).

    Solo se llama cuando blpapi esta disponible (= corriendo localmente).
    En Streamlit Cloud el filesystem es read-only y esta funcion no se invoca.
    """
    if target_path is None:
        target_path = DEFAULT_DATA_PATH
    try:
        target_path.parent.mkdir(parents=True, exist_ok=True)
        export_df = df.copy()
        if rename_map:
            export_df = export_df.rename(columns={k: v for k, v in rename_map.items()
                                                  if k in export_df.columns})
        export_df.index.name = "Date"
        with pd.ExcelWriter(str(target_path), engine="openpyxl") as writer:
            export_df.to_excel(writer, sheet_name="Bloomberg Data")
        return True, str(target_path)
    except Exception as e:
        return False, f"{type(e).__name__}: {e}"

def _get_all_preset_tickers() -> dict[str, str]:
    """Recopila todos los tickers únicos de TODOS los TAB_PRESETS.
    Retorna {ticker: bloomberg_field}."""
    all_tk: dict[str, str] = {}
    for _pval in TAB_PRESETS.values():
        for _tk in _pval["tickers"]:
            t_id = _tk["ticker"]
            if t_id not in all_tk:
                all_tk[t_id] = _tk.get("bloomberg_field", "PX_LAST")
    return all_tk

def _get_all_preset_names() -> dict[str, str]:
    """Recopila {ticker: display_name} de todos los presets."""
    names: dict[str, str] = {}
    for _pval in TAB_PRESETS.values():
        for _tk in _pval["tickers"]:
            t_id = _tk["ticker"]
            if t_id not in names:
                names[t_id] = _tk.get("display_name", t_id)
    return names

def _cache_path_for_date(d: date) -> Path:
    """Retorna la ruta del archivo cache para una fecha dada."""
    _CACHE_DIR.mkdir(exist_ok=True)
    return _CACHE_DIR / f"master_{d.strftime('%Y%m%d')}.pkl"

def load_master_cache(start_date, end_date, force_refresh: bool = False):
    """Carga o crea el cache maestro con todos los tickers de presets.

    Returns:
        (DataFrame completo, error_str o None)
    """
    today = date.today()
    cache_file = _cache_path_for_date(today)

    # Si ya existe el cache de hoy y no se fuerza refresh, leer
    if cache_file.exists() and not force_refresh:
        try:
            df = pd.read_pickle(cache_file)
            df.index = pd.DatetimeIndex(df.index)
            # Verificar si hay tickers nuevos que no estan en el cache
            all_tk = _get_all_preset_tickers()
            missing = {t: f for t, f in all_tk.items() if t not in df.columns}
            if missing:
                st.info(f"🔄 Descargando {len(missing)} tickers nuevos...")
                extra_df, err = load_from_bloomberg(missing, start_date, end_date)
                if extra_df is not None and not extra_df.empty:
                    df = pd.concat([df, extra_df], axis=1)
                    df.to_pickle(cache_file)
            return df, None
        except Exception as e:
            st.warning(f"⚠️ Error leyendo cache, re-descargando: {e}")

    # Descargar todo desde Bloomberg
    all_tk = _get_all_preset_tickers()
    st.info(f"📡 Descargando {len(all_tk)} tickers de todas las categorias...")
    df, err = load_from_bloomberg(all_tk, start_date, end_date)
    if err:
        return None, err
    if df is None or df.empty:
        return None, "Bloomberg no devolvio datos."

    # Guardar cache
    try:
        df.to_pickle(cache_file)
    except Exception:
        pass  # Si no puede guardar, no pasa nada — funciona sin cache

    # Limpiar caches antiguos (mantener solo ultimos 5 dias)
    try:
        for old_file in sorted(_CACHE_DIR.glob("master_*.pkl"))[:-5]:
            old_file.unlink()
    except Exception:
        pass

    return df, None


# ══════════════════════════════════════════════════════════════════════════════
# MOTOR DE VENTANAS
# ══════════════════════════════════════════════════════════════════════════════

def build_target_dates(event_date, lookback, lookforward, frequency):
    """Genera lista de (periodo_relativo, fecha_objetivo)."""
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
    """
    Devuelve la fecha mas cercana anterior a target_ts.
    Si target_ts esta fuera del rango de la serie (antes del primer dato
    o despues del ultimo dato), devuelve None → NaN, sin repetir valores.
    """
    if len(series_index) == 0:
        return None
    # Si la fecha objetivo es posterior al ultimo dato, NO hay observacion
    if target_ts > series_index[-1]:
        return None
    # Si la fecha objetivo es anterior al primer dato, NO hay observacion
    if target_ts < series_index[0]:
        return None
    # Dentro del rango: buscar la observacion previa mas cercana
    # (maneja fines de semana, feriados, gaps normales)
    prior = series_index[series_index <= target_ts]
    if len(prior) == 0:
        return None
    return prior[-1]


def get_period_average(series, period_start, period_end):
    """Calcula el promedio de la serie entre period_start y period_end (inclusive)."""
    mask = (series.index >= period_start) & (series.index <= period_end)
    subset = series[mask]
    if len(subset) == 0:
        return np.nan
    return subset.mean()


def extract_aligned_values(series, event_date, lookback, lookforward, frequency,
                           agg_method="last"):
    """
    Extrae valores alineados para un evento.
    agg_method:
      - "last"    : ultimo dato disponible antes de la fecha objetivo (default)
      - "average" : promedio de todas las observaciones dentro del periodo
    """
    target_dates = build_target_dates(event_date, lookback, lookforward, frequency)
    result = {}

    if agg_method == "average" and frequency != "daily":
        # Para promedio: calcular la media entre dos fechas objetivo consecutivas
        for i, (period, target_ts) in enumerate(target_dates):
            if i == 0:
                prev_ts = target_dates[0][1] - pd.Timedelta(days=1)
            else:
                prev_ts = target_dates[i - 1][1]
            avg_val = get_period_average(series, prev_ts + pd.Timedelta(days=1), target_ts)
            if pd.isna(avg_val):
                mapped = get_prior_observation(target_ts, series.index)
                result[period] = series.loc[mapped] if mapped is not None else np.nan
            else:
                result[period] = avg_val
    else:
        # Metodo default: ultimo dato disponible
        for period, target_ts in target_dates:
            mapped = get_prior_observation(target_ts, series.index)
            result[period] = series.loc[mapped] if mapped is not None else np.nan

    return result


# ══════════════════════════════════════════════════════════════════════════════
# TRANSFORMACIONES
# ══════════════════════════════════════════════════════════════════════════════

def _find_anchor(raw_dict: dict) -> float:
    """Busca el valor en T=0; si no existe, usa el valor disponible mas cercano a T=0."""
    anchor = raw_dict.get(0, np.nan)
    if not pd.isna(anchor):
        return anchor
    # Buscar el periodo mas cercano a 0 que tenga dato
    candidates = [(abs(p), p) for p, v in raw_dict.items() if not pd.isna(v)]
    if not candidates:
        return np.nan
    candidates.sort()
    return raw_dict[candidates[0][1]]


def apply_transformation(raw_dict: dict, field_mode: str, baseline_mode: str) -> dict:
    """
    Transforma valores alineados.
    - price          => nivel real (sin tocar, sin normalizar)
    - price_indexed  => siempre base 100 en T=0
    - cumulative_return / pct_change => (v/anchor - 1)*100
    - absolute_change => v - anchor
    """
    anchor = _find_anchor(raw_dict)
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

        elif field_mode == "price_indexed":
            # Siempre indexa a 100 en T=0
            out[p] = (v / anchor * 100) if anchor != 0 else np.nan

        elif field_mode == "price":
            # Nivel real, sin ninguna normalizacion
            if baseline_mode == "base0":
                out[p] = v - anchor
            else:
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
        "T=0" if p == 0 else (f"{fl}+{p}" if p > 0 else f"{fl}{p}")
        for p in periods
    ]


def create_event_chart(
    ticker_id, display_name, events, aligned_data,
    field_mode, baseline_mode, frequency,
    show_avg=True, highlight_event=None
):
    """Crea figura Plotly para un ticker con una linea por evento + promedio."""
    fig = go.Figure()

    y_title = FIELD_MODE_LABELS.get(field_mode, field_mode)
    all_periods = sorted({p for ev_data in aligned_data.values() for p in ev_data})
    period_labels = build_period_labels(all_periods, frequency)
    label_map = dict(zip(all_periods, period_labels))

    has_highlight = highlight_event is not None

    # ── Lineas por evento ─────────────────────────────────────────────────────
    for i, ev in enumerate(events):
        ev_label = ev.get("label", f"Evento {i+1}")
        ev_date  = ev.get("date", "")
        color    = EVENT_COLORS[i % len(EVENT_COLORS)]
        if ev_label not in aligned_data:
            continue
        ev_data = aligned_data[ev_label]
        y_vals  = [ev_data.get(p, np.nan) for p in all_periods]

        # Determinar estilo segun highlight
        is_highlighted = (ev_label == highlight_event)
        if has_highlight and is_highlighted:
            line_w   = 4.5
            marker_s = 7
            opacity  = 1.0
            color    = PALETTE["rojo"]  # El destacado siempre en rojo
        elif has_highlight and not is_highlighted:
            line_w   = 1.5
            marker_s = 3
            opacity  = 0.30
        else:
            line_w   = 2.5
            marker_s = 5
            opacity  = 1.0

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
            line=dict(color=color, width=line_w),
            marker=dict(size=marker_s, color=color),
            opacity=opacity,
            hovertext=hover, hoverinfo="text", connectgaps=True,
        ))

    # ── Linea promedio ────────────────────────────────────────────────────────
    if show_avg and len(aligned_data) > 1:
        mean_y = []
        for p in all_periods:
            vals = [d.get(p, np.nan) for d in aligned_data.values()]
            clean = [v for v in vals if not pd.isna(v)]
            mean_y.append(np.mean(clean) if clean else np.nan)

        avg_opacity = 0.35 if has_highlight else 0.65
        fig.add_trace(go.Scatter(
            x=all_periods, y=mean_y, mode="lines",
            name="Promedio",
            line=dict(color=PALETTE["azul"], width=3, dash="dot"),
            opacity=avg_opacity,
            hovertemplate="<b>Promedio</b><br>Periodo: %{x}<br>Valor: %{y:.4f}<extra></extra>",
        ))

    # ── Linea vertical T=0 ───────────────────────────────────────────────────
    fig.add_vline(
        x=0, line_dash="dash",
        line_color="rgba(0,55,70,0.45)", line_width=1.8,  # azul corporativo
        annotation_text="  T=0", annotation_position="top",
        annotation_font=dict(size=11, color=PALETTE["azul"]),
    )

    # ── Linea horizontal de referencia ────────────────────────────────────────
    if field_mode in ("cumulative_return", "pct_change", "absolute_change"):
        fig.add_hline(y=0, line_color="rgba(0,55,70,0.25)", line_width=1)
    elif field_mode == "price_indexed":
        fig.add_hline(y=100, line_color="rgba(0,55,70,0.25)", line_width=1)

    # ── Ticks ─────────────────────────────────────────────────────────────────
    step = max(1, len(all_periods) // 14)
    tick_vals = [p for i_p, p in enumerate(all_periods) if i_p % step == 0 or p == 0]
    tick_text = [label_map[p] for p in tick_vals]

    # Colores adaptativos: usar negro en tema claro, funciona en ambos
    _tk_color = "#111111"
    _grid_color = "rgba(0,0,0,0.10)"

    fig.update_layout(
        title=dict(text=f"<b>{display_name or ticker_id}</b>",
                   font=dict(size=16, color=PALETTE["azul"], family="Arial"), x=0),
        xaxis=dict(
            title=dict(text="Periodo relativo al evento",
                       font=dict(size=12, color=_tk_color, family="Arial")),
            tickvals=tick_vals, ticktext=tick_text,
            gridcolor=_grid_color, zeroline=False,
            tickangle=0,  # sin rotar
            tickfont=dict(size=11, color=_tk_color, family="Arial"),
        ),
        yaxis=dict(
            title=dict(text=y_title,
                       font=dict(size=12, color=_tk_color, family="Arial")),
            gridcolor=_grid_color,
            tickfont=dict(size=11, color=_tk_color, family="Arial"),
        ),
        plot_bgcolor="rgba(0,0,0,0)",   # transparente para ambos temas
        paper_bgcolor="rgba(0,0,0,0)",
        hovermode="closest",
        legend=dict(orientation="h", yanchor="bottom", y=-0.35,
                    xanchor="center", x=0.5,
                    font=dict(size=11, color=_tk_color, family="Arial"),
                    bordercolor=PALETTE["azul_grisaceo"], borderwidth=1),
        height=490, margin=dict(l=70, r=30, t=55, b=120),
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
    /* ── Paleta corporativa ── */
    .main-hdr{font-size:1.9rem;font-weight:700;color:#003746;line-height:1.2}
    .sub-hdr{color:#005162;font-size:.92rem;margin-top:4px;margin-bottom:1.2rem}
    .sec-title{font-size:.78rem;font-weight:600;text-transform:uppercase;
               letter-spacing:.06em;color:#003746;margin-bottom:6px;margin-top:2px}
    .status-ok{background:#f0f9f4;border-left:3px solid #00AD59;padding:7px 11px;
               border-radius:4px;color:#003746;font-size:.84rem;margin-bottom:6px}
    .status-warn{background:#fff8f0;border-left:3px solid #FF5F00;padding:7px 11px;
                 border-radius:4px;color:#601636;font-size:.84rem;margin-bottom:6px}
    div[data-testid="stExpander"]{border:1px solid #A0D6E2;border-radius:6px}

    /* Linea decorativa bajo el titulo */
    .title-line{height:3px;border:none;margin:0 0 1.2rem 0;
                background:linear-gradient(90deg,#FF1B44 0%,#003746 40%,#009CC6 100%);
                border-radius:2px}

    /* Sidebar header styling */
    section[data-testid="stSidebar"] .stMarkdown h3{color:#003746}
    section[data-testid="stSidebar"] .stDivider{border-color:#A0D6E2}

    /* Primary button con azul */
    .stButton>button[kind="primary"]{background-color:#003746;border-color:#003746}
    .stButton>button[kind="primary"]:hover{background-color:#005162;border-color:#005162}

    /* ── Fix: forzar scroll en contenedor principal ── */
    [data-testid="stAppViewContainer"] > .main {overflow: auto !important}
    .stMainBlockContainer {overflow: visible !important}
    [data-testid="stVerticalBlockBorderWrapper"] {overflow: visible !important}
    section.main > div {overflow: visible !important}
    </style>
    """, unsafe_allow_html=True)

    st.markdown('<div class="main-hdr">Event Study Dashboard</div>', unsafe_allow_html=True)
    st.markdown('<hr class="title-line">', unsafe_allow_html=True)
    st.markdown(
        '<div class="sub-hdr">Compara el comportamiento historico de activos financieros '
        'alrededor de fechas de eventos clave.</div>',
        unsafe_allow_html=True,
    )

    # ── Init session state ────────────────────────────────────────────────────
    if "_ev_uid_counter" not in st.session_state:
        st.session_state._ev_uid_counter = 0

    def _next_ev_uid():
        st.session_state._ev_uid_counter += 1
        return st.session_state._ev_uid_counter

    if "events" not in st.session_state:
        st.session_state.events = []
    # Migrar eventos sin _uid (por si vienen de sesion anterior)
    for ev in st.session_state.events:
        if "_uid" not in ev:
            ev["_uid"] = _next_ev_uid()
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
        # data_source_kind: identifica el origen real para decidir flujo (PDF categorias vs flat).
        #   "bloomberg" → datos descargados desde Bloomberg (presets soportados)
        #   "default"   → archivo .xlsx/.csv del repo (presets soportados)
        #   "upload"    → CSV/Excel subido por el usuario (PDF en modo flat)
        data_source_kind = "bloomberg" if use_bloomberg else "upload"

        if not use_bloomberg:
            # ── Toggle: usar archivo default del repo ──
            default_files = list_default_data_files()
            has_defaults = len(default_files) > 0

            # data_source_kind: "upload" (usuario sube su CSV) | "default" (archivo del repo)
            # Se persiste en session_state para que el flujo de PDF pueda diferenciar.
            data_source_kind = "upload"
            if has_defaults:
                use_default = st.checkbox(
                    "📦 Usar archivo default del repo",
                    value=False,
                    help=(
                        f"Carga un archivo desde la carpeta `data/` del repositorio "
                        f"({len(default_files)} archivo{'s' if len(default_files) > 1 else ''} "
                        f"disponible{'s' if len(default_files) > 1 else ''}). "
                        f"Estos archivos los mantiene actualizados el dueno del proyecto."
                    ),
                    key="use_default_repo_file",
                )
                if use_default:
                    data_source_kind = "default"
            else:
                st.caption("ℹ️ No hay archivos default en `data/`. Sube tu propio archivo.")

            selected_sheet = None

            if data_source_kind == "default":
                # Selector de archivo dentro de data/
                file_names = [p.name for p in default_files]
                default_idx = 0
                if DEFAULT_DATA_FILENAME in file_names:
                    default_idx = file_names.index(DEFAULT_DATA_FILENAME)
                chosen_name = st.selectbox(
                    "📁 Archivo del repo:",
                    file_names,
                    index=default_idx,
                    key="default_file_selector",
                    help=f"El archivo marcado por defecto es `{DEFAULT_DATA_FILENAME}`.",
                )
                chosen_path = DATA_DIR / chosen_name

                # Sheets si es Excel
                sheets_local = get_local_default_sheets(chosen_path)
                if len(sheets_local) >= 1:
                    selected_sheet = st.selectbox(
                        f"📄 Hoja del Excel ({len(sheets_local)} disponible{'s' if len(sheets_local)>1 else ''}):",
                        sheets_local, index=0, key="default_sheet_selector",
                    )

                file_key = f"DEFAULT_{chosen_name}_{selected_sheet or '0'}"
                data_df, err = load_default_data_file(chosen_path, sheet_name=selected_sheet)
                if err:
                    st.error(f"Error al leer `{chosen_name}`: {err}")
                else:
                    st.markdown(
                        f'<div class="status-ok">✅ {len(data_df.columns)} series · '
                        f'{len(data_df):,} fechas<br>'
                        f'{data_df.index[0].date()} → {data_df.index[-1].date()}<br>'
                        f'<span style="font-size:.78rem;opacity:.8">📦 default repo: {chosen_name}</span></div>',
                        unsafe_allow_html=True,
                    )
                    if st.session_state.last_file_key != file_key:
                        st.session_state.last_file_key = file_key
                        st.session_state.tickers = auto_detect_tickers(data_df)
                        st.rerun()
                    with st.expander("Vista previa"):
                        st.dataframe(data_df.tail(5), use_container_width=True)
            else:
                st.caption("Sube tu archivo de precios historicos.")
                uploaded = st.file_uploader("", type=["csv","xlsx","xls"], label_visibility="collapsed")
                if uploaded:
                    # Detectar hojas del Excel (vacio si es CSV)
                    sheets = get_excel_sheet_names(uploaded)
                    if len(sheets) >= 1:
                        selected_sheet = st.selectbox(
                            f"📄 Hoja del Excel ({len(sheets)} disponible{'s' if len(sheets)>1 else ''}):",
                            sheets,
                            index=0, key="sheet_selector",
                        )

                    # Construir clave unica (archivo + hoja) para auto-deteccion
                    sheet_tag = selected_sheet or "0"
                    file_key = f"{uploaded.name}_{uploaded.size}_{sheet_tag}"

                    data_df, err = load_from_file(uploaded, sheet_name=selected_sheet)
                    if err:
                        st.error(f"Error: {err}")
                    else:
                        st.markdown(
                            f'<div class="status-ok">✅ {len(data_df.columns)} series · '
                            f'{len(data_df):,} fechas<br>'
                            f'{data_df.index[0].date()} → {data_df.index[-1].date()}</div>',
                            unsafe_allow_html=True,
                        )
                        # *** AUTO-DETECTAR tickers al subir archivo / cambiar hoja ***
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
        st.caption(f"{len(st.session_state.events)} eventos activos")

        # Calcular rango activo de fechas (para el aviso ⚠️ fuera de rango).
        #   Bloomberg → [Desde, Hasta]
        #   CSV/Excel → [primera fecha, ultima fecha] del data_df cargado
        _range_start = None
        _range_end = None
        if use_bloomberg:
            _range_start = bbg_start
            _range_end   = bbg_end
        elif data_df is not None and len(data_df) > 0:
            _range_start = data_df.index[0].date()
            _range_end   = data_df.index[-1].date()

        def _ev_out_of_range(ev_date_str: str) -> bool:
            if _range_start is None or _range_end is None:
                return False
            try:
                d = pd.Timestamp(ev_date_str).date()
                return d < _range_start or d > _range_end
            except Exception:
                return False

        # Lista compacta con boton de eliminar por evento
        remove_ev = None
        for i, ev in enumerate(st.session_state.events):
            uid = ev["_uid"]
            out_of_range = _ev_out_of_range(ev["date"])
            c_name, c_date, c_warn, c_del = st.columns([5, 3, 0.7, 1])
            with c_name:
                new_label = st.text_input("ev", value=ev["label"], key=f"el_{uid}", label_visibility="collapsed")
            with c_date:
                new_date = st.text_input("dt", value=ev["date"], key=f"ed_{uid}", label_visibility="collapsed")
            with c_warn:
                if out_of_range:
                    tip = (f"Fuera del rango de datos cargados "
                           f"({_range_start} → {_range_end}). "
                           f"Ajusta el rango o el evento.")
                    st.markdown(
                        f"<div style='font-size:1.15rem;padding-top:4px;text-align:center;' "
                        f"title='{tip}'>⚠️</div>",
                        unsafe_allow_html=True,
                    )
                else:
                    st.markdown("&nbsp;", unsafe_allow_html=True)
            with c_del:
                if st.button("❌", key=f"re_{uid}", help=f"Quitar {ev['label']}"):
                    remove_ev = i
            st.session_state.events[i] = {"label": new_label, "date": new_date, "_uid": uid}

        if remove_ev is not None:
            st.session_state.events.pop(remove_ev)
            st.rerun()

        # Botones de accion
        btn_c1, btn_c2 = st.columns(2)
        with btn_c1:
            if st.button("➕ Agregar", use_container_width=True):
                n = len(st.session_state.events) + 1
                st.session_state.events.append({"label": f"Evento {n}", "date": "2024-01-01", "_uid": _next_ev_uid()})
                st.rerun()
        with btn_c2:
            if st.button("🗑️ Limpiar todos", use_container_width=True):
                st.session_state.events = []
                st.rerun()

        with st.expander("📋 Catalogo de referencia", expanded=False):
            st.caption("Selecciona eventos y presiona el boton para agregarlos.")
            selected = st.multiselect(
                "Eventos a agregar:",
                options=[f"{e['label']} ({e['date']})" for e in REFERENCE_EVENTS],
                key="ref_event_selector",
            )
            if selected:
                if st.button("✅ Agregar seleccionados", use_container_width=True):
                    added = 0
                    existing_dates = {e["date"] for e in st.session_state.events}
                    for sel in selected:
                        for ref_ev in REFERENCE_EVENTS:
                            if f"{ref_ev['label']} ({ref_ev['date']})" == sel:
                                if ref_ev["date"] not in existing_dates:
                                    st.session_state.events.append(
                                        {"label": ref_ev["label"], "date": ref_ev["date"], "_uid": _next_ev_uid()}
                                    )
                                    existing_dates.add(ref_ev["date"])
                                    added += 1
                    if added > 0:
                        st.success(f"✅ {added} agregado{'s' if added > 1 else ''}")
                        st.rerun()
                    else:
                        st.info("Ya estan en tu lista.")

        # ── Destacar un evento (justo despues de eventos) ─────────────────────
        event_labels = [ev["label"] for ev in st.session_state.events]
        highlight_options = ["Ninguno"] + event_labels
        # Usar radio dentro de contenedor con scroll para soportar muchos eventos
        st.markdown(
            '<p style="font-size:0.85rem;font-weight:600;margin:0.6rem 0 0.2rem;">Destacar evento:</p>',
            unsafe_allow_html=True,
        )
        container_height = min(35 * len(highlight_options) + 20, 220)
        with st.container(height=container_height):
            highlight_event = st.radio(
                "Destacar evento:",
                highlight_options,
                index=0,
                label_visibility="collapsed",
                help="El evento seleccionado se muestra con linea gruesa y el resto se atenua.",
            )
        if highlight_event == "Ninguno":
            highlight_event = None

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

        # Metodo de agregacion (relevante cuando freq > frecuencia de datos)
        if frequency != "daily":
            agg_method = st.selectbox(
                "Metodo de agregacion",
                AGG_METHODS,
                format_func=lambda x: AGG_METHOD_LABELS.get(x, x),
                help="Si tus datos son diarios y la frecuencia es mensual/semanal:\n"
                     "- Ultimo dato: toma el precio mas reciente del periodo\n"
                     "- Promedio: calcula el promedio de todas las observaciones del periodo",
            )
        else:
            agg_method = "last"

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
                "asset_type": "equity", "field_mode": "price_indexed", "baseline_mode": "base100",
            })
            st.rerun()

        st.divider()

        # ── Opciones de grafica ───────────────────────────────────────────────
        st.markdown('<div class="sec-title">Opciones de grafica</div>', unsafe_allow_html=True)
        show_avg = st.checkbox("Mostrar linea promedio", value=True)

        st.divider()
        run = st.button("🚀 Ejecutar analisis", type="primary", use_container_width=True)

        # ── Generar reporte PDF ────────────────────────────────────────────────
        # Disponible en TODOS los modos (Bloomberg / default / upload).
        # Modo categorico: Bloomberg o archivo default del repo.
        # Modo flat:       CSV/Excel subido por el usuario (sus tickers en una sola seccion).
        _pdf_available = use_bloomberg or (data_df is not None)
        if _pdf_available:
            st.divider()
            st.markdown('<div class="sec-title">📄 Reporte PDF</div>', unsafe_allow_html=True)
            if data_source_kind == "upload":
                st.caption("Modo *flat*: PDF con tus tickers en una sola sección.")
            else:
                st.caption("Genera un PDF ejecutivo con todas las categorias y graficas.")

            generate_pdf = st.button("📄 Generar Reporte PDF", use_container_width=True)
            if generate_pdf:
                if not st.session_state.events:
                    st.error("❌ Agrega al menos un evento antes de generar el reporte.")
                else:
                    from generate_report import build_full_report
                    pdf_data_df = None
                    pdf_tickers_override = None
                    pdf_subtitle = None

                    if use_bloomberg:
                        with st.spinner("📡 Cargando datos y generando reporte..."):
                            force_refresh = st.session_state.get("_bbg_force_refresh", False)
                            pdf_data_df, err = load_master_cache(
                                bbg_start, bbg_end, force_refresh=force_refresh,
                            )
                        if err:
                            st.error(f"❌ {err}")
                            pdf_data_df = None
                        elif pdf_data_df is None or pdf_data_df.empty:
                            st.error("❌ Sin datos Bloomberg.")
                            pdf_data_df = None
                    else:
                        # Modo CSV/Excel (default o upload) → usar data_df cargado.
                        pdf_data_df = data_df
                        if data_source_kind == "upload":
                            # Modo flat: usar los tickers del sidebar
                            pdf_tickers_override = [
                                {
                                    "ticker": tk["ticker"],
                                    "display_name": tk.get("display_name", tk["ticker"]),
                                    "field_mode": tk.get("field_mode", "price_indexed"),
                                    "baseline_mode": tk.get("baseline_mode", "base100"),
                                    "csv_column": tk.get("csv_column", tk["ticker"]),
                                }
                                for tk in st.session_state.tickers
                            ]
                            if not pdf_tickers_override:
                                st.error("❌ Agrega al menos un ticker en la barra lateral.")
                                pdf_data_df = None
                            pdf_subtitle = "Reporte basado en CSV/Excel del usuario"

                    if pdf_data_df is not None and not pdf_data_df.empty:
                        with st.spinner("📄 Generando graficas y compilando PDF..."):
                            pdf_path = build_full_report(
                                events=st.session_state.events,
                                data_df=pdf_data_df,
                                lookback=lookback,
                                lookforward=lookforward,
                                frequency=frequency,
                                agg_method=agg_method,
                                tickers_override=pdf_tickers_override,
                                report_subtitle=pdf_subtitle,
                            )
                        if pdf_path and pdf_path.exists():
                            with open(pdf_path, "rb") as f:
                                st.download_button(
                                    "📥 Descargar Reporte PDF",
                                    f.read(),
                                    file_name=pdf_path.name,
                                    mime="application/pdf",
                                    use_container_width=True,
                                )
                            st.success("✅ Reporte generado")
                        else:
                            st.error("❌ Error al generar el PDF. Verifica que pdflatex este instalado.")

    # ══════════════════════════════════════════════════════════════════════════
    # CONTENIDO PRINCIPAL
    # ══════════════════════════════════════════════════════════════════════════

    # ══════════════════════════════════════════════════════════════════════════
    # TABS PRINCIPALES (Guia + Presets + Manual)
    # ══════════════════════════════════════════════════════════════════════════
    preset_run = False
    preset_tickers = None
    preset_run_key = None  # nombre de la categoria que se ejecuto (para titular el resultado)

    st.markdown("### ⚡ Analisis rapido por categoria")
    st.caption(
        "Empieza por la **📖 Guía de uso** si es tu primera vez. "
        "Luego elige una categoría. Para análisis manual con tus propios tickers, "
        "configúralos en la barra lateral y presiona **🚀 Ejecutar análisis**."
    )

    guide_tab_name   = "📖 Guía de uso"
    preset_tab_names = [f"{v['icon']} {k}" for k, v in TAB_PRESETS.items()]
    tab_names = [guide_tab_name] + preset_tab_names
    ui_tabs = st.tabs(tab_names)

    # ── Tab 0: Guia de uso ────────────────────────────────────────────────────
    with ui_tabs[0]:
        st.markdown("## 📖 Guía de uso del dashboard")
        st.markdown(
            "Este dashboard compara cómo se comportan distintos activos financieros "
            "alrededor de fechas de eventos clave (crisis, decisiones de política monetaria, "
            "elecciones, etc.). Hay **dos formas de ejecutar un análisis** y es importante "
            "no confundirlas."
        )

        st.markdown("### 🅰️  Análisis por categoría (recomendado para empezar)")
        st.markdown(
            "1. Configura **eventos** en la barra lateral (pestaña 📅 Eventos).\n"
            "2. Elige la fuente de datos en la barra lateral (📦 archivo default del repo, "
            "📂 sube tu propio CSV/Excel, o 🔵 Bloomberg si lo tienes).\n"
            "3. Entra a la categoría que te interesa (ej. **🌎 General**, **📊 Sectores S&P**, etc.).\n"
            "4. Presiona el botón **🚀 Ejecutar {Categoría}** que está dentro de la pestaña.\n\n"
            "👉 Estos botones usan los tickers ya pre-configurados de cada categoría — "
            "**no** necesitas agregar tickers manualmente."
        )

        st.markdown("### 🅱️  Análisis manual (tickers personalizados)")
        st.markdown(
            "Si quieres analizar tickers que no están en ninguna categoría preset:\n\n"
            "1. En la barra lateral, sección **📈 Tickers**, presiona **➕ Agregar ticker**.\n"
            "2. Configura sus parámetros (ticker, tipo de activo, transformación, baseline).\n"
            "3. Si subiste un CSV/Excel propio, los tickers se auto-detectan; sólo edítalos si necesitas.\n"
            "4. Presiona **🚀 Ejecutar análisis** que está **al final de la barra lateral**.\n\n"
            "⚠️ **Si presionas \"Ejecutar análisis\" de la barra lateral sin tener tickers configurados, "
            "no aparecerá nada** — el botón está esperando tu lista de tickers. "
            "Si sólo quieres usar las categorías preset, ignora ese botón y usa los botones "
            "🚀 Ejecutar de cada pestaña."
        )

        st.markdown("### 📅 ¿Cómo agregar eventos?")
        st.markdown(
            "En la barra lateral, sección **📅 Eventos**:\n"
            "- Botón **➕ Agregar** crea un evento vacío que puedes editar (label + fecha YYYY-MM-DD).\n"
            "- El expander **📋 Catálogo de referencia** trae 25 eventos históricos pre-cargados "
            "(COVID, Lehman, Brexit, etc.); selecciónalos y dale **✅ Agregar seleccionados**.\n"
            "- Cada categoría también tiene sus propios eventos default cuando ejecutas su preset.\n\n"
            "Si un evento queda **fuera del rango de tus datos** (Bloomberg Desde/Hasta o el rango "
            "del Excel cargado), aparecerá un ⚠️ junto a la fecha. El análisis igual se ejecuta, "
            "pero ese evento puede salir vacío."
        )

        st.markdown("### 📁 Fuentes de datos")
        st.markdown(
            "**📦 Archivo default del repo** — Hay archivos `.xlsx` en la carpeta `data/` del "
            "repositorio que se mantienen actualizados. Activa el checkbox **\"Usar archivo "
            "default del repo\"** y elige el archivo. Funciona con todas las categorías y con PDF.\n\n"
            "**📂 CSV / Excel propio** — Sube tu archivo. La primera columna deben ser fechas "
            "y cada columna siguiente un ticker. Los tickers se auto-detectan al subir el archivo.\n\n"
            "**🔵 Bloomberg** — Requiere Bloomberg Terminal abierto en la misma máquina y `blpapi` "
            "instalado. Define el rango con **Desde** y **Hasta**. Cuando se corre localmente con "
            "Bloomberg, los datos descargados se **guardan automáticamente** en `data/" +
            DEFAULT_DATA_FILENAME + "` para tener una copia."
        )

        st.markdown("### 📄 Generar Reporte PDF")
        st.markdown(
            "El botón **📄 Generar Reporte PDF** aparece en la barra lateral cuando hay datos "
            "cargados (Bloomberg o CSV/Excel).\n\n"
            "- Con archivo **default** del repo o con **Bloomberg**: el PDF cubre todas las "
            "categorías (General, Sectores, Tasas, FX, Commodities, Europa & EM).\n"
            "- Con un **CSV/Excel propio**: el PDF se genera en modo *flat* — sólo con los "
            "tickers detectados en tu archivo, en una sola sección.\n\n"
            "**Dos modos de PDF según el entorno:**\n"
            "- 🖥️ **Local con LaTeX instalado** → PDF *completo* con portada, índice, "
            "descripciones por categoría y gráficas. Requiere `pdflatex` (ver instrucciones "
            "de instalación en el README).\n"
            "- ☁️ **Streamlit Cloud / sin LaTeX** → PDF *simple*: una gráfica por página, "
            "sin texto. Funciona sin instalar nada extra (usa matplotlib).\n\n"
            "La detección es automática — la app intenta usar LaTeX si está disponible y "
            "cae al modo simple si no."
        )

        st.markdown("### 💡 Tips rápidos")
        st.markdown(
            "- **Frecuencia** (Diaria/Semanal/Mensual/Anual) controla los pasos de la ventana.\n"
            "- **Lookback / Lookforward** son los periodos antes/después del evento.\n"
            "- **Destacar evento** atenúa los demás eventos en la gráfica para resaltar uno.\n"
            "- Usa el expander **📚 Catálogo de Tickers Bloomberg** abajo para encontrar IDs."
        )

    # ── Tabs de presets ───────────────────────────────────────────────────────
    for tab_idx, (preset_key, preset_val) in enumerate(TAB_PRESETS.items()):
        with ui_tabs[tab_idx + 1]:  # +1 por el tab de Guia
            st.markdown(f"**{preset_key}** — {preset_val['description']}")

            # Mostrar tickers en grid compacto
            n_cols = 4
            tk_list = preset_val["tickers"]
            for row_start in range(0, len(tk_list), n_cols):
                cols = st.columns(n_cols)
                for col_idx, tk_p in enumerate(tk_list[row_start:row_start + n_cols]):
                    with cols[col_idx]:
                        at_emoji = {"equity":"📈","rate":"📊","fx":"💱","commodity":"🛢️","other":"📋"}.get(tk_p.get("asset_type",""),"📈")
                        st.markdown(f"{at_emoji} **{tk_p['display_name']}**")
                        st.caption(tk_p["ticker"])

            # Mostrar eventos actuales de la sidebar
            n_ev = len(st.session_state.events)
            if n_ev > 0:
                ev_str = " · ".join(e["label"] for e in st.session_state.events)
                st.markdown(f"<div style='font-size:0.82rem;color:#5A6670;margin:6px 0 10px;'>"
                            f"📅 {ev_str}</div>", unsafe_allow_html=True)
            else:
                st.warning("⚠️ Agrega al menos un evento en la barra lateral antes de ejecutar.")

            # En modo CSV/Excel, avisar si pocas columnas matchean los tickers del preset
            if not use_bloomberg and data_df is not None:
                preset_ticker_ids = {tk_p["ticker"] for tk_p in tk_list}
                avail_cnt = sum(1 for t in preset_ticker_ids if t in data_df.columns)
                if avail_cnt == 0:
                    st.error(
                        f"❌ Ninguno de los {len(preset_ticker_ids)} tickers de **{preset_key}** "
                        f"está en tu archivo. Verifica los headers de tu Excel/CSV o cambia de fuente."
                    )
                elif avail_cnt < len(preset_ticker_ids):
                    st.info(
                        f"ℹ️ {avail_cnt} de {len(preset_ticker_ids)} tickers disponibles en tu archivo. "
                        f"Los faltantes se omitirán."
                    )

            if st.button(f"🚀 Ejecutar {preset_key}", key=f"preset_run_{preset_key}",
                         type="primary", use_container_width=True):
                preset_run = True
                preset_run_key = preset_key
                # Para modo CSV/Excel: setear csv_column = ticker para que el pipeline existente
                # busque la columna por nombre del ticker.
                preset_tickers = [
                    {**dict(tk_p), "csv_column": tk_p["ticker"]}
                    for tk_p in preset_val["tickers"]
                ]

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
    if not run and not preset_run:
        if not use_bloomberg:
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
                st.caption(f"...y {len(REFERENCE_EVENTS)-12} mas. Usa el 📋 Catalogo en la barra lateral.")
        return

    # ══════════════════════════════════════════════════════════════════════════
    # VALIDACION
    # ══════════════════════════════════════════════════════════════════════════

    # Eventos siempre vienen de la sidebar; tickers del preset o de la sidebar
    events = st.session_state.events
    if preset_run and preset_tickers is not None:
        tickers = preset_tickers
    else:
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

    # Si highlight_event no esta en los eventos activos, ignorar
    active_labels = {ev["label"] for ev in valid_events}
    if highlight_event is not None and highlight_event not in active_labels:
        highlight_event = None

    # ══════════════════════════════════════════════════════════════════════════
    # CARGA BLOOMBERG (con cache maestro)
    # ══════════════════════════════════════════════════════════════════════════
    if use_bloomberg:
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

        if preset_run:
            # ── Modo preset: usar cache maestro (descarga UNA vez todos los tickers) ──
            force_refresh = st.session_state.get("_bbg_force_refresh", False)
            if force_refresh:
                st.session_state._bbg_force_refresh = False

            with st.spinner("🔄 Cargando datos Bloomberg (cache maestro)..."):
                master_df, err = load_master_cache(actual_s, actual_e, force_refresh=force_refresh)

            if err:
                st.error(f"❌ {err}")
                return
            if master_df is None or master_df.empty:
                st.error("❌ Bloomberg no devolvio datos.")
                return

            # Filtrar solo las columnas del preset actual
            needed_tickers = [tk["ticker"] for tk in tickers]
            available = [t for t in needed_tickers if t in master_df.columns]
            missing_in_cache = [t for t in needed_tickers if t not in master_df.columns]

            if missing_in_cache:
                for mt in missing_in_cache:
                    st.warning(f"⚠️ Sin datos para **{mt}** en cache.")

            data_df = master_df[available] if available else pd.DataFrame()

            cache_file = _cache_path_for_date(date.today())
            cache_age = ""
            if cache_file.exists():
                import time as _time
                mod_ts = os.path.getmtime(cache_file)
                mod_dt = pd.Timestamp.fromtimestamp(mod_ts)
                cache_age = f" · cache {mod_dt.strftime('%H:%M')}"

            st.success(
                f"✅ Bloomberg: {len(data_df.columns)} series · "
                f"{len(data_df):,} obs (de {len(master_df.columns)} en cache{cache_age})"
            )

            # ── Auto-save local: sobrescribir el Excel default del repo ──
            # Solo cuando blpapi esta disponible (= corriendo en la maquina del usuario,
            # no en Streamlit Cloud que tiene filesystem read-only).
            # NOTA: NO renombramos columnas — los headers quedan como tickers Bloomberg
            # (SPX Index, GT10 Govt, etc.) para que el analisis por categoria pueda
            # hacer lookup por ticker en otros usuarios.
            if BLOOMBERG_AVAILABLE:
                _ok, _info = save_data_to_default_excel(
                    master_df,
                    target_path=DEFAULT_DATA_PATH,
                )
                if _ok:
                    st.caption(f"💾 Copia guardada en `data/{DEFAULT_DATA_FILENAME}`")
                else:
                    st.caption(f"⚠️ No se pudo guardar copia local: {_info}")

            # Boton para forzar re-descarga
            if st.button("🔄 Actualizar datos Bloomberg", help="Re-descarga todos los tickers desde Bloomberg"):
                st.session_state._bbg_force_refresh = True
                st.rerun()

            # ── Descarga Excel con TODOS los tickers del cache ──
            all_names = _get_all_preset_names()
            rename_map = {t: all_names.get(t, t) for t in master_df.columns if t in all_names}
            excel_df = master_df.rename(columns=rename_map)
            excel_df.index.name = "Date"

            xlsx_buf = BytesIO()
            with pd.ExcelWriter(xlsx_buf, engine="openpyxl") as writer:
                excel_df.to_excel(writer, sheet_name="Bloomberg Data")
            xlsx_buf.seek(0)
            today_str = date.today().strftime("%Y%m%d")
            st.download_button(
                f"📥 Descargar Excel completo ({len(excel_df.columns)} tickers)",
                xlsx_buf.getvalue(),
                f"bloomberg_all_tickers_{today_str}.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        else:
            # ── Modo manual: descargar solo los tickers configurados ──
            fields_map = {tk["ticker"]: tk.get("bloomberg_field","PX_LAST") for tk in tickers}
            with st.spinner("🔄 Descargando de Bloomberg..."):
                data_df, err = load_from_bloomberg(fields_map, actual_s, actual_e)
            if err:
                st.error(f"❌ {err}")
                return
            if data_df is None or data_df.empty:
                st.error("❌ Bloomberg no devolvio datos.")
                return
            st.success(f"✅ Bloomberg: {len(data_df.columns)} series · {len(data_df):,} obs")

            # ── Auto-save local: agregar/actualizar estos tickers en el Excel default ──
            # Si ya existe el archivo, hacemos merge para no perder otros tickers.
            # NOTA: NO renombramos columnas — headers quedan como tickers Bloomberg.
            if BLOOMBERG_AVAILABLE:
                merged_df = data_df.copy()
                if DEFAULT_DATA_PATH.exists():
                    try:
                        prev_df, _ = load_default_data_file(DEFAULT_DATA_PATH)
                        if prev_df is not None and not prev_df.empty:
                            new_cols = [c for c in prev_df.columns if c not in merged_df.columns]
                            if new_cols:
                                merged_df = pd.concat([merged_df, prev_df[new_cols]], axis=1)
                    except Exception:
                        pass
                _ok, _info = save_data_to_default_excel(
                    merged_df,
                    target_path=DEFAULT_DATA_PATH,
                )
                if _ok:
                    st.caption(f"💾 Copia guardada en `data/{DEFAULT_DATA_FILENAME}`")

    else:
        # ── Modo CSV/Excel ──
        if data_df is None:
            st.error("❌ Activa **Usar archivo default del repo** o sube un archivo CSV/Excel primero.")
            return

        if preset_run:
            # Filtrar el data_df a las columnas que existen del preset.
            needed = [tk["ticker"] for tk in tickers]
            avail = [t for t in needed if t in data_df.columns]
            missing = [t for t in needed if t not in data_df.columns]

            if not avail:
                st.error(
                    f"❌ Ninguno de los {len(needed)} tickers de **{preset_run_key}** "
                    f"está en tu archivo. Verifica los headers o cambia de fuente."
                )
                return
            if missing:
                _show_missing = ', '.join(missing[:8])
                if len(missing) > 8:
                    _show_missing += f" ...y {len(missing)-8} más"
                st.warning(
                    f"⚠️ {len(missing)} de {len(needed)} tickers sin datos en tu archivo: "
                    f"{_show_missing}"
                )

            # Filtrar tickers a los disponibles para que el pipeline no falle
            tickers = [t for t in tickers if t["ticker"] in avail]
            data_df = data_df[avail]
            st.success(
                f"✅ {preset_run_key}: {len(data_df.columns)} series · {len(data_df):,} obs "
                f"(de archivo cargado)"
            )

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
                raw = extract_aligned_values(
                    series, ev["date"], lookback, lookforward, frequency,
                    agg_method=agg_method,
                )
                # Solo omitir si NO hay ningun dato en toda la ventana
                has_any = any(not pd.isna(v) for v in raw.values())
                if not has_any:
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
            aligned_data, field_mode, baseline_mode, frequency,
            show_avg=show_avg, highlight_event=highlight_event,
        )
        st.plotly_chart(fig, use_container_width=True, key=f"chart_{tk_idx}")

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


# Punto de entrada
if __name__ == "__main__":
    main()
