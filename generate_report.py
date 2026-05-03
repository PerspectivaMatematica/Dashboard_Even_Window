#!/usr/bin/env python3
"""
generate_report.py  —  Reporte PDF ejecutivo del Event Study Dashboard
=====================================================================
Autosuficiente: no requiere streamlit.  Copia las funciones y constantes
necesarias de app.py para poder correr de forma independiente.

Uso:
    python generate_report.py datos.xlsx            # lee Excel exportado
    python generate_report.py                       # usa cache Bloomberg
"""

import os
import sys
import shutil
import subprocess
import tempfile
from datetime import date, timedelta
from pathlib import Path

import numpy as np
import pandas as pd
from dateutil.relativedelta import relativedelta


# ══════════════════════════════════════════════════════════════════════════════
# CONSTANTES (copiadas de app.py para ser autosuficiente)
# ══════════════════════════════════════════════════════════════════════════════

FIELD_MODE_LABELS = {
    "price":             "Precio",
    "price_indexed":     "Precio Indexado (Base 100)",
    "cumulative_return": "Retorno Acumulado (%)",
    "pct_change":        "Cambio Porcentual (%)",
    "absolute_change":   "Cambio Absoluto",
}

FREQ_LABEL = {"daily": "D", "weekly": "S", "monthly": "M", "annual": "A"}

PALETTE = {
    "rojo": "#FF1B44", "azul": "#003746", "azul_digital": "#009CC6",
    "naranja_oscuro": "#FF5F00", "verde_oscuro": "#00AD59", "verde": "#2DDC8E",
    "violeta_oscuro": "#B18DFB", "naranja": "#FA8D5A", "granate": "#601636",
    "granate_claro": "#B77493", "azul_digital_claro": "#005162",
    "azul_grisaceo": "#A0D6E2",
}

EVENT_COLORS = [
    "#003746", "#FF1B44", "#00AD59", "#FF5F00", "#B18DFB", "#009CC6",
    "#601636", "#FA8D5A", "#2DDC8E", "#B77493", "#005162", "#A0D6E2",
]

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
        "tickers": [
            {"ticker": "SPX Index",      "display_name": "S&P 500",       "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "NDX Index",      "display_name": "Nasdaq 100",    "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "VIX Index",      "display_name": "VIX",           "field_mode": "price",         "baseline_mode": "none"},
            {"ticker": "SX5E Index",     "display_name": "Euro Stoxx 50", "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "MEXBOL Index",   "display_name": "IPC Mexico",    "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "GT2 Govt",       "display_name": "US 2Y",         "field_mode": "absolute_change","baseline_mode": "base0"},
            {"ticker": "GT10 Govt",      "display_name": "US 10Y",        "field_mode": "absolute_change","baseline_mode": "base0"},
            {"ticker": "GTMXN2Y Govt",   "display_name": "MX 2Y",         "field_mode": "absolute_change","baseline_mode": "base0"},
            {"ticker": "GTMXN10Y Govt",  "display_name": "Mbono 10Y",     "field_mode": "absolute_change","baseline_mode": "base0"},
            {"ticker": "GDBR10 Index",   "display_name": "Bund 10Y",      "field_mode": "absolute_change","baseline_mode": "base0"},
            {"ticker": "GTII10 Govt",    "display_name": "US TIPS 10Y",   "field_mode": "absolute_change","baseline_mode": "base0"},
            {"ticker": "USGGBE10 Index", "display_name": "US BE 10Y",     "field_mode": "absolute_change","baseline_mode": "base0"},
            {"ticker": "GTMXNII10Y Govt","display_name": "MX Real 10Y",   "field_mode": "absolute_change","baseline_mode": "base0"},
            {"ticker": "USDMXN Curncy",  "display_name": "USD/MXN",       "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "DXY Index",      "display_name": "DXY",           "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "EURUSD Curncy",  "display_name": "EUR/USD",       "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "XAU Curncy",     "display_name": "Oro",           "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "XAG Curncy",     "display_name": "Plata",         "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "CL1 Comdty",     "display_name": "WTI Crudo",     "field_mode": "price_indexed", "baseline_mode": "base100"},
        ],
    },
    "Sectores S&P": {
        "tickers": [
            {"ticker": "S5INFT Index", "display_name": "Tecnologia",          "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "S5FINL Index", "display_name": "Financieros",         "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "S5HLTH Index", "display_name": "Salud",               "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "S5COND Index", "display_name": "Consumo Discrecional","field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "S5CONS Index", "display_name": "Consumo Basico",      "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "S5ENRS Index", "display_name": "Energia",             "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "S5INDU Index", "display_name": "Industriales",        "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "S5MATR Index", "display_name": "Materiales",          "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "S5TELS Index", "display_name": "Comunicaciones",      "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "S5UTIL Index", "display_name": "Utilities",           "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "S5RLST Index", "display_name": "Real Estate",         "field_mode": "price_indexed", "baseline_mode": "base100"},
        ],
    },
    "Tasas US & MX": {
        "tickers": [
            {"ticker": "GT2 Govt",        "display_name": "US Nom 2Y",    "field_mode": "absolute_change", "baseline_mode": "base0"},
            {"ticker": "GT5 Govt",        "display_name": "US Nom 5Y",    "field_mode": "absolute_change", "baseline_mode": "base0"},
            {"ticker": "GT10 Govt",       "display_name": "US Nom 10Y",   "field_mode": "absolute_change", "baseline_mode": "base0"},
            {"ticker": "GT20 Govt",       "display_name": "US Nom 20Y",   "field_mode": "absolute_change", "baseline_mode": "base0"},
            {"ticker": "GT30 Govt",       "display_name": "US Nom 30Y",   "field_mode": "absolute_change", "baseline_mode": "base0"},
            {"ticker": "USGGT02Y Index",  "display_name": "US TIPS 2Y",   "field_mode": "absolute_change", "baseline_mode": "base0"},
            {"ticker": "GTII5 Govt",      "display_name": "US TIPS 5Y",   "field_mode": "absolute_change", "baseline_mode": "base0"},
            {"ticker": "GTII10 Govt",     "display_name": "US TIPS 10Y",  "field_mode": "absolute_change", "baseline_mode": "base0"},
            {"ticker": "GTII20 Govt",     "display_name": "US TIPS 20Y",  "field_mode": "absolute_change", "baseline_mode": "base0"},
            {"ticker": "GTII30 Govt",     "display_name": "US TIPS 30Y",  "field_mode": "absolute_change", "baseline_mode": "base0"},
            {"ticker": "USGGBE02 Index",  "display_name": "US BE 2Y",     "field_mode": "absolute_change", "baseline_mode": "base0"},
            {"ticker": "USGGBE05 Index",  "display_name": "US BE 5Y",     "field_mode": "absolute_change", "baseline_mode": "base0"},
            {"ticker": "USGGBE10 Index",  "display_name": "US BE 10Y",    "field_mode": "absolute_change", "baseline_mode": "base0"},
            {"ticker": "MXIBTIIE Index",  "display_name": "TIIE 28d",     "field_mode": "absolute_change", "baseline_mode": "base0"},
            {"ticker": "GTMXN2Y Govt",    "display_name": "MX Nom 2Y",    "field_mode": "absolute_change", "baseline_mode": "base0"},
            {"ticker": "GTMXN5Y Govt",    "display_name": "MX Nom 5Y",    "field_mode": "absolute_change", "baseline_mode": "base0"},
            {"ticker": "GTMXN10Y Govt",   "display_name": "MX Nom 10Y",   "field_mode": "absolute_change", "baseline_mode": "base0"},
            {"ticker": "GTMXN20Y Govt",   "display_name": "MX Nom 20Y",   "field_mode": "absolute_change", "baseline_mode": "base0"},
            {"ticker": "GTMXN30Y Govt",   "display_name": "MX Nom 30Y",   "field_mode": "absolute_change", "baseline_mode": "base0"},
            {"ticker": "GTMXNII5Y Govt",  "display_name": "MX Real 5Y",   "field_mode": "absolute_change", "baseline_mode": "base0"},
            {"ticker": "GTMXNII10Y Govt", "display_name": "MX Real 10Y",  "field_mode": "absolute_change", "baseline_mode": "base0"},
            {"ticker": "GTMXNII20Y Govt", "display_name": "MX Real 20Y",  "field_mode": "absolute_change", "baseline_mode": "base0"},
            {"ticker": "GTMXNII30Y Govt", "display_name": "MX Real 30Y",  "field_mode": "absolute_change", "baseline_mode": "base0"},
        ],
    },
    "FX": {
        "tickers": [
            {"ticker": "USDMXN Curncy", "display_name": "USD/MXN", "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "EURUSD Curncy", "display_name": "EUR/USD", "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "USDJPY Curncy", "display_name": "USD/JPY", "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "GBPUSD Curncy", "display_name": "GBP/USD", "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "DXY Index",     "display_name": "DXY",     "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "USDBRL Curncy", "display_name": "USD/BRL", "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "USDCNH Curncy", "display_name": "USD/CNH", "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "USDZAR Curncy", "display_name": "USD/ZAR", "field_mode": "price_indexed", "baseline_mode": "base100"},
        ],
    },
    "Commodities": {
        "tickers": [
            {"ticker": "CL1 Comdty",    "display_name": "WTI Crudo",    "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "CO1 Comdty",    "display_name": "Brent",        "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "NG1 Comdty",    "display_name": "Gas Natural",  "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "XAU Curncy",    "display_name": "Oro",          "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "XAG Curncy",    "display_name": "Plata",        "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "HG1 Comdty",    "display_name": "Cobre",        "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "LA1 Comdty",    "display_name": "Aluminio",     "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "LIT US Equity", "display_name": "Litio (ETF)",  "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "C 1 Comdty",    "display_name": "Maiz",         "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "W 1 Comdty",    "display_name": "Trigo",        "field_mode": "price_indexed", "baseline_mode": "base100"},
        ],
    },
    "Europa & EM": {
        "tickers": [
            {"ticker": "SX5E Index",   "display_name": "Euro Stoxx 50", "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "DAX Index",    "display_name": "DAX",           "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "UKX Index",    "display_name": "FTSE 100",      "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "GDBR10 Index", "display_name": "Bund 10Y",      "field_mode": "absolute_change","baseline_mode": "base0"},
            {"ticker": "MEXBOL Index", "display_name": "IPC Mexico",    "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "IBOV Index",   "display_name": "Bovespa",       "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "SHCOMP Index", "display_name": "Shanghai Comp", "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "MXEF Index",   "display_name": "MSCI EM",       "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "NIFTY Index",  "display_name": "Nifty 50",      "field_mode": "price_indexed", "baseline_mode": "base100"},
            {"ticker": "HSI Index",    "display_name": "Hang Seng",     "field_mode": "price_indexed", "baseline_mode": "base100"},
        ],
    },
}


# ══════════════════════════════════════════════════════════════════════════════
# FUNCIONES DE PROCESAMIENTO (copiadas de app.py)
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
    if len(series_index) == 0:
        return None
    if target_ts > series_index[-1]:
        return None
    if target_ts < series_index[0]:
        return None
    prior = series_index[series_index <= target_ts]
    if len(prior) == 0:
        return None
    return prior[-1]


def extract_aligned_values(series, event_date, lookback, lookforward, frequency, agg_method="last"):
    target_dates = build_target_dates(event_date, lookback, lookforward, frequency)
    result = {}
    for period, target_ts in target_dates:
        mapped = get_prior_observation(target_ts, series.index)
        result[period] = series.loc[mapped] if mapped is not None else np.nan
    return result


def _find_anchor(raw_dict):
    anchor = raw_dict.get(0, np.nan)
    if not pd.isna(anchor):
        return anchor
    candidates = [(abs(p), p) for p, v in raw_dict.items() if not pd.isna(v)]
    if not candidates:
        return np.nan
    candidates.sort()
    return raw_dict[candidates[0][1]]


def apply_transformation(raw_dict, field_mode, baseline_mode):
    anchor = _find_anchor(raw_dict)
    if pd.isna(anchor):
        return {p: np.nan for p in raw_dict}
    out = {}
    for p, v in raw_dict.items():
        if pd.isna(v):
            out[p] = np.nan
            continue
        if field_mode in ("cumulative_return", "pct_change"):
            out[p] = (v / anchor - 1) * 100 if anchor != 0 else np.nan
        elif field_mode == "absolute_change":
            out[p] = v - anchor
        elif field_mode == "price_indexed":
            out[p] = (v / anchor * 100) if anchor != 0 else np.nan
        elif field_mode == "price":
            out[p] = v - anchor if baseline_mode == "base0" else v
        else:
            out[p] = v
    return out


def build_period_labels(periods, freq):
    fl = FREQ_LABEL.get(freq, "T")
    return [
        "T=0" if p == 0 else ("{}+{}".format(fl, p) if p > 0 else "{}{}".format(fl, p))
        for p in periods
    ]


# ══════════════════════════════════════════════════════════════════════════════
# MATPLOTLIB CHART
# ══════════════════════════════════════════════════════════════════════════════

def _create_mpl_chart(ticker_id, display_name, events, aligned_data, field_mode, frequency):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    fig, ax = plt.subplots(figsize=(10, 4.2))

    y_label = FIELD_MODE_LABELS.get(field_mode, field_mode)
    all_periods = sorted({p for d in aligned_data.values() for p in d})
    period_labels = build_period_labels(all_periods, frequency)

    for i, ev in enumerate(events):
        ev_label = ev.get("label", "Evento {}".format(i + 1))
        ev_date = ev.get("date", "")
        color = EVENT_COLORS[i % len(EVENT_COLORS)]
        if ev_label not in aligned_data:
            continue
        ev_data = aligned_data[ev_label]
        y_vals = [ev_data.get(p, np.nan) for p in all_periods]
        ax.plot(all_periods, y_vals, color=color, linewidth=1.8, marker="o",
                markersize=3, label="{} ({})".format(ev_label, ev_date), zorder=3)

    if len(aligned_data) > 1:
        mean_y = []
        for p in all_periods:
            vals = [d.get(p, np.nan) for d in aligned_data.values()]
            clean = [v for v in vals if not np.isnan(v)]
            mean_y.append(np.mean(clean) if clean else np.nan)
        ax.plot(all_periods, mean_y, color=PALETTE["azul"], linewidth=2.5,
                linestyle=":", alpha=0.6, label="Promedio", zorder=2)

    ax.axvline(x=0, color=PALETTE["azul"], linewidth=1.2, linestyle="--", alpha=0.5)
    if field_mode in ("cumulative_return", "pct_change", "absolute_change"):
        ax.axhline(y=0, color="#999999", linewidth=0.8, alpha=0.4)

    ax.set_title(display_name, fontsize=13, fontweight="bold", color="#111111", pad=10)
    ax.set_ylabel(y_label, fontsize=9, color="#333333")
    ax.set_xlabel("Periodo relativo al evento", fontsize=9, color="#333333")

    step = max(1, len(all_periods) // 15)
    tick_idx = list(range(0, len(all_periods), step))
    ax.set_xticks([all_periods[i] for i in tick_idx])
    ax.set_xticklabels([period_labels[i] for i in tick_idx], fontsize=7.5, color="#333333")
    ax.tick_params(axis="y", labelsize=8, colors="#333333")
    ax.grid(axis="y", alpha=0.2, color="#000000")
    ax.grid(axis="x", alpha=0.1, color="#000000")
    ax.legend(fontsize=6.5, loc="upper center", bbox_to_anchor=(0.5, -0.18),
              ncol=min(4, len(aligned_data) + 1), frameon=False)
    fig.patch.set_alpha(0)
    ax.set_facecolor("white")
    plt.tight_layout()
    return fig


# ══════════════════════════════════════════════════════════════════════════════
# DETECCION DE pdflatex (LaTeX vs PDF simple)
# ══════════════════════════════════════════════════════════════════════════════

def _has_pdflatex() -> bool:
    """True si pdflatex esta disponible en el PATH del sistema."""
    return shutil.which("pdflatex") is not None


# ══════════════════════════════════════════════════════════════════════════════
# DESCRIPCIONES Y LATEX
# ══════════════════════════════════════════════════════════════════════════════

CAT_DESCRIPTIONS = {
    "General": "Panorama macro con los principales indices bursatiles, tasas de referencia, divisas y materias primas. Permite comparar la reaccion global de los mercados ante cada evento de estres.",
    "Sectores S&P": "Los 11 sectores GICS del S\\&P 500. Muestra como se distribuye el impacto de cada evento entre sectores defensivos y ciclicos.",
    "Tasas US & MX": "Curvas nominales y reales (TIPS / UDIBONOS) de Estados Unidos y Mexico. Permite observar movimientos de flight-to-quality y expectativas de politica monetaria.",
    "FX": "Principales pares de divisas y monedas emergentes. Refleja flujos de capital, aversion al riesgo y diferenciales de tasas.",
    "Commodities": "Energeticos, metales preciosos, industriales y agricolas. Captura disrupciones de oferta, demanda y coberturas de riesgo.",
    "Europa & EM": "Indices europeos y de mercados emergentes. Muestra el contagio y la diversificacion entre bloques economicos.",
}

LATEX_PREAMBLE = r"""
\documentclass[11pt, a4paper]{article}
\usepackage[utf8]{inputenc}
\usepackage[T1]{fontenc}
\usepackage[spanish]{babel}
\usepackage{geometry}
\geometry{margin=2cm, top=2.5cm, bottom=2.5cm}
\usepackage{graphicx}
\usepackage{xcolor}
\usepackage{fancyhdr}
\usepackage{titlesec}
\usepackage{parskip}
\usepackage{hyperref}

\definecolor{azulcorp}{HTML}{003746}
\definecolor{rojocorp}{HTML}{FF1B44}
\definecolor{gris}{HTML}{5A6670}

\pagestyle{fancy}
\fancyhf{}
\fancyhead[L]{\small\color{gris}Event Study Dashboard}
\fancyhead[R]{\small\color{gris}\today}
\fancyfoot[C]{\small\color{gris}\thepage}
\renewcommand{\headrulewidth}{0.4pt}

\titleformat{\section}{\Large\bfseries\color{azulcorp}}{}{0em}{}[\vspace{-0.5em}\color{azulcorp}\rule{\textwidth}{0.8pt}]
\titleformat{\subsection}{\large\bfseries\color{azulcorp}}{}{0em}{}

\hypersetup{colorlinks=true, linkcolor=azulcorp, urlcolor=azulcorp}
"""


# ══════════════════════════════════════════════════════════════════════════════
# BUILD FUNCTION — VARIANTE SIMPLE (sin LaTeX, usa matplotlib PdfPages)
# ══════════════════════════════════════════════════════════════════════════════

def _build_simple_pdf_report(events, data_df, lookback=30, lookforward=30,
                              frequency="daily", agg_method="last",
                              output_dir=None, tickers_override=None):
    """Genera un PDF multi-pagina (una grafica por pagina) usando matplotlib
    PdfPages. NO requiere pdflatex. Pensado para Streamlit Cloud o ambientes
    sin LaTeX. Sin portada, sin texto de referencia — solo las graficas.
    """
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_pdf import PdfPages

    if output_dir is None:
        output_dir = Path(__file__).resolve().parent
    output_dir = Path(output_dir)
    today_str = date.today().strftime("%Y%m%d")

    # Validar eventos
    valid_events = []
    for ev in events:
        try:
            pd.Timestamp(ev["date"])
            valid_events.append(ev)
        except Exception:
            pass
    if not valid_events:
        print("ERROR: No hay eventos validos.")
        return None

    # Determinar secciones
    if tickers_override is not None:
        sections_iter = [("Tus Series", {"tickers": list(tickers_override)})]
    else:
        sections_iter = list(TAB_PRESETS.items())

    # Output path
    pdf_final = output_dir / "Event_Study_Report_{}.pdf".format(today_str)

    total_charts = 0
    try:
        with PdfPages(str(pdf_final)) as pdf:
            for cat_name, cat_val in sections_iter:
                for tk in cat_val["tickers"]:
                    ticker_id = tk["ticker"]
                    display_name = tk.get("display_name", ticker_id)
                    field_mode = tk.get("field_mode", "price_indexed")
                    baseline_mode = tk.get("baseline_mode", "base100")

                    lookup_col = tk.get("csv_column") or ticker_id
                    if lookup_col not in data_df.columns:
                        if ticker_id in data_df.columns:
                            lookup_col = ticker_id
                        else:
                            continue
                    series = data_df[lookup_col].dropna()
                    if len(series) == 0:
                        continue

                    aligned_data = {}
                    for ev in valid_events:
                        try:
                            raw = extract_aligned_values(
                                series, ev["date"], lookback, lookforward,
                                frequency, agg_method,
                            )
                            if any(not pd.isna(v) for v in raw.values()):
                                aligned_data[ev["label"]] = apply_transformation(
                                    raw, field_mode, baseline_mode,
                                )
                        except Exception:
                            pass

                    if not aligned_data:
                        continue

                    fig = _create_mpl_chart(
                        ticker_id, display_name, valid_events,
                        aligned_data, field_mode, frequency,
                    )
                    pdf.savefig(fig, dpi=150, bbox_inches="tight",
                                facecolor="white")
                    plt.close(fig)
                    total_charts += 1

        if total_charts == 0:
            print("ERROR: No se genero ninguna grafica.")
            try:
                pdf_final.unlink()
            except Exception:
                pass
            return None

        print("PDF (modo simple): {}".format(pdf_final))
        return pdf_final

    except Exception as e:
        print("ERROR generando PDF simple: {}".format(e))
        try:
            if pdf_final.exists():
                pdf_final.unlink()
        except Exception:
            pass
        return None


# ══════════════════════════════════════════════════════════════════════════════
# BUILD FUNCTION
# ══════════════════════════════════════════════════════════════════════════════

def build_full_report(events, data_df, lookback=30, lookforward=30,
                      frequency="daily", agg_method="last", output_dir=None,
                      tickers_override=None, report_subtitle=None):
    """Genera el PDF completo. Retorna Path al PDF o None.

    Auto-detecta si pdflatex esta disponible:
    - SI hay pdflatex (corriendo localmente con LaTeX instalado) → genera el
      PDF "completo" con portada, secciones, descripciones y graficas.
    - NO hay pdflatex (Streamlit Cloud sin texlive) → cae al modo simple:
      una grafica por pagina, sin texto, usando matplotlib PdfPages.

    Args:
        events: lista de {label, date}
        data_df: DataFrame de precios (index = fechas, columnas = tickers)
        tickers_override: si se pasa, genera reporte FLAT (una sola seccion
            "Tus Series") en lugar de iterar TAB_PRESETS. Cada elemento debe ser
            {ticker, display_name, field_mode, baseline_mode}. Util para CSVs
            del usuario que no usan tickers Bloomberg.
        report_subtitle: subtitulo opcional para la portada (solo modo LaTeX).
    """
    # ── Dispatch: si no hay pdflatex, usar el generador simple ──
    if not _has_pdflatex():
        print("pdflatex no encontrado → generando PDF en modo simple (matplotlib).")
        return _build_simple_pdf_report(
            events, data_df, lookback, lookforward, frequency, agg_method,
            output_dir, tickers_override,
        )

    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    if output_dir is None:
        output_dir = Path(__file__).resolve().parent
    output_dir = Path(output_dir)
    today_str = date.today().strftime("%Y%m%d")

    tmp_dir = Path(tempfile.mkdtemp(prefix="eventstudy_"))
    img_dir = tmp_dir / "img"
    img_dir.mkdir()

    valid_events = []
    for ev in events:
        try:
            pd.Timestamp(ev["date"])
            valid_events.append(ev)
        except Exception:
            pass
    if not valid_events:
        print("ERROR: No hay eventos validos.")
        return None

    # ── Construir las "secciones" a iterar ────────────────────────────────────
    # Modo categorico (default) → usa TAB_PRESETS.
    # Modo flat (tickers_override) → una sola "seccion" con los tickers del usuario.
    if tickers_override is not None:
        sections_iter = [(
            "Tus Series",
            {"tickers": list(tickers_override)},
        )]
        flat_mode = True
    else:
        sections_iter = list(TAB_PRESETS.items())
        flat_mode = False

    # ── Graficas ──────────────────────────────────────────────────────────────
    latex_sections = []
    total_tickers = 0
    total_charts = 0

    for cat_name, cat_val in sections_iter:
        cat_tickers = cat_val["tickers"]
        desc = CAT_DESCRIPTIONS.get(cat_name, "") if not flat_mode else (
            "Series detectadas en el archivo cargado por el usuario."
        )

        section_tex = "\\section{%s}\n" % cat_name.replace("&", "\\&")
        if desc:
            section_tex += "\\textcolor{gris}{%s}\n\n" % desc
        chart_count = 0

        for tk in cat_tickers:
            ticker_id = tk["ticker"]
            display_name = tk.get("display_name", ticker_id)
            field_mode = tk.get("field_mode", "price_indexed")
            baseline_mode = tk.get("baseline_mode", "base100")

            # En modo flat el "lookup column" puede ser csv_column si esta seteado
            lookup_col = tk.get("csv_column") or ticker_id
            if lookup_col not in data_df.columns:
                if ticker_id in data_df.columns:
                    lookup_col = ticker_id
                else:
                    continue
            series = data_df[lookup_col].dropna()
            if len(series) == 0:
                continue

            aligned_data = {}
            for ev in valid_events:
                try:
                    raw = extract_aligned_values(series, ev["date"], lookback, lookforward, frequency, agg_method)
                    if any(not pd.isna(v) for v in raw.values()):
                        aligned_data[ev["label"]] = apply_transformation(raw, field_mode, baseline_mode)
                except Exception:
                    pass

            if not aligned_data:
                continue

            fig = _create_mpl_chart(ticker_id, display_name, valid_events, aligned_data, field_mode, frequency)
            img_name = "{}_{}.png".format(
                cat_name.replace(" ", "_").replace("&", "and"),
                ticker_id.replace(" ", "_").replace("/", "_").replace(".", "_"),
            )
            fig.savefig(str(img_dir / img_name), dpi=180, bbox_inches="tight", facecolor="white")
            plt.close(fig)

            section_tex += "\\begin{center}\n"
            section_tex += "\\includegraphics[width=0.95\\textwidth]{img/%s}\n" % img_name
            section_tex += "\\end{center}\n\\vspace{0.3cm}\n\n"
            chart_count += 1
            total_charts += 1

        total_tickers += len(cat_tickers)
        if chart_count > 0:
            section_tex += "\\newpage\n"
            latex_sections.append(section_tex)

    if total_charts == 0:
        print("ERROR: No se genero ninguna grafica.")
        shutil.rmtree(tmp_dir, ignore_errors=True)
        return None

    # ── LaTeX ─────────────────────────────────────────────────────────────────
    freq_es = {"daily": "diaria", "weekly": "semanal", "monthly": "mensual", "annual": "anual"}

    # En modo flat el conteo de "categorias" no aplica — usamos "secciones".
    section_label = "secciones" if flat_mode else "categorias"
    subtitle_line = ("Analisis comparativo de activos financieros\\\\ante eventos de estres"
                     if not flat_mode else
                     "Analisis comparativo de tus series\\\\ante eventos definidos")
    if report_subtitle:
        subtitle_line = report_subtitle.replace("\n", "\\\\")

    cover = r"""
\begin{document}
\begin{titlepage}
\centering
\vspace*{4cm}
{\Huge\bfseries\color{azulcorp} Event Study Report}\\[1cm]
{\Large\color{gris} %s}\\[2cm]
{\large\color{gris} %s}\\[1cm]
{\normalsize\color{gris} %d %s \quad|\quad %d tickers \quad|\quad %d eventos}\\[0.5cm]
{\normalsize\color{gris} Ventana: %d / %d periodos (%s)}
\vfill
{\small\color{gris} Generado automaticamente por Event Study Dashboard}
\end{titlepage}
\tableofcontents
\newpage
""" % (subtitle_line, date.today().strftime("%d/%m/%Y"), len(latex_sections), section_label,
       total_tickers, len(valid_events), lookback, lookforward, freq_es.get(frequency, frequency))

    events_sec = "\\section{Eventos Analizados}\n\\begin{itemize}\n"
    for ev in valid_events:
        events_sec += "  \\item \\textbf{%s} --- %s\n" % (ev["label"].replace("&", "\\&"), ev["date"])
    events_sec += "\\end{itemize}\n\\newpage\n\n"

    full_tex = LATEX_PREAMBLE + "\n" + cover + "\n" + events_sec + "\n".join(latex_sections) + "\n\\end{document}\n"

    tex_path = tmp_dir / "report.tex"
    with open(str(tex_path), "w", encoding="utf-8") as f:
        f.write(full_tex)

    # ── Compilar ──────────────────────────────────────────────────────────────
    try:
        for _ in range(2):
            subprocess.run(
                ["pdflatex", "-interaction=nonstopmode", "-output-directory", str(tmp_dir), str(tex_path)],
                capture_output=True, text=True, timeout=120,
            )
        pdf_tmp = tmp_dir / "report.pdf"
        if not pdf_tmp.exists():
            print("ERROR: pdflatex no genero el PDF.")
            return None

        pdf_final = output_dir / "Event_Study_Report_{}.pdf".format(today_str)
        shutil.copy2(str(pdf_tmp), str(pdf_final))
        shutil.rmtree(tmp_dir, ignore_errors=True)
        print("PDF: {}".format(pdf_final))
        return pdf_final
    except FileNotFoundError:
        print("ERROR: pdflatex no encontrado.")
        shutil.rmtree(tmp_dir, ignore_errors=True)
        return None
    except subprocess.TimeoutExpired:
        print("ERROR: pdflatex timeout.")
        shutil.rmtree(tmp_dir, ignore_errors=True)
        return None


# ══════════════════════════════════════════════════════════════════════════════
# CLI
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    events = _EVENTS_GLOBAL

    if len(sys.argv) > 1:
        xlsx_path = sys.argv[1]
        print("Leyendo: {}".format(xlsx_path))
        data_df = pd.read_excel(xlsx_path, index_col=0, sheet_name=0)
        data_df.index = pd.DatetimeIndex(data_df.index)
    else:
        print("Uso: python generate_report.py <archivo.xlsx>")
        print("     O ejecuta desde el dashboard Streamlit.")
        sys.exit(0)

    print("{} tickers, {} obs, {} eventos".format(len(data_df.columns), len(data_df), len(events)))
    pdf = build_full_report(events, data_df)
    if pdf:
        print("Listo: {}".format(pdf))
    else:
        sys.exit(1)
