#!/usr/bin/env python3
"""
Event Study Dashboard
=====================
Dashboard interactivo para comparar el comportamiento histórico de activos
financieros alrededor de fechas de eventos clave.

Fuentes de datos soportadas:
  - CSV / Excel (subida de archivo)
  - Bloomberg Terminal via blpapi

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

# ── Bloomberg (opcional) ───────────────────────────────────────────────────────
try:
    import blpapi
    BLOOMBERG_AVAILABLE = True
except ImportError:
    BLOOMBERG_AVAILABLE = False

# ══════════════════════════════════════════════════════════════════════════════
# CONSTANTES
# ══════════════════════════════════════════════════════════════════════════════

ASSET_TYPES   = ["equity", "rate", "fx", "commodity", "other"]
FIELD_MODES   = ["cumulative_return", "absolute_change", "pct_change", "price"]
BASELINE_MODES = ["base100", "base0", "none"]
FREQUENCIES   = ["daily", "weekly", "monthly", "annual"]

FREQ_LABEL = {"daily": "D", "weekly": "W", "monthly": "M", "annual": "Y"}

FIELD_MODE_LABELS = {
    "price":              "Precio Raw",
    "cumulative_return":  "Retorno Acumulado (%)",
    "pct_change":         "Cambio Porcentual (%)",
    "absolute_change":    "Cambio Absoluto",
}

BASELINE_LABELS = {
    "base100": "Base 100 (indexado al evento)",
    "base0":   "Base 0 (cambio desde el evento)",
    "none":    "Sin normalización",
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
]


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


def load_from_bloomberg(fields_map: dict, start_date: date, end_date: date):
    """
    Descarga datos históricos desde Bloomberg API.

    Parameters
    ----------
    fields_map : dict  {ticker_str: field_str}   e.g. {"SPX Index": "PX_LAST"}
    start_date / end_date : date

    Returns
    -------
    pd.DataFrame  (columnas = tickers), str error or None
    """
    if not BLOOMBERG_AVAILABLE:
        return None, (
            "blpapi no está instalado.\n"
            "Instálalo con:\n"
            "  pip install --index-url=https://bcms.bloomberg.com/pip/simple/ blpapi\n"
            "Luego reinicia la aplicación."
        )

    try:
        opts = blpapi.SessionOptions()
        opts.setServerHost("localhost")
        opts.setServerPort(8194)

        session = blpapi.Session(opts)
        if not session.start():
            return None, (
                "No se pudo iniciar la sesión Bloomberg. "
                "Asegúrate de que Bloomberg Terminal esté abierto."
            )
        if not session.openService("//blp/refdata"):
            session.stop()
            return None, "No se pudo abrir el servicio Bloomberg refdata."

        service = session.getService("//blp/refdata")

        # Agrupa por campo para minimizar requests
        by_field: dict[str, list[str]] = {}
        for ticker, field in fields_map.items():
            by_field.setdefault(field, []).append(ticker)

        all_series: dict[str, pd.Series] = {}

        for field, field_tickers in by_field.items():
            req = service.createRequest("HistoricalDataRequest")
            for t in field_tickers:
                req.getElement("securities").appendValue(t)
            req.getElement("fields").appendValue(field)
            req.set("startDt",  start_date.strftime("%Y%m%d"))
            req.set("endDt",    end_date.strftime("%Y%m%d"))
            req.set("periodicitySelection", "DAILY")
            req.set("nonTradingDayFillOption", "ACTIVE_DAYS_ONLY")

            session.sendRequest(req)

            while True:
                ev = session.nextEvent(3000)
                for msg in ev:
                    if msg.messageType() == blpapi.Name("HistoricalDataResponse"):
                        sec_data  = msg.getElement("securityData")
                        sec_name  = sec_data.getElementAsString("security")

                        if sec_data.hasElement("securityError"):
                            err_txt = sec_data.getElement("securityError").getElementAsString("message")
                            st.warning(f"Bloomberg: {sec_name} → {err_txt}")
                            continue

                        field_data = sec_data.getElement("fieldData")
                        dates_list, vals_list = [], []

                        for j in range(field_data.numValues()):
                            pt = field_data.getValue(j)
                            try:
                                d = pt.getElementAsDatetime("date")
                                v = pt.getElementAsFloat(field)
                                dates_list.append(
                                    pd.Timestamp(d.year, d.month, d.day)
                                )
                                vals_list.append(v)
                            except Exception:
                                pass

                        if dates_list:
                            all_series[sec_name] = pd.Series(
                                vals_list, index=dates_list, name=sec_name
                            )

                if ev.eventType() == blpapi.Event.RESPONSE:
                    break

        session.stop()

        if not all_series:
            return None, "Bloomberg no devolvió datos. Revisa los tickers y el rango de fechas."

        df = pd.DataFrame(all_series).sort_index()
        return df, None

    except Exception as e:
        return None, f"Error Bloomberg ({type(e).__name__}): {e}"


# ══════════════════════════════════════════════════════════════════════════════
# MOTOR DE VENTANAS DE EVENTOS
# ══════════════════════════════════════════════════════════════════════════════

def build_target_dates(event_date, lookback: int, lookforward: int, frequency: str):
    """Devuelve lista de (período_relativo, fecha_objetivo) para una ventana."""
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
    """Devuelve la última fecha disponible <= target_ts."""
    prior = series_index[series_index <= target_ts]
    return prior[-1] if len(prior) > 0 else None


def extract_aligned_values(series: pd.Series, event_date, lookback, lookforward, frequency):
    """
    Extrae valores alineados (sin transformar) para un evento.
    Retorna dict {período: valor}.
    """
    target_dates = build_target_dates(event_date, lookback, lookforward, frequency)
    result = {}
    for (period, target_ts) in target_dates:
        mapped = get_prior_observation(target_ts, series.index)
        result[period] = series.loc[mapped] if mapped is not None else np.nan
    return result


# ══════════════════════════════════════════════════════════════════════════════
# TRANSFORMACIONES
# ══════════════════════════════════════════════════════════════════════════════

def apply_transformation(raw_dict: dict, field_mode: str, baseline_mode: str) -> dict:
    """
    Transforma valores alineados según field_mode y baseline_mode.
    raw_dict: {período: valor_raw},  período 0 = anchor del evento.
    """
    anchor = raw_dict.get(0, np.nan)
    if pd.isna(anchor):
        return {p: np.nan for p in raw_dict}
    if anchor == 0 and field_mode in ("cumulative_return", "pct_change"):
        return {p: np.nan for p in raw_dict}

    out = {}
    for p, v in raw_dict.items():
        if pd.isna(v):
            out[p] = np.nan
            continue

        if field_mode in ("cumulative_return", "pct_change"):
            out[p] = (v / anchor - 1) * 100
        elif field_mode == "absolute_change":
            out[p] = v - anchor
        elif field_mode == "price":
            if baseline_mode == "base100":
                out[p] = v / anchor * 100
            elif baseline_mode == "base0":
                out[p] = v - anchor
            else:
                out[p] = v
        else:
            out[p] = v

    return out


# ══════════════════════════════════════════════════════════════════════════════
# PLOTLY
# ══════════════════════════════════════════════════════════════════════════════

def build_period_labels(periods: list, freq: str) -> list:
    fl = FREQ_LABEL.get(freq, "T")
    labels = []
    for p in periods:
        if p == 0:
            labels.append("Evento (T0)")
        elif p > 0:
            labels.append(f"{fl}+{p}")
        else:
            labels.append(f"{fl}{p}")
    return labels


def create_event_chart(ticker_id, display_name, events, aligned_data, field_mode, frequency):
    """Crea figura Plotly para un ticker con una línea por evento."""
    fig = go.Figure()

    y_title = FIELD_MODE_LABELS.get(field_mode, field_mode)
    all_periods = sorted({p for ev_data in aligned_data.values() for p in ev_data})
    period_labels = build_period_labels(all_periods, frequency)
    label_map = dict(zip(all_periods, period_labels))

    for i, ev in enumerate(events):
        ev_label = ev.get("label", f"Evento {i+1}")
        ev_date  = ev.get("date", "")
        color    = EVENT_COLORS[i % len(EVENT_COLORS)]

        if ev_label not in aligned_data:
            continue

        ev_data = aligned_data[ev_label]
        y_vals  = [ev_data.get(p, np.nan) for p in all_periods]

        hover_texts = []
        for j, p in enumerate(all_periods):
            v = ev_data.get(p, np.nan)
            val_str = f"{v:,.4f}" if not pd.isna(v) else "Sin dato"
            hover_texts.append(
                f"<b>{ev_label}</b><br>"
                f"Fecha evento: {ev_date}<br>"
                f"Período: {label_map[p]}<br>"
                f"{y_title}: {val_str}"
            )

        fig.add_trace(go.Scatter(
            x=all_periods,
            y=y_vals,
            mode="lines+markers",
            name=f"{ev_label} ({ev_date})",
            line=dict(color=color, width=2.5),
            marker=dict(size=5, color=color),
            hovertext=hover_texts,
            hoverinfo="text",
            connectgaps=False,
        ))

    # Línea vertical en el anchor (T=0)
    fig.add_vline(
        x=0,
        line_dash="dash",
        line_color="rgba(80,80,80,0.55)",
        line_width=1.8,
        annotation_text="  T=0",
        annotation_position="top",
        annotation_font=dict(size=11, color="#666"),
    )

    # Línea horizontal de referencia
    if field_mode in ("cumulative_return", "pct_change", "absolute_change"):
        fig.add_hline(y=0, line_color="rgba(150,150,150,0.35)", line_width=1)
    elif field_mode == "price" and "base100" in (
        [tk.get("baseline_mode") for tk in st.session_state.get("tickers", [])]
    ):
        fig.add_hline(y=100, line_color="rgba(150,150,150,0.35)", line_width=1)

    # Ticks del eje X
    step = max(1, len(all_periods) // 12)
    tick_vals = [p for i, p in enumerate(all_periods) if i % step == 0 or p == 0]
    tick_text = [label_map[p] for p in tick_vals]

    fig.update_layout(
        title=dict(
            text=f"<b>{display_name or ticker_id}</b>",
            font=dict(size=15, color="#0d1b2a"),
            x=0,
        ),
        xaxis=dict(
            title="Período relativo al evento",
            tickvals=tick_vals,
            ticktext=tick_text,
            gridcolor="#ebebeb",
            zeroline=False,
            tickfont=dict(size=11),
        ),
        yaxis=dict(
            title=y_title,
            gridcolor="#ebebeb",
            tickfont=dict(size=11),
        ),
        plot_bgcolor="white",
        paper_bgcolor="white",
        hovermode="closest",
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.42,
            xanchor="center",
            x=0.5,
            font=dict(size=11),
            bordercolor="#ddd",
            borderwidth=1,
        ),
        height=490,
        margin=dict(l=70, r=30, t=55, b=140),
    )
    return fig


# ══════════════════════════════════════════════════════════════════════════════
# CSV DE MUESTRA
# ══════════════════════════════════════════════════════════════════════════════

def generate_sample_csv() -> str:
    """Genera CSV sintético realista para testing (2015-2024)."""
    np.random.seed(42)
    dates = pd.bdate_range("2015-01-02", "2024-12-31")
    n = len(dates)
    dates_str = [d.strftime("%Y-%m-%d") for d in dates]

    def idx(d):
        return dates_str.index(d) if d in dates_str else None

    # SPX Index
    r = np.random.normal(0.0004, 0.010, n)
    for sd, sv in [("2020-03-16", -0.12), ("2020-03-17", -0.06),
                   ("2022-02-24", -0.030), ("2023-03-10", -0.025)]:
        i = idx(sd)
        if i is not None:
            r[i] = sv
    spx = 2000.0 * np.exp(np.cumsum(r))

    # US 10Y Yield
    ch = np.random.normal(0.00005, 0.035, n)
    for sd, sv in [("2020-03-16", -0.25), ("2022-02-24", 0.08), ("2022-06-15", 0.15)]:
        i = idx(sd)
        if i is not None:
            ch[i] = sv
    us10y = np.maximum(0.05, 2.0 + np.cumsum(ch))

    # USDMXN
    fx = np.random.normal(0.00008, 0.007, n)
    for sd, sv in [("2020-03-16", 0.085), ("2022-02-24", 0.012)]:
        i = idx(sd)
        if i is not None:
            fx[i] = sv
    usdmxn = 14.5 * np.exp(np.cumsum(fx))

    # Gold
    gr = np.random.normal(0.0003, 0.009, n)
    for sd, sv in [("2020-03-16", -0.04), ("2022-02-24", 0.03)]:
        i = idx(sd)
        if i is not None:
            gr[i] = sv
    gold = 1180.0 * np.exp(np.cumsum(gr))

    df = pd.DataFrame({
        "SPX Index":        np.round(spx, 2),
        "USGG10YR Index":   np.round(us10y, 3),
        "USDMXN Curncy":    np.round(usdmxn, 4),
        "XAU Curncy":       np.round(gold, 2),
    }, index=dates)
    df.index.name = "Date"
    return df.to_csv()


# ══════════════════════════════════════════════════════════════════════════════
# APP PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════════

def main():
    st.set_page_config(
        page_title="Event Study Dashboard",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    st.markdown("""
    <style>
    .main-hdr  { font-size:1.9rem; font-weight:700; color:#0d1b2a; line-height:1.2; }
    .sub-hdr   { color:#555; font-size:0.92rem; margin-top:4px; margin-bottom:1.4rem; }
    .sec-title { font-size:0.78rem; font-weight:600; text-transform:uppercase;
                 letter-spacing:.06em; color:#888; margin-bottom:6px; margin-top:2px; }
    .status-ok   { background:#e8f5e9; border-left:3px solid #43a047; padding:7px 11px;
                   border-radius:4px; color:#1b5e20; font-size:0.84rem; margin-bottom:6px; }
    .status-warn { background:#fffde7; border-left:3px solid #fbc02d; padding:7px 11px;
                   border-radius:4px; color:#6d4c00; font-size:0.84rem; margin-bottom:6px; }
    div[data-testid="stExpander"] { border:1px solid #e8e8e8; border-radius:6px; }
    </style>
    """, unsafe_allow_html=True)

    st.markdown('<div class="main-hdr">📊 Event Study Dashboard</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="sub-hdr">Compara el comportamiento histórico de activos financieros '
        'alrededor de fechas de eventos clave — con soporte Bloomberg y CSV/Excel.</div>',
        unsafe_allow_html=True,
    )

    # ── Estado inicial ─────────────────────────────────────────────────────────
    if "events" not in st.session_state:
        st.session_state.events = [
            {"label": "COVID Crash",       "date": "2020-03-16"},
            {"label": "Russia-Ukraine",    "date": "2022-02-24"},
            {"label": "SVB Crisis",        "date": "2023-03-10"},
        ]
    if "tickers" not in st.session_state:
        st.session_state.tickers = [
            {
                "ticker": "SPX Index", "display_name": "S&P 500",
                "csv_column": "SPX Index", "bloomberg_field": "PX_LAST",
                "asset_type": "equity",
                "field_mode": "cumulative_return", "baseline_mode": "base100",
            },
            {
                "ticker": "USGG10YR Index", "display_name": "US 10Y Yield",
                "csv_column": "USGG10YR Index", "bloomberg_field": "PX_LAST",
                "asset_type": "rate",
                "field_mode": "absolute_change", "baseline_mode": "base0",
            },
        ]

    # ══════════════════════════════════════════════════════════════════════════
    # SIDEBAR
    # ══════════════════════════════════════════════════════════════════════════
    with st.sidebar:
        st.markdown("### ⚙️ Configuración")

        # ── Fuente de datos ──────────────────────────────────────────────────
        st.markdown('<div class="sec-title">Fuente de datos</div>', unsafe_allow_html=True)
        data_source  = st.radio("", ["📂 CSV / Excel", "🔵 Bloomberg"], label_visibility="collapsed")
        use_bloomberg = "Bloomberg" in data_source
        data_df = None

        if not use_bloomberg:
            st.caption("Sube tu archivo de precios históricos.")
            uploaded = st.file_uploader("", type=["csv", "xlsx", "xls"], label_visibility="collapsed")
            if uploaded:
                data_df, err = load_from_file(uploaded)
                if err:
                    st.error(f"Error al cargar: {err}")
                else:
                    st.markdown(
                        f'<div class="status-ok">✅ {len(data_df.columns)} series · '
                        f'{len(data_df):,} fechas<br>'
                        f'{data_df.index[0].date()} → {data_df.index[-1].date()}</div>',
                        unsafe_allow_html=True,
                    )
                    with st.expander("Vista previa de datos"):
                        st.dataframe(data_df.tail(5), use_container_width=True)
        else:
            if BLOOMBERG_AVAILABLE:
                st.markdown('<div class="status-ok">✅ blpapi detectado</div>', unsafe_allow_html=True)
            else:
                st.markdown('<div class="status-warn">⚠️ blpapi no encontrado</div>', unsafe_allow_html=True)
                with st.expander("Instrucciones de instalación"):
                    st.code(
                        "pip install --index-url=https://bcms.bloomberg.com/pip/simple/ blpapi",
                        language="bash",
                    )
                    st.caption("Requiere Bloomberg Terminal abierto en el mismo equipo.")
            st.caption("Rango de fechas para descargar:")
            bbg_start = st.date_input("Desde", value=date(2010, 1, 1), key="bbg_s")
            bbg_end   = st.date_input("Hasta", value=date.today(), key="bbg_e")

        st.divider()

        # ── Eventos ──────────────────────────────────────────────────────────
        st.markdown('<div class="sec-title">📅 Eventos</div>', unsafe_allow_html=True)

        remove_ev = None
        for i, ev in enumerate(st.session_state.events):
            with st.expander(f"🔴 {ev['label']}", expanded=(i == 0)):
                c1, c2 = st.columns([3, 2])
                with c1:
                    new_label = st.text_input("Nombre", value=ev["label"], key=f"el_{i}")
                with c2:
                    new_date = st.text_input("Fecha", value=ev["date"], key=f"ed_{i}",
                                             help="Formato YYYY-MM-DD")
                st.session_state.events[i] = {"label": new_label, "date": new_date}
                if len(st.session_state.events) > 1:
                    if st.button("🗑 Eliminar evento", key=f"re_{i}", use_container_width=True):
                        remove_ev = i

        if remove_ev is not None:
            st.session_state.events.pop(remove_ev)
            st.rerun()

        if st.button("➕ Agregar evento", use_container_width=True):
            n = len(st.session_state.events) + 1
            st.session_state.events.append({"label": f"Evento {n}", "date": "2024-01-01"})
            st.rerun()

        st.divider()

        # ── Ventana de análisis ───────────────────────────────────────────────
        st.markdown('<div class="sec-title">🪟 Ventana de análisis</div>', unsafe_allow_html=True)
        frequency = st.selectbox(
            "Frecuencia",
            FREQUENCIES,
            format_func=lambda x: {
                "daily": "Diaria", "weekly": "Semanal",
                "monthly": "Mensual", "annual": "Anual"
            }[x],
        )
        c1, c2 = st.columns(2)
        with c1:
            lookback    = st.number_input("Lookback",    min_value=1, max_value=500, value=30)
        with c2:
            lookforward = st.number_input("Lookforward", min_value=1, max_value=500, value=30)

        st.divider()

        # ── Tickers ───────────────────────────────────────────────────────────
        st.markdown('<div class="sec-title">📈 Tickers</div>', unsafe_allow_html=True)

        remove_tk = None
        for i, tk in enumerate(st.session_state.tickers):
            lbl = tk.get("display_name") or tk["ticker"]
            with st.expander(f"📈 {lbl}", expanded=(i == 0)):
                tk["ticker"]       = st.text_input("Ticker ID", value=tk["ticker"], key=f"tt_{i}")
                tk["display_name"] = st.text_input("Nombre",    value=tk.get("display_name",""), key=f"tn_{i}")

                if not use_bloomberg:
                    if data_df is not None and len(data_df.columns) > 0:
                        cols = list(data_df.columns)
                        def_col = tk.get("csv_column", cols[0])
                        if def_col not in cols:
                            def_col = cols[0]
                        tk["csv_column"] = st.selectbox(
                            "Columna del archivo", cols,
                            index=cols.index(def_col), key=f"tc_{i}"
                        )
                    else:
                        tk["csv_column"] = st.text_input(
                            "Nombre de columna", value=tk.get("csv_column",""), key=f"tc_{i}"
                        )
                else:
                    tk["bloomberg_field"] = st.text_input(
                        "Campo Bloomberg",
                        value=tk.get("bloomberg_field","PX_LAST"), key=f"tbf_{i}",
                        help="Ej: PX_LAST  |  YLD_YTM_MID  |  PX_BID",
                    )

                tk["asset_type"] = st.selectbox(
                    "Tipo de activo", ASSET_TYPES,
                    index=ASSET_TYPES.index(tk.get("asset_type","equity")),
                    format_func=str.capitalize, key=f"ta_{i}",
                )
                defs = ASSET_DEFAULTS.get(tk["asset_type"], ASSET_DEFAULTS["other"])

                tk["field_mode"] = st.selectbox(
                    "Transformación", FIELD_MODES,
                    index=FIELD_MODES.index(tk.get("field_mode", defs["field_mode"])),
                    format_func=lambda x: FIELD_MODE_LABELS.get(x, x), key=f"tfm_{i}",
                )
                tk["baseline_mode"] = st.selectbox(
                    "Modo baseline", BASELINE_MODES,
                    index=BASELINE_MODES.index(tk.get("baseline_mode", defs["baseline_mode"])),
                    format_func=lambda x: BASELINE_LABELS.get(x, x), key=f"tbm_{i}",
                )
                st.session_state.tickers[i] = tk

                if len(st.session_state.tickers) > 1:
                    if st.button("🗑 Eliminar ticker", key=f"rt_{i}", use_container_width=True):
                        remove_tk = i

        if remove_tk is not None:
            st.session_state.tickers.pop(remove_tk)
            st.rerun()

        if st.button("➕ Agregar ticker", use_container_width=True):
            n = len(st.session_state.tickers) + 1
            st.session_state.tickers.append({
                "ticker": f"TICKER_{n}", "display_name": f"Activo {n}",
                "csv_column": "", "bloomberg_field": "PX_LAST",
                "asset_type": "equity",
                "field_mode": "cumulative_return", "baseline_mode": "base100",
            })
            st.rerun()

        st.divider()
        run = st.button("🚀 Ejecutar análisis", type="primary", use_container_width=True)

    # ══════════════════════════════════════════════════════════════════════════
    # PANTALLA INICIAL (sin ejecutar)
    # ══════════════════════════════════════════════════════════════════════════
    if not run:
        col_a, col_b = st.columns([3, 2])

        with col_a:
            st.info("👈 Configura los parámetros en el panel izquierdo y presiona **🚀 Ejecutar análisis**.")

            with st.expander("📋 Formato esperado del CSV/Excel"):
                st.markdown("""
**Estructura del archivo:**
- **Primera columna**: Fechas (`YYYY-MM-DD` u otro formato estándar)
- **Columnas siguientes**: Una serie por columna, con el nombre del ticker como encabezado
- Los valores deben ser numéricos (precios, yields, tipos de cambio, etc.)

El archivo puede ser `.csv`, `.xlsx` o `.xls`.
                """)
                sample_preview = pd.DataFrame({
                    "SPX Index":      [3257.85, 3265.35, 2480.64, 2304.92],
                    "USGG10YR Index": [1.88,    1.90,    0.73,    0.76],
                    "USDMXN Curncy":  [18.87,   18.90,   24.51,   23.04],
                    "XAU Curncy":     [1520.0,  1547.8,  1477.2,  1680.0],
                }, index=pd.to_datetime(["2020-01-02","2020-01-03","2020-03-16","2020-04-01"]))
                sample_preview.index.name = "Date"
                st.dataframe(sample_preview, use_container_width=True)

                st.download_button(
                    "⬇️ Descargar CSV de muestra (datos sintéticos 2015-2024)",
                    generate_sample_csv(),
                    "sample_event_study_data.csv",
                    "text/csv",
                    use_container_width=True,
                )

        with col_b:
            st.markdown("#### 💡 Eventos de referencia")
            st.markdown("""
| Evento | Fecha |
|--------|-------|
| COVID Crash | 2020-03-16 |
| Russia-Ukraine | 2022-02-24 |
| SVB Crisis | 2023-03-10 |
| Fed pivot (75 bps) | 2022-06-15 |
| Lehman Brothers | 2008-09-15 |
| Taper Tantrum | 2013-05-22 |
| Brexit referendum | 2016-06-23 |
| 9/11 | 2001-09-11 |
            """)
        return

    # ══════════════════════════════════════════════════════════════════════════
    # VALIDACIÓN
    # ══════════════════════════════════════════════════════════════════════════
    events  = st.session_state.events
    tickers = st.session_state.tickers

    valid_events = []
    for ev in events:
        try:
            pd.Timestamp(ev["date"])
            valid_events.append(ev)
        except Exception:
            st.warning(f"⚠️ Fecha inválida para '{ev['label']}': '{ev['date']}' — omitido.")

    if not valid_events:
        st.error("❌ Ningún evento tiene una fecha válida. Usa el formato YYYY-MM-DD.")
        return
    if not tickers:
        st.error("❌ Agrega al menos un ticker en el panel de configuración.")
        return

    # ══════════════════════════════════════════════════════════════════════════
    # CARGA BLOOMBERG (si aplica)
    # ══════════════════════════════════════════════════════════════════════════
    if use_bloomberg:
        fields_map = {tk["ticker"]: tk.get("bloomberg_field","PX_LAST") for tk in tickers}
        all_event_ts = [pd.Timestamp(ev["date"]) for ev in valid_events]

        buf_map = {
            "daily":   (timedelta(days=lookback+10),          timedelta(days=lookforward+10)),
            "weekly":  (timedelta(weeks=lookback+2),          timedelta(weeks=lookforward+2)),
            "monthly": (relativedelta(months=lookback+1),     relativedelta(months=lookforward+1)),
            "annual":  (relativedelta(years=lookback+1),      relativedelta(years=lookforward+1)),
        }
        buf_back, buf_fwd = buf_map.get(frequency, (timedelta(days=90), timedelta(days=90)))

        needed_start = (min(all_event_ts) - buf_back).date()
        needed_end   = (max(all_event_ts) + buf_fwd).date()
        actual_start = min(bbg_start, needed_start)
        actual_end   = max(bbg_end,   needed_end)

        with st.spinner("🔄 Descargando datos de Bloomberg…"):
            data_df, err = load_from_bloomberg(fields_map, actual_start, actual_end)

        if err:
            st.error(f"❌ {err}")
            return
        if data_df is None or data_df.empty:
            st.error("❌ Bloomberg no devolvió datos.")
            return

        st.success(
            f"✅ Bloomberg: {len(data_df.columns)} series · "
            f"{len(data_df):,} observaciones · "
            f"{data_df.index[0].date()} → {data_df.index[-1].date()}"
        )

    elif data_df is None:
        st.error("❌ Sube un archivo CSV o Excel antes de ejecutar el análisis.")
        return

    # ══════════════════════════════════════════════════════════════════════════
    # PROCESAMIENTO Y VISUALIZACIÓN POR TICKER
    # ══════════════════════════════════════════════════════════════════════════
    st.markdown("---")
    st.markdown("## 📊 Resultados del Event Study")

    for tk_idx, tk in enumerate(tickers):
        ticker_id    = tk["ticker"]
        display_name = tk.get("display_name") or ticker_id
        field_mode   = tk.get("field_mode",    "cumulative_return")
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
            st.warning(f"⚠️ Serie vacía para '{display_name}'.")
            continue

        # Construir ventanas alineadas
        aligned_data: dict[str, dict] = {}
        skipped: list[str] = []

        for ev in valid_events:
            ev_label = ev["label"]
            try:
                raw = extract_aligned_values(
                    series, ev["date"], lookback, lookforward, frequency
                )
                if pd.isna(raw.get(0, np.nan)):
                    skipped.append(ev_label)
                    continue
                aligned_data[ev_label] = apply_transformation(raw, field_mode, baseline_mode)
            except Exception as e:
                st.warning(f"⚠️ '{ev_label}' — '{display_name}': {e}")

        if skipped:
            st.caption(
                f"⚠️ Sin datos en el anchor para '{display_name}': {', '.join(skipped)}"
            )
        if not aligned_data:
            st.error(f"❌ Sin datos procesados para **'{display_name}'**. "
                     "Verifica el rango de fechas y los eventos.")
            continue

        # Gráfico
        fig = create_event_chart(
            ticker_id, display_name, valid_events,
            aligned_data, field_mode, frequency
        )
        st.plotly_chart(fig, use_container_width=True)

        # Tabla de datos
        with st.expander(f"📋 Datos tabulares — {display_name}"):
            all_periods   = sorted({p for d in aligned_data.values() for p in d})
            period_labels = build_period_labels(all_periods, frequency)

            rows = {
                ev_label: [ev_data.get(p, np.nan) for p in all_periods]
                for ev_label, ev_data in aligned_data.items()
            }
            result_df = pd.DataFrame(rows, index=period_labels)
            result_df.index.name = "Período"

            st.dataframe(
                result_df.style.format("{:.4f}", na_rep="—"),
                use_container_width=True,
            )
            st.download_button(
                f"⬇️ Descargar datos — {display_name}",
                result_df.to_csv(),
                f"{ticker_id.replace(' ','_')}_event_study.csv",
                "text/csv",
                key=f"dl_{tk_idx}",
            )

    st.success("✅ Análisis completado.")


if __name__ == "__main__":
    main()
