import html
import io

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(
    page_title="Dashboard Cukai",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================
# Constants
# ============================================================
COL_MONTH = "BULAN"
COL_COMPANY = "NAMA SYARIKAT"
COL_VIOLATION = "JENIS PELANGGARAN SYARAT"
COL_CODE = "KOD HASIL"
COL_AMOUNT = "JUMLAH CUKAI"
COL_DATE = "TARIKH BAYARAN"
REQUIRED = [COL_MONTH, COL_COMPANY, COL_VIOLATION, COL_CODE, COL_AMOUNT, COL_DATE]

SUPPORTED = ["csv", "tsv", "txt", "xlsx", "xlsm", "xls", "ods", "json"]
EXCEL_ENGINES = {"xlsx": "openpyxl", "xlsm": "openpyxl", "xls": "xlrd", "ods": "odf"}

MONTHS = {
    "JANUARI": 1, "JAN": 1, "JANUARY": 1,
    "FEBRUARI": 2, "FEB": 2, "FEBRUARY": 2,
    "MAC": 3, "MAR": 3, "MARCH": 3,
    "APRIL": 4, "APR": 4,
    "MEI": 5, "MAY": 5,
    "JUN": 6, "JUNE": 6,
    "JULAI": 7, "JUL": 7, "JULY": 7,
    "OGOS": 8, "OGO": 8, "AUG": 8, "AUGUST": 8,
    "SEPTEMBER": 9, "SEP": 9, "SEPT": 9,
    "OKTOBER": 10, "OKT": 10, "OCT": 10, "OCTOBER": 10,
    "NOVEMBER": 11, "NOV": 11,
    "DISEMBER": 12, "DIS": 12, "DEC": 12, "DECEMBER": 12,
}
MONTH_LABELS = ["Jan", "Feb", "Mac", "Apr", "Mei", "Jun", "Jul", "Ogo", "Sep", "Okt", "Nov", "Dis"]

PALETTE = ["#4F46E5", "#06B6D4", "#10B981", "#F59E0B", "#EF4444",
           "#8B5CF6", "#EC4899", "#14B8A6", "#F97316", "#64748B"]
PRIMARY, GREEN, RED = PALETTE[0], "#10B981", "#EF4444"

# ============================================================
# Styling
# ============================================================
st.markdown(
    """
<style>
.block-container {padding-top: 1.5rem; padding-bottom: 3rem; max-width: 1400px;}
.hero {background: linear-gradient(120deg, #4F46E5 0%, #06B6D4 100%);
       padding: 1.5rem 2rem; border-radius: 18px; color: #fff; margin-bottom: 1.2rem;}
.hero .title {font-size: 2rem; font-weight: 800; line-height: 1.2;}
.hero .sub {margin-top: .35rem; opacity: .92; font-size: .95rem;}
.kpi {background: #fff; border: 1px solid #E5E7EB; border-radius: 16px;
      padding: 1rem 1.2rem; box-shadow: 0 1px 3px rgba(0,0,0,.06); height: 100%;}
.kpi .label {font-size: .74rem; color: #6B7280; text-transform: uppercase; letter-spacing: .06em;}
.kpi .value {font-size: 1.55rem; font-weight: 700; color: #111827; margin-top: .15rem;}
.kpi .delta {font-size: .8rem; margin-top: .25rem; font-weight: 600;}
.kpi .body {font-size: .9rem; color: #374151; margin-top: .3rem;}
.up {color: #059669;} .down {color: #DC2626;} .flat {color: #6B7280;}
button[data-baseweb="tab"] {font-weight: 600;}
h5 {margin-top: .4rem;}
</style>
""",
    unsafe_allow_html=True,
)


def style(fig, height=380, legend=True):
    layout = dict(
        template="plotly_white",
        height=height,
        margin=dict(l=10, r=10, t=30, b=10),
        font=dict(family="Inter, Segoe UI, sans-serif", size=13),
        colorway=PALETTE,
        showlegend=legend,
        hoverlabel=dict(bgcolor="white"),
    )
    if legend:
        layout["legend"] = dict(orientation="h", yanchor="bottom", y=1.02, x=0)
    fig.update_layout(**layout)
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(gridcolor="#EEF2F7")
    return fig


def show(fig):
    st.plotly_chart(fig, width="stretch")


def kpi(col, label, value, delta=None, sub=""):
    d_html = ""
    if delta is not None:
        cls = "up" if delta > 0 else "down" if delta < 0 else "flat"
        arrow = "▲" if delta > 0 else "▼" if delta < 0 else "■"
        d_html = f'<div class="delta {cls}">{arrow} {abs(delta):.1f}% {html.escape(sub)}</div>'
    elif sub:
        d_html = f'<div class="delta flat">{html.escape(sub)}</div>'
    col.markdown(
        f'<div class="kpi"><div class="label">{html.escape(label)}</div>'
        f'<div class="value">{html.escape(value)}</div>{d_html}</div>',
        unsafe_allow_html=True,
    )


def rm(x):
    return f"RM {x:,.2f}"


def short(x):
    a = abs(x)
    if a >= 1e9:
        return f"{x / 1e9:.1f}B"
    if a >= 1e6:
        return f"{x / 1e6:.1f}M"
    if a >= 1e3:
        return f"{x / 1e3:.1f}K"
    return f"{x:,.0f}"


# ============================================================
# Data loading
# ============================================================
@st.cache_data(show_spinner=False)
def excel_sheets(name, data):
    ext = name.rsplit(".", 1)[-1].lower()
    return pd.ExcelFile(io.BytesIO(data), engine=EXCEL_ENGINES[ext]).sheet_names


@st.cache_data(show_spinner="Reading file…")
def read_file(name, data, sheet=None, header_row=1):
    ext = name.rsplit(".", 1)[-1].lower()
    hdr = int(header_row) - 1
    if ext in EXCEL_ENGINES:
        df = pd.read_excel(io.BytesIO(data), sheet_name=sheet or 0,
                           header=hdr, engine=EXCEL_ENGINES[ext])
    elif ext == "json":
        try:
            df = pd.read_json(io.BytesIO(data))
        except ValueError:
            df = pd.read_json(io.BytesIO(data), lines=True)
    else:  # csv, tsv, txt
        df = None
        for enc in ("utf-8-sig", "latin-1"):
            try:
                df = pd.read_csv(io.BytesIO(data), sep=None, engine="python",
                                 encoding=enc, header=hdr)
                break
            except UnicodeDecodeError:
                continue
        if df is None:
            raise ValueError("Could not decode the file. Try saving it as UTF-8.")
    # Normalise headers: strip tabs / extra spaces, uppercase
    df.columns = [" ".join(str(c).split()).upper() for c in df.columns]
    return df


# ============================================================
# Cleaning helpers
# ============================================================
def clean_amount(s):
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce")
    cleaned = s.astype(str).str.replace(r"[^\d.\-]", "", regex=True)
    return pd.to_numeric(cleaned, errors="coerce")


def month_number(v):
    if pd.isna(v):
        return None
    t = str(v).strip().upper()
    if t in MONTHS:
        return MONTHS[t]
    try:
        n = int(float(t))
        if 1 <= n <= 12:
            return n
    except ValueError:
        pass
    return None


def build_period_date(df, basis):
    """First day of the month for each row, from TARIKH BAYARAN or BULAN."""
    pay = pd.to_datetime(df[COL_DATE], dayfirst=True, errors="coerce")
    if basis == COL_DATE:
        return pay.dt.to_period("M").dt.to_timestamp()
    months = pd.to_numeric(df[COL_MONTH].map(month_number), errors="coerce")
    combined = pd.to_datetime(
        pd.DataFrame({"year": pay.dt.year, "month": months, "day": 1}), errors="coerce"
    )
    direct = pd.to_datetime(df[COL_MONTH], dayfirst=True, errors="coerce")
    direct = direct.where(months.isna()).dt.to_period("M").dt.to_timestamp()
    return combined.fillna(direct)


def aggregate(d, monthly):
    if d.empty:
        return pd.DataFrame(columns=["LABEL", "JUMLAH", "REKOD"])
    if monthly:
        g = d.groupby("PERIOD").agg(JUMLAH=(COL_AMOUNT, "sum"), REKOD=(COL_AMOUNT, "size"))
        g = g.reindex(pd.date_range(g.index.min(), g.index.max(), freq="MS"), fill_value=0)
        g = g.rename_axis("PERIOD").reset_index()
        g["LABEL"] = g["PERIOD"].dt.strftime("%b %Y")
    else:
        g = d.groupby("TAHUN").agg(JUMLAH=(COL_AMOUNT, "sum"), REKOD=(COL_AMOUNT, "size"))
        g = g.reindex(range(int(g.index.min()), int(g.index.max()) + 1), fill_value=0)
        g = g.rename_axis("TAHUN").reset_index()
        g["LABEL"] = g["TAHUN"].astype(str)
    return g


# ============================================================
# Chart builders
# ============================================================
def trend_chart(t, monthly, color=PRIMARY):
    fig = go.Figure()
    if monthly:
        fig.add_trace(go.Scatter(
            x=t["LABEL"], y=t["JUMLAH"], mode="lines+markers", fill="tozeroy",
            line=dict(color=color, width=3), marker=dict(size=6),
            fillcolor="rgba(79,70,229,.12)",
            hovertemplate="%{x}<br>RM %{y:,.2f}<extra></extra>",
        ))
    else:
        fig.add_trace(go.Bar(
            x=t["LABEL"], y=t["JUMLAH"], marker_color=color,
            text=t["JUMLAH"].map(short), textposition="outside", cliponaxis=False,
            hovertemplate="%{x}<br>RM %{y:,.2f}<extra></extra>",
        ))
    # Category axis: only periods that exist in the data are shown
    fig.update_xaxes(type="category")
    fig.update_yaxes(title="RM")
    return style(fig, 380, legend=False)


def hbar(d, label, color=PRIMARY, height=420):
    d = d.sort_values(COL_AMOUNT)
    full = d[label].astype(str)
    lbl = full.where(full.str.len() <= 40, full.str.slice(0, 38) + "…")
    fig = go.Figure(go.Bar(
        x=d[COL_AMOUNT], y=lbl, customdata=full, orientation="h", marker_color=color,
        text=d[COL_AMOUNT].map(short), textposition="outside", cliponaxis=False,
        hovertemplate="%{customdata}<br>RM %{x:,.2f}<extra></extra>",
    ))
    fig.update_yaxes(automargin=True)
    return style(fig, height, legend=False)


# ============================================================
# Sidebar – step 1: file
# ============================================================
with st.sidebar:
    st.markdown("## 📊 Dashboard Cukai")
    st.caption("Upload → choose view → explore")
    uploaded = st.file_uploader(
        "1 · Upload data file", type=SUPPORTED,
        help="CSV, TSV, TXT, Excel (.xlsx/.xlsm/.xls), OpenDocument (.ods) or JSON.",
    )
    sheet, header_row = None, 1
    if uploaded is not None:
        file_bytes = uploaded.getvalue()
        ext = uploaded.name.rsplit(".", 1)[-1].lower()
        if ext in EXCEL_ENGINES:
            try:
                sheets = excel_sheets(uploaded.name, file_bytes)
            except Exception as e:
                st.error(f"Could not open the workbook: {e}")
                st.stop()
            sheet = st.selectbox("Sheet", sheets) if len(sheets) > 1 else sheets[0]
        if ext != "json":
            with st.expander("Advanced file options"):
                header_row = st.number_input(
                    "Header row number", min_value=1, max_value=50, value=1,
                    help="Change this if your column titles are not on the first row.",
                )

# ============================================================
# Load data or show landing page
# ============================================================
if uploaded is not None:
    try:
        raw = read_file(uploaded.name, file_bytes, sheet, header_row)
    except Exception as e:
        st.error(f"Failed to read the file: {e}")
        st.stop()
    source_name = uploaded.name
else:
    st.markdown(
        '<div class="hero"><div class="title">Dashboard Cukai</div>'
        '<div class="sub">Upload a file and get an interactive monthly or yearly dashboard instantly.</div></div>',
        unsafe_allow_html=True,
    )
    c1, c2, c3 = st.columns(3)
    c1.markdown('<div class="kpi"><div class="label">Step 1</div><div class="value">Upload</div>'
                '<div class="body">CSV, TSV, Excel, ODS or JSON from the sidebar.</div></div>',
                unsafe_allow_html=True)
    c2.markdown('<div class="kpi"><div class="label">Step 2</div><div class="value">Choose view</div>'
                '<div class="body">Switch between Bulanan (monthly) and Tahunan (yearly).</div></div>',
                unsafe_allow_html=True)
    c3.markdown('<div class="kpi"><div class="label">Step 3</div><div class="value">Explore</div>'
                '<div class="body">Filter, drill into companies, and download results.</div></div>',
                unsafe_allow_html=True)
    st.markdown("")
    st.markdown("**Expected columns:** " + " · ".join(f"`{c}`" for c in REQUIRED))
    st.stop()

missing = [c for c in REQUIRED if c not in raw.columns]
if missing:
    st.error(f"Missing column(s): {', '.join(missing)}")
    st.write("Columns found in your file:", list(raw.columns))
    st.info("If the column titles are not on the first row, set **Header row number** under "
            "*Advanced file options* in the sidebar.")
    st.stop()

# ============================================================
# Sidebar – step 2: view
# ============================================================
with st.sidebar:
    st.markdown("---")
    view = st.radio("2 · Dashboard view", ["Bulanan", "Tahunan"], horizontal=True)
    monthly = view == "Bulanan"
    basis = st.selectbox(
        "Time axis built from", [COL_DATE, COL_MONTH],
        help="TARIKH BAYARAN by default. Choose BULAN if that column defines your reporting month.",
    )
    top_n = st.slider("Top N items in charts", 1, 10, 10)

# ============================================================
# Prepare data
# ============================================================
df = raw.copy()
for c in (COL_COMPANY, COL_VIOLATION):
    df[c] = (df[c].fillna("").astype(str).str.split().str.join(" ")
             .replace("", "(Kosong)"))
df[COL_CODE] = (df[COL_CODE].fillna("(Kosong)").astype(str).str.strip()
                .str.replace(r"\.0$", "", regex=True))
df[COL_AMOUNT] = clean_amount(df[COL_AMOUNT])
df["PERIOD"] = build_period_date(df, basis)
df["TAHUN"] = df["PERIOD"].dt.year
df["BULAN_NO"] = df["PERIOD"].dt.month
df["TEMPOH"] = df["PERIOD"].dt.strftime("%Y-%m")

bad_dates = int(df["PERIOD"].isna().sum())
bad_amounts = int(df[COL_AMOUNT].isna().sum())
df[COL_AMOUNT] = df[COL_AMOUNT].fillna(0)

# ============================================================
# Sidebar – step 3: filters
# ============================================================
with st.sidebar:
    st.markdown("---")
    years = sorted(df["TAHUN"].dropna().astype(int).unique())
    sel_years = st.multiselect("3 · Tahun (empty = all)", years)
    with st.expander("More filters (empty = all)", expanded=False):
        sel_companies = st.multiselect("Nama Syarikat", sorted(df[COL_COMPANY].unique()))
        sel_violations = st.multiselect("Jenis Pelanggaran Syarat", sorted(df[COL_VIOLATION].unique()))
        sel_codes = st.multiselect("Kod Hasil", sorted(df[COL_CODE].unique()))

f = df
if sel_years:
    f = f[f["TAHUN"].isin(sel_years)]
if sel_companies:
    f = f[f[COL_COMPANY].isin(sel_companies)]
if sel_violations:
    f = f[f[COL_VIOLATION].isin(sel_violations)]
if sel_codes:
    f = f[f[COL_CODE].isin(sel_codes)]

if f.empty:
    st.warning("No data matches the current filters.")
    st.stop()

dated = f.dropna(subset=["PERIOD"]).copy()
dated["TAHUN"] = dated["TAHUN"].astype(int)
dated["BULAN_NO"] = dated["BULAN_NO"].astype(int)
trend = aggregate(dated, monthly)
period_name = "bulan" if monthly else "tahun"

# ============================================================
# Header + KPIs
# ============================================================
span = ""
if not dated.empty:
    span = f" · {dated['PERIOD'].min():%b %Y} – {dated['PERIOD'].max():%b %Y}"
st.markdown(
    f'<div class="hero"><div class="title">Dashboard Cukai · {view}</div>'
    f'<div class="sub">{html.escape(source_name)} · {len(f):,} rekod{span}</div></div>',
    unsafe_allow_html=True,
)

# Which years exist in the uploaded file (before filters)?
year_counts = df["TAHUN"].dropna().astype(int).value_counts().sort_index()
if not year_counts.empty:
    st.caption("Tahun dalam fail: " + " · ".join(f"{y} ({n:,} rekod)" for y, n in year_counts.items()))

if bad_dates or bad_amounts:
    st.warning(
        f"{bad_dates} row(s) have an unreadable date (excluded from time charts) and "
        f"{bad_amounts} row(s) have an unreadable amount (counted as RM 0)."
    )

total = f[COL_AMOUNT].sum()
k = st.columns(4)
kpi(k[0], "Jumlah Cukai", rm(total))
kpi(k[1], "Bilangan Rekod", f"{len(f):,}")
kpi(k[2], "Bilangan Syarikat", f"{f[COL_COMPANY].nunique():,}")
if len(trend) >= 1:
    last = trend.iloc[-1]
    delta = None
    sub = ""
    if len(trend) >= 2 and trend.iloc[-2]["JUMLAH"] != 0:
        prev = trend.iloc[-2]
        delta = (last["JUMLAH"] - prev["JUMLAH"]) / abs(prev["JUMLAH"]) * 100
        sub = f"vs {prev['LABEL']}"
    kpi(k[3], f"{period_name.capitalize()} terkini · {last['LABEL']}", rm(last["JUMLAH"]), delta, sub)
else:
    kpi(k[3], f"{period_name.capitalize()} terkini", "—")

st.markdown("")

# ============================================================
# Tabs
# ============================================================
tab_overview, tab_trend, tab_company, tab_vio, tab_data = st.tabs(
    ["📈 Ringkasan", "🔎 Trend", "🏢 Syarikat", "⚠️ Pelanggaran & Kod", "🗂️ Data"]
)

# ---------- Overview ----------
with tab_overview:
    st.markdown(f"##### Jumlah cukai mengikut {period_name}")
    if trend.empty:
        st.info("No rows with a valid date to plot.")
    else:
        show(trend_chart(trend, monthly))

    c1, c2 = st.columns([3, 2])
    with c1:
        st.markdown(f"##### Top {top_n} syarikat")
        comp = (f.groupby(COL_COMPANY, as_index=False)[COL_AMOUNT].sum()
                .nlargest(top_n, COL_AMOUNT))
        show(hbar(comp, COL_COMPANY))
    with c2:
        st.markdown("##### Pecahan jenis pelanggaran")
        v = f.groupby(COL_VIOLATION)[COL_AMOUNT].sum()
        v = v[v > 0].nlargest(top_n).reset_index()
        if v.empty:
            st.info("No positive amounts to display.")
        else:
            fig = px.pie(v, names=COL_VIOLATION, values=COL_AMOUNT, hole=0.55,
                         color_discrete_sequence=PALETTE)
            fig.update_traces(textinfo="percent", textposition="inside",
                              hovertemplate="%{label}<br>RM %{value:,.2f}<extra></extra>")
            style(fig, 420)
            fig.update_layout(legend=dict(orientation="h", y=-0.12, x=0))
            show(fig)

# ---------- Trend ----------
with tab_trend:
    if trend.empty:
        st.info("No rows with a valid date to plot.")
    else:
        t = trend.copy()
        t["PCT"] = t["JUMLAH"].pct_change().replace([np.inf, -np.inf], np.nan) * 100
        t["KUMULATIF"] = t["JUMLAH"].cumsum()

        c1, c2 = st.columns(2)
        with c1:
            st.markdown(f"##### Pertumbuhan ({'bulan ke bulan' if monthly else 'tahun ke tahun'}, %)")
            fig = go.Figure(go.Bar(
                x=t["LABEL"], y=t["PCT"],
                marker_color=[GREEN if (p >= 0) else RED for p in t["PCT"].fillna(0)],
                hovertemplate="%{x}<br>%{y:.1f}%<extra></extra>",
            ))
            fig.update_xaxes(type="category")
            show(style(fig, 360, legend=False))
        with c2:
            st.markdown("##### Jumlah kumulatif")
            fig = go.Figure(go.Scatter(
                x=t["LABEL"], y=t["KUMULATIF"], mode="lines", fill="tozeroy",
                line=dict(color=PALETTE[1], width=3), fillcolor="rgba(6,182,212,.12)",
                hovertemplate="%{x}<br>RM %{y:,.2f}<extra></extra>",
            ))
            fig.update_xaxes(type="category")
            show(style(fig, 360, legend=False))

        st.markdown("##### Corak bermusim: tahun × bulan")
        pv = (dated.pivot_table(index="TAHUN", columns="BULAN_NO", values=COL_AMOUNT, aggfunc="sum")
              .reindex(columns=range(1, 13)))
        pv.columns = MONTH_LABELS
        pv.index = pv.index.astype(str)
        fig = px.imshow(pv, aspect="auto", text_auto=".2s",
                        color_continuous_scale=["#EEF2FF", PRIMARY],
                        labels=dict(color="RM"))
        show(style(fig, 80 + 55 * len(pv), legend=False))

# ---------- Company ----------
with tab_company:
    comp_tbl = (f.groupby(COL_COMPANY)
                .agg(JUMLAH=(COL_AMOUNT, "sum"), REKOD=(COL_AMOUNT, "size"),
                     PURATA=(COL_AMOUNT, "mean"), TERAKHIR=("PERIOD", "max"))
                .sort_values("JUMLAH", ascending=False).reset_index())
    comp_tbl["PERATUS"] = comp_tbl["JUMLAH"] / total * 100 if total else 0

    st.markdown("##### Ranking syarikat")
    st.dataframe(
        comp_tbl, width="stretch", hide_index=True, height=360,
        column_config={
            COL_COMPANY: st.column_config.TextColumn("Nama Syarikat"),
            "JUMLAH": st.column_config.NumberColumn("Jumlah Cukai", format="RM %.2f"),
            "REKOD": st.column_config.NumberColumn("Rekod"),
            "PURATA": st.column_config.NumberColumn("Purata", format="RM %.2f"),
            "TERAKHIR": st.column_config.DatetimeColumn("Tempoh Terakhir", format="MMM YYYY"),
            "PERATUS": st.column_config.ProgressColumn(
                "Kongsian (%)", format="%.1f%%", min_value=0, max_value=100),
        },
    )

    st.markdown("##### Drill-down satu syarikat")
    company = st.selectbox("Pilih syarikat", comp_tbl[COL_COMPANY], label_visibility="collapsed")
    cd = dated[dated[COL_COMPANY] == company]
    m1, m2, m3 = st.columns(3)
    kpi(m1, "Jumlah Cukai", rm(f.loc[f[COL_COMPANY] == company, COL_AMOUNT].sum()))
    kpi(m2, "Bilangan Rekod", f"{(f[COL_COMPANY] == company).sum():,}")
    kpi(m3, "Pelanggaran berbeza", f"{f.loc[f[COL_COMPANY] == company, COL_VIOLATION].nunique():,}")
    st.markdown("")
    if cd.empty:
        st.info("This company has no rows with a valid date.")
    else:
        show(trend_chart(aggregate(cd, monthly), monthly, color=PALETTE[2]))

# ---------- Violations & codes ----------
with tab_vio:
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("##### Pembahagian pelanggaran syarat")
        vd = f.groupby(COL_VIOLATION)[COL_AMOUNT].sum()
        vd = vd[vd > 0].sort_values(ascending=False).reset_index()
        if vd.empty:
            st.info("No positive amounts to display.")
        else:
            fig = px.pie(vd, names=COL_VIOLATION, values=COL_AMOUNT, hole=0.55,
                         color_discrete_sequence=PALETTE)
            fig.update_traces(
                textinfo="percent", textposition="inside",
                hovertemplate="%{label}<br>RM %{value:,.2f}<br>%{percent}<extra></extra>",
            )
            style(fig, 420)
            fig.update_layout(legend=dict(orientation="h", y=-0.12, x=0))
            show(fig)
    with c2:
        st.markdown(f"##### Top {top_n} kod hasil")
        cc = f.groupby(COL_CODE, as_index=False)[COL_AMOUNT].sum().nlargest(top_n, COL_AMOUNT)
        show(hbar(cc, COL_CODE, PALETTE[5]))

    st.markdown(f"##### Pelanggaran mengikut {period_name}")
    if dated.empty:
        st.info("No rows with a valid date to plot.")
    else:
        top_v = f.groupby(COL_VIOLATION)[COL_AMOUNT].sum().nlargest(8).index
        sd = dated.assign(
            JENIS=np.where(dated[COL_VIOLATION].isin(top_v), dated[COL_VIOLATION], "Lain-lain"))
        key = "PERIOD" if monthly else "TAHUN"
        stack = sd.groupby([key, "JENIS"], as_index=False)[COL_AMOUNT].sum()
        stack["TEMPOH_LABEL"] = (stack[key].dt.strftime("%b %Y") if monthly
                                 else stack[key].astype(str))
        fig = px.bar(stack, x="TEMPOH_LABEL", y=COL_AMOUNT, color="JENIS",
                     color_discrete_sequence=PALETTE,
                     labels={COL_AMOUNT: "RM", "TEMPOH_LABEL": period_name.capitalize(), "JENIS": ""})
        fig.update_xaxes(type="category")
        show(style(fig, 420))

# ---------- Data ----------
def to_excel(frames):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for name, d in frames.items():
            d.to_excel(w, sheet_name=name[:31], index=False)
    return buf.getvalue()


with tab_data:
    summary = trend[["LABEL", "JUMLAH", "REKOD"]].rename(
        columns={"LABEL": "Tempoh", "JUMLAH": "Jumlah Cukai (RM)", "REKOD": "Bilangan Rekod"})
    export_df = f.drop(columns=["PERIOD"])

    st.markdown(f"##### Ringkasan mengikut {period_name}")
    st.dataframe(summary, width="stretch", hide_index=True,
                 column_config={"Jumlah Cukai (RM)": st.column_config.NumberColumn(format="RM %.2f")})

    with st.expander("Lihat data mentah (ditapis)"):
        st.dataframe(export_df, width="stretch", hide_index=True)

    d1, d2, _ = st.columns([1, 1, 3])
    d1.download_button(
        "⬇️ Download CSV", export_df.to_csv(index=False).encode("utf-8-sig"),
        file_name="data_ditapis.csv", mime="text/csv",
    )
    d2.download_button(
        "⬇️ Download Excel",
        to_excel({"Data Ditapis": export_df, "Ringkasan": summary,
                  "Syarikat": comp_tbl.drop(columns=["TERAKHIR"])}),
        file_name="dashboard_cukai.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
