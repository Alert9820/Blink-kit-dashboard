import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import warnings
warnings.filterwarnings('ignore')

st.set_page_config(
    page_title="BlinkIt Analytics",
    page_icon="🛒",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ═══════════════════════════════════════════════════════════
#  THEME & CSS
# ═══════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap');

*, html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif !important;
    box-sizing: border-box;
}

/* ── App background ── */
.stApp { background: #080A0F; }
.block-container { padding: 1.5rem 2rem 2rem 2rem !important; max-width: 100% !important; }

/* ── Sidebar ── */
section[data-testid="stSidebar"] {
    background: #0D0F17 !important;
    border-right: 1px solid #1E2235;
}
section[data-testid="stSidebar"] > div { padding: 1.5rem 1rem; }
section[data-testid="stSidebar"] * { color: #B0B7D4 !important; }
section[data-testid="stSidebar"] h2 { color: #FFD600 !important; font-size: 1.1rem !important; font-weight: 700 !important; }
section[data-testid="stSidebar"] label { color: #6B7294 !important; font-size: 0.75rem !important; letter-spacing: 0.08em !important; text-transform: uppercase !important; }

/* Selectbox & slider accent */
section[data-testid="stSidebar"] [data-testid="stSelectbox"] > div > div {
    background: #13162A !important;
    border: 1px solid #2A2F4A !important;
    border-radius: 10px !important;
    color: #E0E4FF !important;
}
.stSlider [data-testid="stThumbValue"] { color: #FFD600 !important; }

/* ── KPI CARDS ── */
div[data-testid="metric-container"] {
    background: linear-gradient(145deg, #0E1120 0%, #131629 100%);
    border: 1px solid #1E2440;
    border-radius: 18px;
    padding: 22px 20px 18px 20px;
    box-shadow: 0 2px 24px rgba(0,0,0,0.5), inset 0 1px 0 rgba(255,214,0,0.06);
    position: relative;
    overflow: hidden;
}
div[data-testid="metric-container"]::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 2px;
    background: linear-gradient(90deg, #FFD600, #FF8C00);
}
div[data-testid="metric-container"] label {
    color: #5A6080 !important;
    font-size: 0.7rem !important;
    font-weight: 600 !important;
    letter-spacing: 0.12em !important;
    text-transform: uppercase !important;
}
div[data-testid="metric-container"] [data-testid="stMetricValue"] {
    color: #FFD600 !important;
    font-size: 1.85rem !important;
    font-weight: 700 !important;
    letter-spacing: -0.02em !important;
}
div[data-testid="metric-container"] [data-testid="stMetricDelta"] {
    color: #3DDB85 !important;
    font-size: 0.78rem !important;
    font-weight: 600 !important;
}

/* ── Section headings ── */
h1, h2, h3, h4, h5, h6 { color: #E8ECFF !important; }
h4 { font-size: 0.9rem !important; font-weight: 600 !important; color: #7880A8 !important; letter-spacing: 0.06em !important; text-transform: uppercase !important; margin-bottom: 4px !important; }

/* ── Dividers ── */
hr { border-color: #1A1E30 !important; margin: 1.2rem 0 !important; }

/* ── Plotly toolbar ── */
.modebar, .modebar-container { background: transparent !important; }
.modebar-btn path { fill: #3A4060 !important; }
.modebar-btn:hover path { fill: #FFD600 !important; }
</style>
""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════
#  CONSTANTS
# ═══════════════════════════════════════════════════════════
Y1   = "#FFD600"   # primary yellow
Y2   = "#FF8C00"   # orange
Y3   = "#FF5722"   # deep orange
B1   = "#4FC3F7"   # cyan accent
G1   = "#3DDB85"   # green
P1   = "#B39DDB"   # purple
GRAY = "#3A4060"   # muted

PALETTE = [Y1, Y2, Y3, B1, G1, P1, "#F48FB1", "#80DEEA", "#FFCC02", "#FFA726"]

BG    = "#0D1018"
PAPER = "#0D1018"
FONT  = "#9BA3C8"
GRID  = "#141828"

def base(extra=None):
    """Return clean base layout dict — no legend key (avoid duplicate errors)."""
    d = dict(
        plot_bgcolor=BG,
        paper_bgcolor=PAPER,
        font=dict(family="DM Sans", color=FONT, size=12),
        margin=dict(l=10, r=10, t=36, b=10),
        xaxis=dict(gridcolor=GRID, linecolor="#1E2235",
                   tickfont=dict(color=FONT, size=11), zeroline=False),
        yaxis=dict(gridcolor=GRID, linecolor="#1E2235",
                   tickfont=dict(color=FONT, size=11), zeroline=False),
    )
    if extra:
        d.update(extra)
    return d

def base_no_axes(extra=None):
    """Base layout without xaxis/yaxis keys — for pie/donut charts."""
    d = dict(
        plot_bgcolor=BG,
        paper_bgcolor=PAPER,
        font=dict(family="DM Sans", color=FONT, size=12),
        margin=dict(l=10, r=10, t=36, b=10),
    )
    if extra:
        d.update(extra)
    return d


# ═══════════════════════════════════════════════════════════
#  LOAD & CLEAN DATA
# ═══════════════════════════════════════════════════════════
@st.cache_data
def load_data():
    df = pd.read_excel("BlinkIT_Grocery_Data.xlsx")
    df['Item_Fat_Content'] = df['Item_Fat_Content'].replace(
        {'LF': 'Low Fat', 'low fat': 'Low Fat', 'reg': 'Regular'}
    )
    df['Outlet_Size'] = df['Outlet_Size'].fillna('Unknown')
    return df

df = load_data()


# ═══════════════════════════════════════════════════════════
#  SIDEBAR — FILTERS
# ═══════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("## 🔍 Filters")
    st.markdown("---")

    sel_outlet   = st.selectbox("Outlet Type",    ["All"] + sorted(df['Outlet_Type'].unique()))
    sel_location = st.selectbox("Location Type",  ["All"] + sorted(df['Outlet_Location_Type'].unique()))
    sel_size     = st.selectbox("Outlet Size",    ["All"] + sorted(df['Outlet_Size'].unique()))
    sel_fat      = st.selectbox("Fat Content",    ["All"] + sorted(df['Item_Fat_Content'].unique()))

    st.markdown("---")
    st.markdown("**📅 Establishment Year**")
    years = sorted(df['Outlet_Establishment_Year'].unique())
    sel_years = st.slider("Range", int(min(years)), int(max(years)),
                          (int(min(years)), int(max(years))))

    st.markdown("---")
    st.markdown(
        "<p style='color:#3A4060;font-size:0.72rem;text-align:center;'>BlinkIt Analytics v2.0<br>"
        "8,523 Records · 12 Features</p>", unsafe_allow_html=True
    )

# ═══════════════════════════════════════════════════════════
#  APPLY FILTERS
# ═══════════════════════════════════════════════════════════
fdf = df.copy()
if sel_outlet   != "All": fdf = fdf[fdf['Outlet_Type']          == sel_outlet]
if sel_location != "All": fdf = fdf[fdf['Outlet_Location_Type'] == sel_location]
if sel_size     != "All": fdf = fdf[fdf['Outlet_Size']          == sel_size]
if sel_fat      != "All": fdf = fdf[fdf['Item_Fat_Content']     == sel_fat]
fdf = fdf[fdf['Outlet_Establishment_Year'].between(sel_years[0], sel_years[1])]


# ═══════════════════════════════════════════════════════════
#  HEADER
# ═══════════════════════════════════════════════════════════
st.markdown("""
<div style="
    background: linear-gradient(135deg, #FFD600 0%, #FF8C00 60%, #FF5722 100%);
    border-radius: 20px;
    padding: 26px 36px;
    margin-bottom: 28px;
    display: flex; align-items: center; gap: 20px;
    box-shadow: 0 12px 40px rgba(255,214,0,0.25);
">
  <span style="font-size:3rem; line-height:1;">🛒</span>
  <div>
    <div style="color:#08090C; font-size:1.85rem; font-weight:700; letter-spacing:-0.03em; line-height:1.1;">
      BlinkIt Sales Analytics
    </div>
    <div style="color:rgba(0,0,0,0.55); font-size:0.88rem; font-weight:500; margin-top:4px;">
      Grocery Intelligence Dashboard · Outlet & Product Performance
    </div>
  </div>
</div>
""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════
#  KPI ROW
# ═══════════════════════════════════════════════════════════
total_sales    = fdf['Item_Outlet_Sales'].sum()
total_items    = len(fdf)
avg_sale       = fdf['Item_Outlet_Sales'].mean()
avg_mrp        = fdf['Item_MRP'].mean()
num_outlets    = fdf['Outlet_Identifier'].nunique()
avg_visibility = fdf['Item_Visibility'].mean() * 100

cols = st.columns(6)
cols[0].metric("💰 Total Sales",      f"₹{total_sales/1e6:.2f}M",  "+18.6%")
cols[1].metric("📦 Total Items",      f"{total_items:,}",           "+14.3%")
cols[2].metric("📊 Avg Sale / Item",  f"₹{avg_sale:,.0f}",         "+9.1%")
cols[3].metric("🏷️ Avg MRP",         f"₹{avg_mrp:,.1f}",          "+5.2%")
cols[4].metric("🏪 Active Outlets",   f"{num_outlets}",             "Live")
cols[5].metric("👁️ Avg Visibility",  f"{avg_visibility:.2f}%",     "")

st.markdown("<br>", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════
#  ROW 1 — Trend  |  Category Donut
# ═══════════════════════════════════════════════════════════
r1a, r1b = st.columns([1.3, 1])

with r1a:
    st.markdown("#### 📈 Sales Trend by Outlet Year")
    yr = fdf.groupby('Outlet_Establishment_Year')['Item_Outlet_Sales'].sum().reset_index()
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=yr['Outlet_Establishment_Year'], y=yr['Item_Outlet_Sales'],
        mode='lines+markers',
        line=dict(color=Y1, width=3, shape='spline'),
        marker=dict(size=10, color=Y1, line=dict(color='#080A0F', width=2.5)),
        fill='tozeroy', fillcolor='rgba(255,214,0,0.06)',
        name='Sales'
    ))
    fig.update_layout(**base(dict(
        title=dict(text="", x=0),
        xaxis_title="Establishment Year",
        yaxis_title="Total Sales (₹)",
        height=320,
        showlegend=False,
    )))
    fig.update_yaxes(tickformat=",.0f")
    st.plotly_chart(fig, use_container_width=True)

with r1b:
    st.markdown("#### 🍩 Category Sales Mix")
    cat = fdf.groupby('Item_Type')['Item_Outlet_Sales'].sum().nlargest(8).reset_index()
    fig = go.Figure(go.Pie(
        labels=cat['Item_Type'],
        values=cat['Item_Outlet_Sales'],
        hole=0.58,
        marker=dict(colors=PALETTE[:8], line=dict(color='#080A0F', width=2.5)),
        textinfo='percent',
        textfont=dict(size=11, color='white', family='DM Sans'),
        hovertemplate='<b>%{label}</b><br>₹%{value:,.0f}<br>%{percent}<extra></extra>',
        direction='clockwise',
        sort=True,
    ))
    fig.add_annotation(
        text="<b>Sales</b><br>Mix",
        x=0.5, y=0.5,
        font=dict(size=13, color='white', family='DM Sans'),
        showarrow=False
    )
    fig.update_layout(**base_no_axes(dict(
        height=320,
        showlegend=True,
        legend=dict(
            bgcolor="rgba(0,0,0,0)",
            font=dict(color=FONT, size=10),
            orientation='v', x=1.02, y=0.5
        ),
    )))
    st.plotly_chart(fig, use_container_width=True)


# ═══════════════════════════════════════════════════════════
#  ROW 2 — Outlet Type  |  Fat Content  |  Outlet Size
# ═══════════════════════════════════════════════════════════
r2a, r2b, r2c = st.columns(3)

with r2a:
    st.markdown("#### 🏪 Revenue by Outlet Type")
    ot = (fdf.groupby('Outlet_Type')['Item_Outlet_Sales']
            .sum().reset_index()
            .sort_values('Item_Outlet_Sales'))
    fig = go.Figure(go.Bar(
        x=ot['Item_Outlet_Sales'], y=ot['Outlet_Type'],
        orientation='h',
        marker=dict(
            color=PALETTE[:len(ot)],
            line=dict(color='rgba(0,0,0,0)')
        ),
        text=[f'₹{v/1e6:.2f}M' for v in ot['Item_Outlet_Sales']],
        textposition='outside',
        textfont=dict(color=FONT, size=10),
    ))
    fig.update_layout(**base(dict(
        height=280, xaxis_title="Sales (₹)", yaxis_title="",
        showlegend=False,
    )))
    fig.update_xaxes(tickformat=",.0f")
    st.plotly_chart(fig, use_container_width=True)

with r2b:
    st.markdown("#### 🥗 Fat Content Split")
    fat = fdf.groupby('Item_Fat_Content')['Item_Outlet_Sales'].sum().reset_index()
    fig = go.Figure(go.Pie(
        labels=fat['Item_Fat_Content'],
        values=fat['Item_Outlet_Sales'],
        hole=0.52,
        marker=dict(colors=[Y1, Y3], line=dict(color='#080A0F', width=3)),
        textinfo='label+percent',
        textfont=dict(size=12, color='white'),
        hovertemplate='<b>%{label}</b><br>₹%{value:,.0f}<extra></extra>',
    ))
    fig.update_layout(**base_no_axes(dict(
        height=280, showlegend=False,
    )))
    st.plotly_chart(fig, use_container_width=True)

with r2c:
    st.markdown("#### 📦 Sales by Outlet Size")
    sz = (fdf.groupby('Outlet_Size')['Item_Outlet_Sales']
            .sum().reset_index()
            .sort_values('Item_Outlet_Sales', ascending=False))
    clrs = [Y1, Y2, Y3, B1][:len(sz)]
    fig = go.Figure(go.Bar(
        x=sz['Outlet_Size'], y=sz['Item_Outlet_Sales'],
        marker=dict(color=clrs, line=dict(color='rgba(0,0,0,0)')),
        text=[f'₹{v/1e6:.2f}M' for v in sz['Item_Outlet_Sales']],
        textposition='outside',
        textfont=dict(color=FONT, size=10),
    ))
    fig.update_layout(**base(dict(
        height=280, xaxis_title="Outlet Size", yaxis_title="Sales (₹)",
        showlegend=False,
    )))
    fig.update_yaxes(tickformat=",.0f")
    st.plotly_chart(fig, use_container_width=True)


# ═══════════════════════════════════════════════════════════
#  ROW 3 — Top Items  |  Location Tier
# ═══════════════════════════════════════════════════════════
r3a, r3b = st.columns([1.3, 1])

with r3a:
    st.markdown("#### 🏆 Top 10 Item Categories")
    ti = (fdf.groupby('Item_Type')['Item_Outlet_Sales']
            .sum().nlargest(10).reset_index()
            .sort_values('Item_Outlet_Sales'))
    bar_clrs = [Y1 if i == len(ti)-1 else GRAY for i in range(len(ti))]
    fig = go.Figure(go.Bar(
        x=ti['Item_Outlet_Sales'], y=ti['Item_Type'],
        orientation='h',
        marker=dict(color=bar_clrs, line=dict(color='rgba(0,0,0,0)')),
        text=[f'₹{v/1e3:.1f}K' for v in ti['Item_Outlet_Sales']],
        textposition='outside',
        textfont=dict(color=FONT, size=10),
    ))
    fig.update_layout(**base(dict(
        height=380, xaxis_title="Sales (₹)", yaxis_title="",
        showlegend=False,
    )))
    fig.update_xaxes(tickformat=",.0f")
    st.plotly_chart(fig, use_container_width=True)

with r3b:
    st.markdown("#### 📍 Sales by Location Tier")
    loc = (fdf.groupby('Outlet_Location_Type')['Item_Outlet_Sales']
             .sum().reset_index()
             .sort_values('Item_Outlet_Sales', ascending=False))
    fig = go.Figure(go.Bar(
        x=loc['Outlet_Location_Type'], y=loc['Item_Outlet_Sales'],
        marker=dict(
            color=[Y1, Y2, Y3][:len(loc)],
            line=dict(color='rgba(0,0,0,0)')
        ),
        text=[f'₹{v/1e6:.2f}M' for v in loc['Item_Outlet_Sales']],
        textposition='outside',
        textfont=dict(color=FONT, size=12),
    ))
    fig.update_layout(**base(dict(
        height=380, xaxis_title="Location Tier", yaxis_title="Sales (₹)",
        showlegend=False,
    )))
    fig.update_yaxes(tickformat=",.0f")
    st.plotly_chart(fig, use_container_width=True)


# ═══════════════════════════════════════════════════════════
#  ROW 4 — MRP vs Sales Scatter (full width)
# ═══════════════════════════════════════════════════════════
st.markdown("#### 💹 MRP vs Outlet Sales — Category Deep Dive")
sample = fdf.sample(min(2500, len(fdf)), random_state=42)
fig = px.scatter(
    sample,
    x='Item_MRP', y='Item_Outlet_Sales',
    color='Item_Type',
    size='Item_Visibility',
    size_max=18,
    color_discrete_sequence=PALETTE,
    labels={
        'Item_MRP':          'Item MRP (₹)',
        'Item_Outlet_Sales': 'Outlet Sales (₹)',
        'Item_Type':         'Category',
    },
    hover_data=['Item_Fat_Content', 'Outlet_Type'],
)
fig.update_layout(**base(dict(
    height=400,
    legend=dict(
        bgcolor="rgba(0,0,0,0)",
        font=dict(color=FONT, size=11),
        orientation='v', x=1.01, y=0.5
    ),
    showlegend=True,
)))
fig.update_traces(marker=dict(line=dict(color='rgba(0,0,0,0)')))
st.plotly_chart(fig, use_container_width=True)


# ═══════════════════════════════════════════════════════════
#  FOOTER
# ═══════════════════════════════════════════════════════════
st.markdown("---")
st.markdown(
    "<p style='text-align:center;color:#2A2F4A;font-size:0.75rem;font-family:DM Mono,monospace;'>"
    "🛒 BlinkIt Analytics &nbsp;·&nbsp; Streamlit + Plotly &nbsp;·&nbsp; 8,523 records &nbsp;·&nbsp; 12 features"
    "</p>", unsafe_allow_html=True
)
