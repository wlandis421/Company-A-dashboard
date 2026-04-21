"""
Company A — Value Creation Dashboard
Streamlit app — reads CompanyA_Value_Tracking_v2.xlsx
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
import os

st.set_page_config(
    page_title="Company A | Value Creation Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Styling ────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main { background-color: #f8f9fa; }
    .stMetric { background: white; border-radius: 8px; padding: 12px; box-shadow: 0 1px 4px rgba(0,0,0,0.08); }
    h1, h2, h3 { color: #1F3864; }
    .kpi-card { background: white; border-radius: 10px; padding: 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); text-align: center; }
    .kpi-value { font-size: 2rem; font-weight: 700; color: #1F3864; }
    .kpi-label { font-size: 0.85rem; color: #595959; margin-top: 4px; }
    .section-header { background: #1F3864; color: white; padding: 10px 18px; border-radius: 6px; margin: 18px 0 10px 0; font-weight: 600; }
    div[data-testid="stMetricValue"] { font-size: 1.6rem !important; color: #1F3864 !important; }
    div[data-testid="stMetricLabel"] { color: #595959 !important; font-size: 0.8rem !important; }
</style>
""", unsafe_allow_html=True)

# ── Load data ──────────────────────────────────────────────────────────────────
EXCEL_FILE = "CompanyA_Value_Tracking_v2.xlsx"

@st.cache_data(ttl=60)
def load_data(filepath):
    xl = pd.ExcelFile(filepath)
    log = pd.read_excel(xl, sheet_name="Initiative Log", header=1, nrows=25)
    log.columns = [str(c).strip().replace("\n"," ") for c in log.columns]

    # Rename to clean keys
    col_map = {
        log.columns[0]:  "ID",
        log.columns[1]:  "BU",
        log.columns[2]:  "Segment",
        log.columns[3]:  "Category",
        log.columns[4]:  "Topic",
        log.columns[5]:  "SubCat",
        log.columns[6]:  "Name",
        log.columns[7]:  "Owner",
        log.columns[8]:  "Dept",
        log.columns[9]:  "GL",
        log.columns[10]: "Description",
        log.columns[11]: "Status",
        log.columns[12]: "StartDate",
        log.columns[13]: "EndDate",
        log.columns[14]: "Baseline",
        log.columns[15]: "GrossSave",
        log.columns[16]: "OneTime",
        log.columns[17]: "UpsideTarget",
        log.columns[18]: "Yr1Pct",
        log.columns[19]: "Yr2Pct",
        log.columns[20]: "Yr3Pct",
        log.columns[21]: "PnLLine",
        log.columns[22]: "Notes",
    }
    log = log.rename(columns=col_map)
    log = log[log["ID"].notna() & log["ID"].astype(str).str.startswith("I-")]
    log["GrossSave"]    = pd.to_numeric(log["GrossSave"],    errors="coerce").fillna(0)
    log["OneTime"]      = pd.to_numeric(log["OneTime"],      errors="coerce").fillna(0)
    log["UpsideTarget"] = pd.to_numeric(log["UpsideTarget"], errors="coerce").fillna(0)
    log["Baseline"]     = pd.to_numeric(log["Baseline"],     errors="coerce").fillna(0)
    log["Yr1Pct"]       = pd.to_numeric(log["Yr1Pct"],       errors="coerce").fillna(0)
    log["Yr2Pct"]       = pd.to_numeric(log["Yr2Pct"],       errors="coerce").fillna(0)
    log["Yr3Pct"]       = pd.to_numeric(log["Yr3Pct"],       errors="coerce").fillna(0)
    log["NetSave"]  = log["GrossSave"] - log["OneTime"]
    log["Yr1Save"]  = (log["GrossSave"] * log["Yr1Pct"]).astype(int)
    log["Yr2Save"]  = (log["GrossSave"] * log["Yr2Pct"]).astype(int)
    log["Yr3Save"]  = log["GrossSave"].astype(int)
    log["Upside"]   = (log["GrossSave"] * 1.15).astype(int)
    log["Baseline_Scn"] = (log["GrossSave"] * 0.85).astype(int)
    return log

def fmt_m(v):
    """Format as $Xm or $X.Xm"""
    v = v / 1e6
    if abs(v) >= 10:
        return f"${v:.1f}M"
    return f"${v:.2f}M"

def fmt_k(v):
    if abs(v) >= 1e6: return fmt_m(v)
    return f"${v/1e3:.0f}K"

# ── Load ────────────────────────────────────────────────────────────────────────
if not os.path.exists(EXCEL_FILE):
    st.error(f"❌ Excel file not found: `{EXCEL_FILE}`\n\nMake sure the file is in the same folder as this app.")
    st.stop()

df = load_data(EXCEL_FILE)

# ── Sidebar ─────────────────────────────────────────────────────────────────────
st.sidebar.image("https://img.icons8.com/color/96/bar-chart.png", width=60)
st.sidebar.title("Company A\nValue Creation")
st.sidebar.caption("PE Cost Transformation Program")
st.sidebar.markdown("---")

page = st.sidebar.radio("View", [
    "📊 Master Dashboard",
    "🏭 By Workstream",
    "📋 Initiative Tracker",
    "📈 P&L & Valuation",
])

st.sidebar.markdown("---")
st.sidebar.markdown("**Filters**")
seg_filter  = st.sidebar.multiselect("Segment", sorted(df["Segment"].dropna().unique()), default=sorted(df["Segment"].dropna().unique()))
stat_filter = st.sidebar.multiselect("Status",  sorted(df["Status"].dropna().unique()),  default=sorted(df["Status"].dropna().unique()))
topic_filter= st.sidebar.multiselect("Topic",   sorted(df["Topic"].dropna().unique()),   default=sorted(df["Topic"].dropna().unique()))

# Apply filters
mask = (df["Segment"].isin(seg_filter)) & (df["Status"].isin(stat_filter)) & (df["Topic"].isin(topic_filter))
fd = df[mask].copy()

st.sidebar.markdown("---")
st.sidebar.caption(f"Showing **{len(fd)}** of **{len(df)}** initiatives")
st.sidebar.caption("Upload updated Excel to refresh.")

# ── Color maps ─────────────────────────────────────────────────────────────────
STATUS_COLORS = {"Active":"#375623","Pipeline":"#7F6000","Complete":"#1F3864","Cancelled":"#C00000"}
SEG_COLORS    = {"Dx":"#1F3864","BH":"#2E75B6","GY":"#7030A0"}
TOPIC_COLORS  = {"Manufacturing":"#C55A11","G&A / HC":"#375623","Procurement":"#7030A0",
                 "Marketing":"#2E75B6","Technology":"#1F6B75","Distribution":"#C00000","Public-to-Private":"#1F3864"}

NAVY = "#1F3864"

# ══════════════════════════════════════════════════════════════════════════════
# PAGE 1: MASTER DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════
if page == "📊 Master Dashboard":
    st.markdown("# 📊 Master Value Creation Dashboard")
    st.caption("Company A  ·  PE Cost Transformation Program  ·  All figures USD")

    # ── KPI Row ────────────────────────────────────────────────────────────────
    k1,k2,k3,k4,k5,k6 = st.columns(6)
    total_gross = fd["GrossSave"].sum()
    total_ot    = fd["OneTime"].sum()
    total_net   = fd["NetSave"].sum()
    total_yr1   = fd["Yr1Save"].sum()
    hc_gross    = fd[fd["Category"]=="HC"]["GrossSave"].sum()
    nhc_gross   = fd[fd["Category"]=="Non-HC"]["GrossSave"].sum()
    total_ftes  = len(df[df["Category"]=="HC"])  # proxy — 1 initiative ~1 FTE group

    k1.metric("Gross Annual Savings", fmt_m(total_gross))
    k2.metric("HC Savings (SG&A)",    fmt_m(hc_gross))
    k3.metric("Non-HC Savings",        fmt_m(nhc_gross))
    k4.metric("One-Time Costs",        fmt_m(total_ot),  delta=f"-{fmt_m(total_ot)} invested", delta_color="inverse")
    k5.metric("Net Annual Savings",    fmt_m(total_net))
    k6.metric("Year 1 Realized",       fmt_m(total_yr1))

    st.markdown("---")

    col_left, col_right = st.columns([1.1, 0.9])

    with col_left:
        # Savings by Segment bar
        seg_df = fd.groupby("Segment").agg(
            HC=("GrossSave", lambda x: x[fd.loc[x.index,"Category"]=="HC"].sum()),
            NonHC=("GrossSave", lambda x: x[fd.loc[x.index,"Category"]=="Non-HC"].sum()),
        ).reset_index()
        seg_melt = seg_df.melt(id_vars="Segment", var_name="Type", value_name="Savings")

        fig_seg = px.bar(seg_melt, x="Segment", y="Savings", color="Type",
                         title="Gross Savings by Segment (HC vs Non-HC)",
                         color_discrete_map={"HC":"#375623","NonHC":"#2E75B6"},
                         labels={"Savings":"$ Savings","Type":"Category"},
                         text_auto=False, barmode="stack")
        fig_seg.update_traces(texttemplate="")
        fig_seg.update_layout(height=340, plot_bgcolor="white", paper_bgcolor="white",
                              legend=dict(orientation="h",yanchor="bottom",y=1.02),
                              yaxis_tickprefix="$", yaxis_tickformat=",.0f",
                              title_font_color=NAVY)
        st.plotly_chart(fig_seg, use_container_width=True)

        # Savings ramp
        ramp = pd.DataFrame({
            "Year": ["Year 1","Year 2","Year 3 (Run-Rate)"],
            "HC":   [fd[fd["Category"]=="HC"]["Yr1Save"].sum(),
                     fd[fd["Category"]=="HC"]["Yr2Save"].sum(),
                     fd[fd["Category"]=="HC"]["GrossSave"].sum()],
            "Non-HC COGS":[fd[(fd["Category"]=="Non-HC")&(fd["PnLLine"]=="COGS")]["Yr1Save"].sum(),
                           fd[(fd["Category"]=="Non-HC")&(fd["PnLLine"]=="COGS")]["Yr2Save"].sum(),
                           fd[(fd["Category"]=="Non-HC")&(fd["PnLLine"]=="COGS")]["GrossSave"].sum()],
            "Non-HC SGA": [fd[(fd["Category"]=="Non-HC")&(fd["PnLLine"]=="SGA")]["Yr1Save"].sum(),
                           fd[(fd["Category"]=="Non-HC")&(fd["PnLLine"]=="SGA")]["Yr2Save"].sum(),
                           fd[(fd["Category"]=="Non-HC")&(fd["PnLLine"]=="SGA")]["GrossSave"].sum()],
            "Non-HC R&D": [fd[(fd["Category"]=="Non-HC")&(fd["PnLLine"]=="RD")]["Yr1Save"].sum(),
                           fd[(fd["Category"]=="Non-HC")&(fd["PnLLine"]=="RD")]["Yr2Save"].sum(),
                           fd[(fd["Category"]=="Non-HC")&(fd["PnLLine"]=="RD")]["GrossSave"].sum()],
        })
        ramp_melt = ramp.melt(id_vars="Year", var_name="Category", value_name="Savings")
        fig_ramp = px.bar(ramp_melt, x="Year", y="Savings", color="Category",
                          title="3-Year Savings Ramp by P&L Line",
                          color_discrete_sequence=["#375623","#C55A11","#2E75B6","#7030A0"],
                          barmode="stack")
        fig_ramp.update_layout(height=320, plot_bgcolor="white", paper_bgcolor="white",
                               yaxis_tickprefix="$", yaxis_tickformat=",.0f",
                               legend=dict(orientation="h",yanchor="bottom",y=1.02),
                               title_font_color=NAVY)
        st.plotly_chart(fig_ramp, use_container_width=True)

    with col_right:
        # Status donut
        stat_df = fd.groupby("Status")["GrossSave"].sum().reset_index()
        fig_pie = px.pie(stat_df, names="Status", values="GrossSave",
                         title="Savings by Initiative Status",
                         color="Status",
                         color_discrete_map=STATUS_COLORS,
                         hole=0.55)
        fig_pie.update_traces(textinfo="label+percent", textposition="outside")
        fig_pie.update_layout(height=300, showlegend=False, title_font_color=NAVY,
                              paper_bgcolor="white")
        st.plotly_chart(fig_pie, use_container_width=True)

        # Topic breakdown
        topic_df = fd.groupby("Topic")["GrossSave"].sum().sort_values(ascending=True).reset_index()
        fig_topic = px.bar(topic_df, y="Topic", x="GrossSave",
                           title="Gross Savings by Workstream",
                           orientation="h",
                           color="Topic",
                           color_discrete_map=TOPIC_COLORS,
                           labels={"GrossSave":"Gross Savings ($)","Topic":""})
        fig_topic.update_layout(height=360, plot_bgcolor="white", paper_bgcolor="white",
                                showlegend=False, xaxis_tickprefix="$",
                                xaxis_tickformat=",.0f", title_font_color=NAVY)
        st.plotly_chart(fig_topic, use_container_width=True)

    # Owner table
    st.markdown("### Savings by Owner")
    owner_df = fd.groupby("Owner").agg(
        Initiatives=("ID","count"),
        GrossSavings=("GrossSave","sum"),
        OneTimeCost=("OneTime","sum"),
        NetSavings=("NetSave","sum"),
        Yr1Realized=("Yr1Save","sum"),
    ).reset_index().sort_values("GrossSavings", ascending=False)
    owner_df["GrossSavings"] = owner_df["GrossSavings"].map(fmt_m)
    owner_df["OneTimeCost"]  = owner_df["OneTimeCost"].map(fmt_m)
    owner_df["NetSavings"]   = owner_df["NetSavings"].map(fmt_m)
    owner_df["Yr1Realized"]  = owner_df["Yr1Realized"].map(fmt_m)
    st.dataframe(owner_df, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# PAGE 2: BY WORKSTREAM
# ══════════════════════════════════════════════════════════════════════════════
elif page == "🏭 By Workstream":
    st.markdown("# 🏭 Savings by Workstream")

    topics = sorted(fd["Topic"].dropna().unique())
    sel_topic = st.selectbox("Select Workstream", topics)
    tdf = fd[fd["Topic"]==sel_topic]

    if tdf.empty:
        st.warning("No initiatives match current filters for this workstream.")
    else:
        gross_t  = tdf["GrossSave"].sum()
        ot_t     = tdf["OneTime"].sum()
        net_t    = tdf["NetSave"].sum()
        yr1_t    = tdf["Yr1Save"].sum()
        upside_t = tdf["Upside"].sum()
        base_t   = tdf["Baseline_Scn"].sum()

        c1,c2,c3,c4,c5,c6 = st.columns(6)
        c1.metric("# Initiatives", len(tdf))
        c2.metric("Gross Annual", fmt_m(gross_t))
        c3.metric("One-Time Cost", fmt_m(ot_t))
        c4.metric("Net Annual", fmt_m(net_t))
        c5.metric("Year 1 Realized", fmt_m(yr1_t))
        c6.metric("Upside Scenario", fmt_m(upside_t))

        st.markdown("---")

        col1, col2 = st.columns(2)
        with col1:
            # Scenario waterfall
            fig_scn = go.Figure(go.Bar(
                x=["Downside\n(85%)", "Base Plan\n(100%)", "Upside\n(115%)"],
                y=[base_t, gross_t, upside_t],
                marker_color=["#C00000","#2E75B6","#375623"],
                text=[fmt_m(base_t), fmt_m(gross_t), fmt_m(upside_t)],
                textposition="outside",
            ))
            fig_scn.update_layout(title=f"{sel_topic} — Scenario Analysis",
                                  height=320, plot_bgcolor="white", paper_bgcolor="white",
                                  yaxis_tickprefix="$", yaxis_tickformat=",.0f",
                                  title_font_color=NAVY)
            st.plotly_chart(fig_scn, use_container_width=True)

        with col2:
            # Status breakdown
            s_df = tdf.groupby("Status")["GrossSave"].sum().reset_index()
            fig_s = px.pie(s_df, names="Status", values="GrossSave",
                           color="Status", color_discrete_map=STATUS_COLORS,
                           title=f"{sel_topic} — Status Mix", hole=0.5)
            fig_s.update_layout(height=320, paper_bgcolor="white", title_font_color=NAVY)
            st.plotly_chart(fig_s, use_container_width=True)

        # Initiative detail table
        st.markdown(f"### {sel_topic} — Initiative Detail")
        show_cols = ["ID","Name","Segment","Owner","Status","PnLLine",
                     "GrossSave","OneTime","NetSave","Yr1Save","Upside","Baseline_Scn"]
        disp = tdf[show_cols].copy()
        disp.columns = ["ID","Initiative","Segment","Owner","Status","P&L Line",
                        "Gross Savings","One-Time","Net Savings","Yr1 Realized","Upside ($)","Baseline ($)"]
        for col in ["Gross Savings","One-Time","Net Savings","Yr1 Realized","Upside ($)","Baseline ($)"]:
            disp[col] = disp[col].map(fmt_m)

        def color_status(val):
            colors = {"Active":"background-color:#E2EFDA","Pipeline":"background-color:#FFF2CC",
                      "Complete":"background-color:#DEEAF1","Cancelled":"background-color:#FCE4D6"}
            return colors.get(val,"")

        st.dataframe(disp.style.applymap(color_status, subset=["Status"]),
                     use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# PAGE 3: INITIATIVE TRACKER
# ══════════════════════════════════════════════════════════════════════════════
elif page == "📋 Initiative Tracker":
    st.markdown("# 📋 Initiative Tracker")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Initiatives", len(fd))
    col2.metric("Active", len(fd[fd["Status"]=="Active"]))
    col3.metric("Pipeline", len(fd[fd["Status"]=="Pipeline"]))
    col4.metric("Complete", len(fd[fd["Status"]=="Complete"]))

    st.markdown("---")

    # Status × Topic heatmap
    heat = fd.pivot_table(index="Topic", columns="Status", values="GrossSave",
                          aggfunc="sum", fill_value=0).reset_index()
    heat_melt = fd.groupby(["Topic","Status"])["GrossSave"].sum().reset_index()
    fig_heat = px.density_heatmap(heat_melt, x="Status", y="Topic", z="GrossSave",
                                  title="Savings Heatmap: Topic × Status",
                                  color_continuous_scale="Blues",
                                  labels={"GrossSave":"Gross Savings ($)"})
    fig_heat.update_layout(height=360, paper_bgcolor="white", title_font_color=NAVY)
    st.plotly_chart(fig_heat, use_container_width=True)

    # Full initiative table
    st.markdown("### All Initiatives")
    show = ["ID","Name","Topic","Segment","Owner","Status","PnLLine",
            "GrossSave","OneTime","NetSave","Yr1Save"]
    disp = fd[show].copy()
    disp.columns = ["ID","Initiative","Workstream","Segment","Owner","Status","P&L Line",
                    "Gross Savings ($)","One-Time ($)","Net Savings ($)","Yr1 ($)"]
    for col in ["Gross Savings ($)","One-Time ($)","Net Savings ($)","Yr1 ($)"]:
        disp[col] = disp[col].map(fmt_m)

    def color_status(val):
        colors={"Active":"background-color:#E2EFDA;color:#375623",
                "Pipeline":"background-color:#FFF2CC;color:#7F6000",
                "Complete":"background-color:#DEEAF1;color:#1F3864",
                "Cancelled":"background-color:#FCE4D6;color:#C00000"}
        return colors.get(val,"")

    st.dataframe(disp.style.applymap(color_status, subset=["Status"]),
                 use_container_width=True, hide_index=True, height=600)

# ══════════════════════════════════════════════════════════════════════════════
# PAGE 4: P&L & VALUATION
# ══════════════════════════════════════════════════════════════════════════════
elif page == "📈 P&L & Valuation":
    st.markdown("# 📈 P&L Impact & Valuation")
    st.caption("Finance adjusts baseline P&L from cost transformation savings below.")

    # ── P&L Impact ─────────────────────────────────────────────────────────────
    st.markdown("### P&L Savings by Line Item")

    cogs_save = fd[fd["PnLLine"]=="COGS"]["GrossSave"].sum()
    sga_save  = fd[fd["PnLLine"]=="SGA"]["GrossSave"].sum()
    rd_save   = fd[fd["PnLLine"]=="RD"]["GrossSave"].sum()
    total_save= cogs_save + sga_save + rd_save

    c1,c2,c3,c4 = st.columns(4)
    c1.metric("COGS Savings",    fmt_m(cogs_save), f"{cogs_save/total_save*100:.0f}% of total")
    c2.metric("SG&A Savings",    fmt_m(sga_save),  f"{sga_save/total_save*100:.0f}% of total")
    c3.metric("R&D Savings",     fmt_m(rd_save),   f"{rd_save/total_save*100:.0f}% of total")
    c4.metric("Total Impact",    fmt_m(total_save))

    col1, col2 = st.columns(2)
    with col1:
        # P&L bridge waterfall
        pnl_line_df = pd.DataFrame({
            "Line":   ["COGS Savings","SG&A Savings","R&D Savings"],
            "Amount": [cogs_save, sga_save, rd_save],
            "Color":  ["#C55A11","#2E75B6","#7030A0"],
        })
        fig_bridge = px.bar(pnl_line_df, x="Line", y="Amount", color="Line",
                            title="Run-Rate Savings by P&L Line",
                            color_discrete_sequence=["#C55A11","#2E75B6","#7030A0"],
                            text=[fmt_m(v) for v in [cogs_save,sga_save,rd_save]])
        fig_bridge.update_traces(textposition="outside")
        fig_bridge.update_layout(height=340, plot_bgcolor="white", paper_bgcolor="white",
                                 showlegend=False, yaxis_tickprefix="$",
                                 yaxis_tickformat=",.0f", title_font_color=NAVY)
        st.plotly_chart(fig_bridge, use_container_width=True)

    with col2:
        # EBITDA bridge
        baseline_rev  = 850_000_000
        baseline_cogs = -520_000_000
        baseline_sga  = -185_000_000
        baseline_rd   = -68_000_000
        baseline_ebitda = baseline_rev + baseline_cogs + baseline_sga + baseline_rd

        post_cogs   = baseline_cogs + cogs_save
        post_sga    = baseline_sga  + sga_save
        post_rd     = baseline_rd   + rd_save
        post_ebitda = baseline_rev + post_cogs + post_sga + post_rd

        fig_ebitda = go.Figure(go.Waterfall(
            name="EBITDA Bridge",
            orientation="v",
            measure=["absolute","relative","relative","relative","total"],
            x=["Baseline\nEBITDA","COGS\nSavings","SG&A\nSavings","R&D\nSavings","Post-Transformation\nEBITDA"],
            y=[baseline_ebitda/1e6, cogs_save/1e6, sga_save/1e6, rd_save/1e6, 0],
            connector={"line":{"color":"#BFBFBF"}},
            decreasing={"marker":{"color":"#C00000"}},
            increasing={"marker":{"color":"#375623"}},
            totals={"marker":{"color":"#1F3864"}},
            text=[fmt_m(v) for v in [baseline_ebitda, cogs_save, sga_save, rd_save, post_ebitda]],
            textposition="outside",
        ))
        fig_ebitda.update_layout(title="EBITDA Waterfall Bridge ($M)",
                                 height=340, plot_bgcolor="white", paper_bgcolor="white",
                                 yaxis_title="$M", title_font_color=NAVY)
        st.plotly_chart(fig_ebitda, use_container_width=True)

    # ── Valuation ──────────────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("### Valuation & Value Creation")
    st.caption("Adjust assumptions below to see real-time scenario impact.")

    va1, va2, va3, va4 = st.columns(4)
    entry_ebitda = va1.number_input("Entry EBITDA ($M)", value=85.0, step=1.0)
    entry_mult   = va2.number_input("Entry Multiple (x)", value=12.5, step=0.5)
    net_debt     = va3.number_input("Net Debt ($M)", value=420.0, step=10.0)
    hold_yrs     = va4.number_input("Hold Period (yrs)", value=5, step=1)

    ebitda_uplift = total_save / 1e6
    revised_ebitda = entry_ebitda + ebitda_uplift
    entry_ev    = entry_ebitda * entry_mult
    entry_eq    = entry_ev - net_debt

    st.markdown(f"""
    > **EBITDA Uplift from Cost Savings:** {fmt_m(total_save)} → Revised EBITDA: **${revised_ebitda:.1f}M**
    > (Entry EBITDA ${entry_ebitda:.0f}M + ${ebitda_uplift:.1f}M savings)
    """)

    scenarios = {
        "Downside":  {"exit_ebitda": revised_ebitda * 0.85, "exit_mult": 11.5,  "color": "#C00000"},
        "Base Plan": {"exit_ebitda": revised_ebitda,         "exit_mult": entry_mult, "color": "#2E75B6"},
        "Upside":    {"exit_ebitda": revised_ebitda,         "exit_mult": 14.5,  "color": "#375623"},
    }

    sc1, sc2, sc3 = st.columns(3)
    cols_scn = [sc1, sc2, sc3]
    scn_results = {}
    for i, (scn, params) in enumerate(scenarios.items()):
        exit_ev  = params["exit_ebitda"] * params["exit_mult"]
        exit_eq  = exit_ev - net_debt
        moic     = exit_eq / entry_eq if entry_eq > 0 else 0
        val_crtd = exit_eq - entry_eq
        scn_results[scn] = {"exit_ev":exit_ev,"exit_eq":exit_eq,"moic":moic,"val_created":val_crtd}
        cols_scn[i].markdown(f"""
        <div style="background:white;border-radius:10px;padding:16px;border-left:4px solid {params['color']};box-shadow:0 2px 8px rgba(0,0,0,0.07)">
        <b style="color:{params['color']};font-size:1.1rem">{scn}</b><br>
        <span style="font-size:0.8rem;color:#595959">Exit EBITDA: ${params['exit_ebitda']:.1f}M @ {params['exit_mult']:.1f}x</span><br><br>
        <b>Exit EV:</b> ${exit_ev:.0f}M<br>
        <b>Exit Equity:</b> ${exit_eq:.0f}M<br>
        <b>MOIC:</b> {moic:.2f}x<br>
        <b>Value Created:</b> ${val_crtd:.0f}M
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # MOIC comparison chart
    moic_df = pd.DataFrame({
        "Scenario": list(scn_results.keys()),
        "MOIC":     [v["moic"] for v in scn_results.values()],
        "Color":    [s["color"] for s in scenarios.values()],
    })
    fig_moic = px.bar(moic_df, x="Scenario", y="MOIC", color="Scenario",
                      color_discrete_sequence=[s["color"] for s in scenarios.values()],
                      title="MOIC by Scenario", text=[f"{v:.2f}x" for v in moic_df["MOIC"]],
                      labels={"MOIC":"MOIC (x)"})
    fig_moic.update_traces(textposition="outside")
    fig_moic.add_hline(y=2.0, line_dash="dash", line_color="#7F6000",
                       annotation_text="2.0x Target", annotation_position="right")
    fig_moic.update_layout(height=320, plot_bgcolor="white", paper_bgcolor="white",
                            showlegend=False, title_font_color=NAVY)
    st.plotly_chart(fig_moic, use_container_width=True)

    # Value attribution table
    st.markdown("### Value Attribution")
    base_mult_val = scenarios["Base Plan"]["exit_mult"]
    attr_data = {
        "Value Driver": [
            "Cost Savings → EBITDA Uplift",
            "Multiple Expansion (Base → Upside)",
            "NWC Improvement (est.)",
            "Capex Optimization (est.)",
        ],
        "$ Value ($M)": [
            round(ebitda_uplift * base_mult_val, 1),
            round(entry_ebitda * (14.5 - entry_mult), 1),
            35.0,
            15.0,
        ],
        "Commentary": [
            f"${ebitda_uplift:.1f}M EBITDA uplift × {base_mult_val:.1f}x exit multiple",
            f"Multiple expansion {entry_mult:.1f}x → 14.5x on entry EBITDA",
            "Working capital release: procurement & inventory",
            "Reduced capex from facility consolidation",
        ]
    }
    attr_df = pd.DataFrame(attr_data)
    attr_df.loc[len(attr_df)] = ["**Total Value Created**", round(attr_df["$ Value ($M)"].sum(),1), "Sum of all levers"]
    st.dataframe(attr_df, use_container_width=True, hide_index=True)

# ── Footer ─────────────────────────────────────────────────────────────────────
st.markdown("---")
st.caption("Company A Value Creation Dashboard · Data source: CompanyA_Value_Tracking_v2.xlsx · Built with Streamlit + Claude")
