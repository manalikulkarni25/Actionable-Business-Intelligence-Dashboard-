# ==========================================================
# MS-CIT Admissions — Standalone HTML with Slicers (Panel)
# Saves a single HTML to your Desktop and opens it.
# Run: python C:\Users\manalik\Desktop\Dashboard.py
# ==========================================================
import os, re, sys, traceback, pathlib, webbrowser
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.express as px
import panel as pn

# -------------------------
# SETTINGS
# -------------------------
FILE_PATH = r"C:\Users\manalik\Desktop\Admissions_Data.xlsx"  # can be .xlsx or .csv
DEFAULT_REGION_FALLBACK = "District"  # used if no RLC column exists

# -------------------------
# INIT PANEL
# -------------------------
pn.extension("plotly")

# -------------------------
# Load file (.xlsx or .csv)
# -------------------------
def load_table(path: str) -> pd.DataFrame:
    p = str(path).lower()
    if p.endswith(".csv"):
        # adjust encoding/delimiter if needed
        return pd.read_csv(path, encoding="utf-8", engine="python")
    return pd.read_excel(path, engine="openpyxl")

df = load_table(FILE_PATH)

# -------------------------
# Clean & normalize columns
# -------------------------
df.columns = (
    df.columns.astype(str)
      .str.strip()
      .str.replace("\xa0", " ", regex=False)
      .str.replace(r"\s+", " ", regex=True)
)

rename_map = {
    "RLC Region (RLC Name)": "RLC_Name",
    "LLC Region(LLC Name)": "LLC_Name",
    "Center Name": "ALC_Name",
    "LearnerDistrict": "District",
    "Tp Name": "Training_Partner",
}
df.rename(columns=rename_map, inplace=True)

for col in ["District", "RLC_Name", "LLC_Name", "ALC_Name", "Batch", "Gender"]:
    if col not in df.columns:
        df[col] = "Unknown"
    df[col] = df[col].fillna("Unknown")

# -------------------------
# Derive Year/Month from Batch
# -------------------------
MONTH_MAP = {
    "jan":1,"january":1,"feb":2,"february":2,"mar":3,"march":3,"apr":4,"april":4,
    "may":5,"jun":6,"june":6,"jul":7,"july":7,"aug":8,"august":8,"sep":9,"sept":9,
    "september":9,"oct":10,"october":10,"nov":11,"november":11,"dec":12,"december":12,
}

def extract_year(text):
    if isinstance(text, str):
        m = re.search(r"(20\d{2})", text)
        if m: return int(m.group(1))
    return np.nan

def extract_month(text):
    if not isinstance(text, str): return np.nan
    s = text.lower()
    for k,v in MONTH_MAP.items():
        if k in s: return v
    m = re.search(r"[^0-9](1[0-2]|0?[1-9])([^0-9]|$)", s)
    if m: return int(m.group(1))
    return np.nan

current_year = datetime.now().year
df["Year"]  = df["Batch"].apply(extract_year).fillna(current_year).astype(int)
df["Month"] = df["Batch"].apply(extract_month).fillna(1).astype(int)
df["Actual_Admissions"] = 1  # each row = one admission

# -------------------------
# Region column detection
# -------------------------
def pick_region_column(frame: pd.DataFrame) -> str:
    for c in ["RLC_Name", "RLC Region (RLC Name)", "RLC Region"]:
        if c in frame.columns:
            return c
    return DEFAULT_REGION_FALLBACK

REGION_COL = pick_region_column(df)

# -------------------------
# Target logic helpers
# -------------------------
def compute_region_targets(frame: pd.DataFrame, year: int, region_col: str) -> pd.DataFrame:
    prev_years = [y for y in sorted(frame["Year"].unique()) if y < year]
    prev2 = prev_years[-2:]  # last two historical years
    hist = (
        frame[frame["Year"].isin(prev2)]
        .groupby(region_col, as_index=False)["Actual_Admissions"].sum()
        .rename(columns={"Actual_Admissions": "Target"})
    )
    if hist.empty:
        # fallback: proportional split of current distribution (or zeros)
        cur = frame[frame["Year"] == year].groupby(region_col, as_index=False)["Actual_Admissions"].sum()
        total_cur = cur["Actual_Admissions"].sum()
        if total_cur > 0:
            cur["Target"] = total_cur
        else:
            regions = frame[region_col].unique().tolist()
            cur = pd.DataFrame({region_col: regions, "Target": 0})
        hist = cur[[region_col, "Target"]]
    return hist

# -------------------------
# Widgets (slicers) — work in saved HTML too
# -------------------------
years = sorted(df["Year"].unique())
year_sel   = pn.widgets.Select(name="Year", value=max(years), options=years)
regions = ["All"] + sorted(df[REGION_COL].dropna().unique().tolist())
region_sel = pn.widgets.Select(name=REGION_COL, value="All", options=regions)

# -------------------------
# Reactive blocks
# -------------------------
@pn.depends(year_sel, region_sel)
def kpis_and_insights(year, region_choice):
    sub = df[df["Year"] == year].copy()
    if region_choice != "All":
        sub = sub[sub[REGION_COL] == region_choice]

    targets = compute_region_targets(df, year, REGION_COL)
    actuals = sub.groupby(REGION_COL, as_index=False)["Actual_Admissions"].sum().rename(columns={"Actual_Admissions":"Actual"})
    perf = pd.merge(actuals, targets, on=REGION_COL, how="outer").fillna(0)
    perf["Achievement_%"] = np.where(perf["Target"]>0, 100*perf["Actual"]/perf["Target"], np.nan)
    perf["Shortfall_%"]   = np.where(perf["Target"]>0, 100*(perf["Target"]-perf["Actual"])/perf["Target"], np.nan)
    perf["Status"]        = np.where(perf["Actual"]>=perf["Target"], "Achieved", "Shortfall")

    actual_total = int(sub["Actual_Admissions"].sum())
    scope_regions = perf[REGION_COL].unique()
    target_total = float(targets[targets[REGION_COL].isin(scope_regions)]["Target"].sum())
    pct = (100*actual_total/target_total) if target_total>0 else np.nan

    # KPI cards
    kpi_target = pn.indicators.Number(name="🎯 Target (scope)", value=target_total, format="{:,.0f}")
    kpi_actual = pn.indicators.Number(name="🧮 Actual", value=actual_total, format="{:,.0f}")
    kpi_pct    = pn.indicators.Number(name="📈 % Achieved", value=(pct if not np.isnan(pct) else 0), format="{:.1f}%")
    prog = pn.indicators.Progress(name="Target Completion", value=int(min(max(pct,0),100)) if not np.isnan(pct) else 0)

    # Insights
    insights = []
    insights.append(f"Overall progress: <b>{pct:.1f}%</b> of target (Year {year})." if target_total>0 else "Target unavailable due to insufficient history.")
    shortfall = perf[perf["Status"]=="Shortfall"].sort_values("Shortfall_%", ascending=False)
    if not shortfall.empty:
        top3 = ", ".join(shortfall[REGION_COL].head(3).tolist())
        insights.append(f"Focus regions (Top shortfalls): <b>{top3}</b>.")
    if "Gender" in sub.columns and sub["Actual_Admissions"].sum()>0:
        g = sub.groupby("Gender")["Actual_Admissions"].sum().sort_values(ascending=False)
        if len(g)>1:
            lowest = g.index[-1]
            share = 100*g.iloc[-1]/g.sum()
            if share < 35:
                insights.append(f"Low participation among <b>{lowest}</b> learners (~{share:.1f}%). Plan targeted outreach.")

    card = pn.pane.HTML(
        "<ul style='margin:0;padding-left:18px'>" + "".join([f"<li>{x}</li>" for x in insights]) + "</ul>",
        styles={"background":"#f4f6ff","border":"1px solid #dfe3ff","border-radius":"8px","padding":"12px","color":"#1e1e1e"}
    )

    return pn.Column(
        pn.Row(kpi_target, kpi_actual, kpi_pct, sizing_mode="stretch_width"),
        prog,
        pn.pane.Markdown("### Insights"),
        card,
    )

@pn.depends(year_sel, region_sel)
def charts(year, region_choice):
    sub = df[df["Year"] == year].copy()
    if region_choice != "All":
        sub = sub[sub[REGION_COL] == region_choice]

    targets = compute_region_targets(df, year, REGION_COL)
    actuals = sub.groupby(REGION_COL, as_index=False)["Actual_Admissions"].sum().rename(columns={"Actual_Admissions":"Actual"})
    perf = pd.merge(actuals, targets, on=REGION_COL, how="outer").fillna(0)

    fig_bar = px.bar(
        perf.melt(id_vars=[REGION_COL], value_vars=["Actual","Target"], var_name="Metric", value_name="Count"),
        x=REGION_COL, y="Count", color="Metric", barmode="group",
        title=f"Target vs Actual by {REGION_COL} (Year {year})"
    )

    if "Gender" in sub.columns:
        fig_gender = px.pie(sub, names="Gender", title="Gender-wise Distribution")
    else:
        fig_gender = px.pie(values=[1], names=["Data Not Available"], title="Gender-wise Distribution")

    if sub["Month"].nunique() > 1:
        monthly = sub.groupby("Month", as_index=False)["Actual_Admissions"].sum()
        fig_trend = px.line(monthly, x="Month", y="Actual_Admissions", markers=True, title=f"Monthly Admissions (Year {year})")
    else:
        year_total = sub["Actual_Admissions"].sum()
        fig_trend = px.bar(pd.DataFrame({"Year":[year],"Actual_Admissions":[year_total]}),
                           x="Year", y="Actual_Admissions", title="Year-wise Admissions")

    return pn.Row(
        pn.pane.Plotly(fig_bar, config={"displaylogo": False}),
        pn.Spacer(width=20),
        pn.Column(pn.pane.Plotly(fig_gender, config={"displaylogo": False}),
                  pn.pane.Plotly(fig_trend, config={"displaylogo": False}))
    )

@pn.depends(year_sel, region_sel)
def tables(year, region_choice):
    sub = df[df["Year"] == year].copy()
    if region_choice != "All":
        sub = sub[sub[REGION_COL] == region_choice]

    targets = compute_region_targets(df, year, REGION_COL)
    actuals = sub.groupby(REGION_COL, as_index=False)["Actual_Admissions"].sum().rename(columns={"Actual_Admissions":"Actual"})
    perf = pd.merge(actuals, targets, on=REGION_COL, how="outer").fillna(0)
    perf["Achievement_%"] = np.where(perf["Target"]>0, 100*perf["Actual"]/perf["Target"], np.nan)
    perf["Shortfall_%"]   = np.where(perf["Target"]>0, 100*(perf["Target"]-perf["Actual"])/perf["Target"], np.nan)
    perf["Status"]        = np.where(perf["Actual"]>=perf["Target"], "Achieved", "Shortfall")

    def _status_color(v): return "background-color: #d1ffd1" if v=="Achieved" else "background-color: #ffd6d6"
    styled = (
        perf[[REGION_COL,"Actual","Target","Achievement_%","Shortfall_%","Status"]]
        .rename(columns={REGION_COL:"Region"})
        .style.format({"Actual":"{:,.0f}","Target":"{:,.0f}","Achievement_%":"{:,.1f}%","Shortfall_%":"{:,.1f}%"})
        .apply(lambda s: [_status_color(v) for v in s] if s.name=="Status" else [""]*len(s))
    )
    table_perf = pn.pane.HTML(styled.to_html(), sizing_mode="stretch_width")

    shortfall = perf[perf["Status"]=="Shortfall"].sort_values("Shortfall_%", ascending=False)
    if shortfall.empty:
        short_html = "<i>No regions below target.</i>"
    else:
        short_html = shortfall[[REGION_COL,"Actual","Target","Shortfall_%"]] \
            .rename(columns={REGION_COL:"Region"}) \
            .style.format({"Actual":"{:,.0f}","Target":"{:,.0f}","Shortfall_%":"{:,.1f}%"}) \
            .to_html()
    table_short = pn.pane.HTML(short_html, sizing_mode="stretch_width")

    return pn.Column(
        pn.pane.Markdown("### Region Performance (Conditional Formatting)"),
        table_perf,
        pn.pane.Markdown("### Shortfall Table (Top Laggards)"),
        table_short
    )

# -------------------------
# Template (layout + sidebar)
# -------------------------
tmpl = pn.template.FastListTemplate(
    site="MS-CIT",
    title="Admissions Dashboard — Program Manager View",
    sidebar=[pn.pane.Markdown("### Filters"), year_sel, region_sel],
    main=[kpis_and_insights, pn.layout.HSpacer(height=10), charts, pn.layout.HSpacer(height=10), tables],
    theme="default",
)
tmpl.config.raw_css.append("""
/* Improve contrast for insights card */
.bk.panel-models-pane-HTML { font-size: 14px; }
""")

# -------------------------
# SAVE (no embedding on templates)
# -------------------------
try:
    OUTPUT = pathlib.Path.home() / "Desktop" / "Admissions_Dashboard_Standalone.html"
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)

    # ✅ Correct for Templates: NO embed=True here
    pn.save(tmpl, OUTPUT, resources="inline")

    print("\n✅ Saved dashboard:")
    print(str(OUTPUT))

    # Try to open automatically
    opened = False
    try:
        opened = webbrowser.open_new_tab(OUTPUT.as_uri())
    except Exception:
        pass
    if not opened and sys.platform.startswith("win"):
        try:
            os.startfile(str(OUTPUT))
            opened = True
        except Exception:
            pass
    if not opened:
        print("ℹ️ Could not auto-open. Please open the file from your Desktop.")

except Exception:
    print("\n❌ Failed to build dashboard. Full error:\n")
    print(traceback.format_exc())
    raise