# demand_control_app.py
import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go

# ============================================================
# SECTION 1: PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Demand Control System",
    page_icon="🏭",
    layout="wide"
)

st.title("🏭 Demand Control System")
st.caption("BGSA440AC Production Planning")

# ============================================================
# SECTION 2: SIDEBAR INPUT [1]
# ============================================================

with st.sidebar:
    st.header("⚙️ Input Parameters")

    # ── Product Selection ──────────────────────────────────
    st.subheader("1️⃣ Product")
    product = st.selectbox(
        "Select Basic Type",
        ["M6226", "M6225"]
    )

    # ── Supply Chain Parameters ────────────────────────────
    st.subheader("2️⃣ Process Flow Parameters [1]")

    with st.expander("🏭 Foundry (UMC) - SIN"):
        foundry_ct      = st.number_input(
            "Cycle Time (days)",
            value=90, min_value=1,
            key="f_ct"
        )
        foundry_transit = st.number_input(
            "Transit Time (days)",
            value=5, min_value=0,
            key="f_tr"
        )
        foundry_yield   = st.slider(
            "Yield %",
            min_value=80.0, max_value=100.0,
            value=100.0, step=0.1,
            key="f_y"
        ) / 100

    with st.expander("⚙️ Bump @ Amkor - Taiwan"):
        bump_ct      = st.number_input(
            "Cycle Time (days)",
            value=14, min_value=1,
            key="b_ct"
        )
        bump_transit = st.number_input(
            "Transit Time (days)",
            value=2, min_value=0,
            key="b_tr"
        )
        bump_yield   = st.slider(
            "Yield %",
            min_value=80.0, max_value=100.0,
            value=99.9, step=0.1,
            key="b_y"
        ) / 100

    with st.expander("🔬 Testing @ Amkor - Taiwan"):
        test_ct      = st.number_input(
            "Cycle Time (days)",
            value=7, min_value=1,
            key="t_ct"
        )
        test_transit = st.number_input(
            "Transit Time (days)",
            value=1, min_value=0,
            key="t_tr"
        )
        test_yield   = st.slider(
            "Yield %",
            min_value=80.0, max_value=100.0,
            value=96.0, step=0.1,
            key="t_y"
        ) / 100

    with st.expander("📦 DPS @ Amkor - Taiwan"):
        dps_ct      = st.number_input(
            "Cycle Time (days)",
            value=9, min_value=1,
            key="d_ct"
        )
        dps_transit = st.number_input(
            "Transit Time (days)",
            value=5, min_value=0,
            key="d_tr"
        )
        dps_yield   = st.slider(
            "Yield %",
            min_value=80.0, max_value=100.0,
            value=98.0, step=0.1,
            key="d_y"
        ) / 100

    # ── Product Attributes ─────────────────────────────────
    st.subheader("3️⃣ Product Attributes [1]")

    cpw = st.number_input(
        "CPW (Chip per Wafer)",
        value=36000 if product == "M6226" else 60000,
        min_value=1000,
        step=1000
    )
    tbase = st.number_input(
        "Tbase (Test Hour / Wafer)",
        value=6.5 if product == "M6226" else 8.5,
        min_value=0.1,
        step=0.5
    )
    weekly_output = st.number_input(
        "Weekly Output / Tester",
        value=22 if product == "M6226" else 20,
        min_value=1
    )
    max_testers = st.number_input(
        "Available Testers [1]",
        value=19,
        min_value=1
    )
    initial_stock = st.number_input(
        "Existing DC Inventories (pcs) [1]",
        value=0,
        min_value=0,
        step=100000
    )

    # ── REACH Target ───────────────────────────────────────
    st.subheader("4️⃣ REACH Target")
    reach_target = st.slider(
        "Target REACH Level",
        min_value=1.0, max_value=8.0,
        value=4.5, step=0.5
    )
    reach_min = st.number_input(
        "Min REACH",
        value=4.0, min_value=0.0
    )
    reach_max = st.number_input(
        "Max REACH",
        value=5.0, min_value=0.0
    )

    # ── Run Button ─────────────────────────────────────────
    st.divider()
    run = st.button(
        "🚀 Run Calculation",
        type="primary",
        use_container_width=True
    )

# ============================================================
# SECTION 3: DEMAND INPUT (Main Area)
# ============================================================

st.subheader("📥 Weekly Demand Input")

month_plan = [
    ("Jan'26", 5, 1_857_600),
    ("Feb'26", 4, 2_322_000),
    ("Mar'26", 4, 3_168_420),
    ("Apr'26", 5, 3_168_420),
    ("May'26", 4, 3_500_000),
    ("Jun'26", 4, 3_500_000),
    ("Jul'26", 5, 3_200_000),
    ("Aug'26", 4, 3_200_000),
    ("Sep'26", 4, 3_000_000),
    ("Oct'26", 5, 2_800_000),
    ("Nov'26", 4, 2_800_000),
    ("Dec'26", 4, 2_800_000),
]

months = []
demand_values = []
for month, weeks, weekly_demand in month_plan:
    months.extend([month] * weeks)
    demand_values.extend([weekly_demand] * weeks)

default_demand = {
    "CW": list(range(1, len(months) + 1)),
    "Month": months,
    "Demand": demand_values,
}
demand_df = pd.DataFrame(default_demand)

# Allow user to edit demand
edited_demand = st.data_editor(
    demand_df,
    use_container_width=True,
    height=200,
    column_config={
        "CW":     st.column_config.NumberColumn("CW", disabled=True),
        "Month":  st.column_config.TextColumn("Month", disabled=True),
        "Demand": st.column_config.NumberColumn(
            "Weekly Demand (pcs)",
            min_value=0,
            format="%d"
        )
    }
)

# ============================================================
# SECTION 4: CORE CALCULATIONS
# ============================================================

def get_params():
    """整合所有参数"""
    return {
        "stages": {
            "Foundry": {
                "ct":      foundry_ct / 7,
                "transit": foundry_transit / 7,
                "yield":   foundry_yield
            },
            "Bump": {
                "ct":      bump_ct / 7,
                "transit": bump_transit / 7,
                "yield":   bump_yield
            },
            "Testing": {
                "ct":      test_ct / 7,
                "transit": test_transit / 7,
                "yield":   test_yield
            },
            "DPS": {
                "ct":      dps_ct / 7,
                "transit": dps_transit / 7,
                "yield":   dps_yield
            }
        },
        "cpw":          cpw,
        "tbase":        tbase,
        "weekly_output": weekly_output,
        "total_yield":  (foundry_yield * bump_yield *
                        test_yield * dps_yield),
        "total_lt":     int((foundry_ct + foundry_transit +
                            bump_ct + bump_transit +
                            test_ct + test_transit +
                            dps_ct + dps_transit) / 7),
        "output_lag":   int((dps_ct + dps_transit) / 7),
    }


def calc_all(demand_series, p, init_stock):
    """所有计算"""

    n = len(demand_series)

    # ── Wafer Start ────────────────────────────────────────
    wafer_start = np.zeros(n)
    for i in range(n):
        future = i + p["total_lt"]
        d = (demand_series.iloc[future]
             if future < n
             else demand_series.iloc[-1])
        wafer_start[i] = d / p["cpw"] / p["total_yield"]

    # ── Bump Demand (wafers) ───────────────────────────────
    post_bump_y = p["stages"]["Testing"]["yield"] * \
                  p["stages"]["DPS"]["yield"]
    bump_lag    = int(
        p["stages"]["Testing"]["ct"] +
        p["stages"]["Testing"]["transit"] +
        p["stages"]["DPS"]["ct"] +
        p["stages"]["DPS"]["transit"]
    )
    bump = np.zeros(n)
    for i in range(n):
        future = i + bump_lag
        d = (demand_series.iloc[future]
             if future < n
             else demand_series.iloc[-1])
        bump[i] = d / p["cpw"] / post_bump_y

    # ── Sort Demand (wafers) ───────────────────────────────
    post_sort_y = p["stages"]["Testing"]["yield"] * \
                  p["stages"]["DPS"]["yield"]
    sort_lag    = int(
        p["stages"]["DPS"]["ct"] +
        p["stages"]["DPS"]["transit"]
    )
    sort = np.zeros(n)
    for i in range(n):
        future = i + sort_lag
        d = (demand_series.iloc[future]
             if future < n
             else demand_series.iloc[-1])
        sort[i] = d / p["cpw"] / post_sort_y

    # ── DPS Demand (pcs) ───────────────────────────────────
    dps = (sort * p["cpw"] *
           p["stages"]["Testing"]["yield"] *
           p["stages"]["DPS"]["yield"])

    # ── Tester Demand ──────────────────────────────────────
    tester = wafer_start * p["tbase"] / p["weekly_output"]

    # ── DC Stock ───────────────────────────────────────────
    stock = np.zeros(n)
    lag   = p["output_lag"]
    for i in range(n):
        dps_out = (wafer_start[i - lag] * p["cpw"] *
                   p["total_yield"] if i >= lag else 0)
        prev      = stock[i-1] if i > 0 else init_stock
        stock[i]  = max(0, prev + dps_out -
                        demand_series.iloc[i])

    # ── REACH Level (4-week avg) [1] ───────────────────────
    reach = np.zeros(n)
    for i in range(n):
        end      = min(i + 4, n)
        avg_d    = demand_series.iloc[i:end].mean()
        reach[i] = stock[i] / avg_d if avg_d > 0 else 0

    return {
        "wafer_start": wafer_start,
        "bump":        bump,
        "sort":        sort,
        "dps":         dps,
        "tester":      tester,
        "stock":       stock,
        "reach":       reach,
    }

# ============================================================
# SECTION 5: OUTPUT
# ============================================================

if run:
    demand_series = edited_demand["Demand"].reset_index(drop=True)
    p             = get_params()

    with st.spinner("⏳ Calculating..."):
        results = calc_all(demand_series, p, initial_stock)

    st.success("✅ Calculation Complete!")
    st.divider()

    # ── KPI Cards ──────────────────────────────────────────
    st.subheader("📊 Summary KPIs")
    k1, k2, k3, k4, k5 = st.columns(5)

    reach_arr  = results["reach"]
    tester_arr = results["tester"]

    k1.metric(
        "Avg REACH Level",
        f"{reach_arr.mean():.2f}",
        f"Target: {reach_target}"
    )
    k2.metric(
        "Weeks OK",
        f"{((reach_arr >= reach_min) & (reach_arr <= reach_max)).sum()}",
        f"/ 52 weeks"
    )
    k3.metric(
        "Weeks Below Min",
        f"{(reach_arr < reach_min).sum()}",
        f"REACH < {reach_min}"
    )
    k4.metric(
        "Max Tester Demand",
        f"{tester_arr.max():.1f}",
        f"/ {max_testers} available [1]"
    )
    k5.metric(
        "Total Wafer Start",
        f"{results['wafer_start'].sum():.0f}",
        "wafers"
    )

    st.divider()

    # ── Table 1: Wafer Start/Out ───────────────────────────
    st.subheader("📋 Table 1: Wafer Start & Wafer Out / Month")

    table1 = edited_demand.copy()
    table1["Wafer_Start (WSPM)"] = results["wafer_start"].round(1)
    table1["Tester_Demand"]      = results["tester"].round(2)
    table1["Tester_Status"]      = [
        "🔴 OVER" if t > max_testers else "✅ OK"
        for t in results["tester"]
    ]

    st.dataframe(
        table1,
        use_container_width=True,
        height=300
    )

    st.divider()

    # ── Table 3: Wafer Demand by Process ───────────────────
    st.subheader("📋 Table 3: Wafer Demand / Week by Process")

    table3 = pd.DataFrame({
        "CW":               edited_demand["CW"],
        "Month":            edited_demand["Month"],
        "Bump (wafers)":    results["bump"].round(1),
        "Sort (wafers)":    results["sort"].round(1),
        "DPS (pcs)":        results["dps"].round(0).astype(int),
    })

    st.dataframe(
        table3,
        use_container_width=True,
        height=300
    )

    st.divider()

    # ── Graphs ─────────────────────────────────────────────
    st.subheader("📈 Output Graphs")

    tab1, tab2, tab3, tab4 = st.tabs([
        "① Tester Demand",
        "② VRFC Demand",
        "③ REACH Level",
        "④ DC Stock"
    ])

    cw = edited_demand["CW"].tolist()

    # Graph 1: Tester [1]
    with tab1:
        st.markdown("**# of Tester by Basic Type**")
        fig1 = go.Figure()
        fig1.add_trace(go.Bar(
            x=cw, y=results["tester"],
            name=product,
            marker_color=[
                "crimson" if t > max_testers
                else "steelblue"
                for t in results["tester"]
            ]
        ))
        fig1.add_hline(
            y=max_testers,
            line_dash="dash",
            line_color="red",
            annotation_text=f"Max: {max_testers} [1]"
        )
        fig1.update_layout(
            xaxis_title="Calendar Week",
            yaxis_title="No. of Testers",
            height=400
        )
        st.plotly_chart(fig1, use_container_width=True)

    # Graph 2: VRFC Demand
    with tab2:
        st.markdown("**VRFC Demand (Customer Demand)**")
        fig2 = go.Figure()
        fig2.add_trace(go.Bar(
            x=cw,
            y=demand_series,
            name="Weekly Demand",
            marker_color="steelblue"
        ))
        fig2.update_layout(
            xaxis_title="Calendar Week",
            yaxis_title="Demand (pcs)",
            height=400
        )
        st.plotly_chart(fig2, use_container_width=True)

    # Graph 3: REACH [1]
    with tab3:
        st.markdown("**REACH Development**")
        fig3 = go.Figure()
        fig3.add_trace(go.Scatter(
            x=cw, y=results["reach"],
            mode="lines+markers",
            name="REACH Level",
            line=dict(color="royalblue", width=2)
        ))
        for y, color, label in [
            (reach_min,    "red",    f"Min: {reach_min}"),
            (reach_target, "green",  f"Target: {reach_target}"),
            (reach_max,    "orange", f"Max: {reach_max}"),
        ]:
            fig3.add_hline(
                y=y, line_dash="dash",
                line_color=color,
                annotation_text=label
            )
        fig3.update_layout(
            xaxis_title="Calendar Week",
            yaxis_title="REACH Level",
            height=400
        )
        st.plotly_chart(fig3, use_container_width=True)

    # Graph 4: DC Stock
    with tab4:
        st.markdown("**DC Stock Development**")
        fig4 = go.Figure()
        fig4.add_trace(go.Scatter(
            x=cw, y=results["stock"],
            fill="tozeroy",
            name="DC Stock",
            line=dict(color="mediumseagreen", width=2),
            fillcolor="rgba(60,179,113,0.15)"
        ))
        fig4.update_layout(
            xaxis_title="Calendar Week",
            yaxis_title="Stock (pcs)",
            height=400
        )
        st.plotly_chart(fig4, use_container_width=True)

    st.divider()

    # ── Export ─────────────────────────────────────────────
    st.subheader("💾 Export Results")

    final_df = pd.DataFrame({
        "CW":            edited_demand["CW"],
        "Month":         edited_demand["Month"],
        "Demand":        demand_series,
        "Wafer_Start":   results["wafer_start"].round(1),
        "Bump (wafers)": results["bump"].round(1),
        "Sort (wafers)": results["sort"].round(1),
        "DPS (pcs)":     results["dps"].round(0),
        "Tester":        results["tester"].round(2),
        "DC_Stock":      results["stock"].round(0),
        "REACH":         results["reach"].round(2),
        "REACH_Status":  [
            "🔴 LOW"  if r < reach_min  else
            "🟡 HIGH" if r > reach_max  else
            "✅ OK"
            for r in results["reach"]
        ]
    })

    col1, col2 = st.columns(2)

    with col1:
        csv = final_df.to_csv(index=False)
        st.download_button(
            "⬇️ Download CSV",
            csv,
            "demand_control_output.csv",
            use_container_width=True
        )

    with col2:
        final_df.to_excel("output.xlsx", index=False)
        with open("output.xlsx", "rb") as f:
            st.download_button(
                "⬇️ Download Excel",
                f,
                "demand_control_output.xlsx",
                use_container_width=True
            )

else:
    # 未运行时显示提示
    st.info(
        "👈 设置左边参数后点击 "
        "**Run Calculation** 开始计算"
    )