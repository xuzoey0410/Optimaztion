from io import BytesIO
from pathlib import Path
import tempfile

import pandas as pd
import plotly.express as px
import streamlit as st

import output123 as planner


APP_DIR = Path(__file__).parent
SAMPLE_INPUT = APP_DIR / "input_sample.xlsx"
LOCAL_INPUT = APP_DIR / "All_Output1.xlsx"


st.set_page_config(page_title="Production Planner", layout="wide")
st.title("Production Planner")


def sheet_name(excel_file, target):
    lookup = {s.strip().lower(): s for s in excel_file.sheet_names}
    return lookup.get(target.strip().lower())


@st.cache_data(show_spinner=False)
def load_tables(file_bytes):
    if file_bytes:
        source = BytesIO(file_bytes)
    elif SAMPLE_INPUT.exists():
        source = SAMPLE_INPUT
    elif LOCAL_INPUT.exists():
        source = LOCAL_INPUT
    else:
        return None

    xl = pd.ExcelFile(source)

    def read(target, required=True):
        name = sheet_name(xl, target)
        if name is None:
            if required:
                raise ValueError(f"Missing sheet: {target}")
            return pd.DataFrame()
        return pd.read_excel(xl, sheet_name=name)

    return {
        "flow": read(planner.FLOW_SHEET),
        "product": read(planner.PRODUCT_SHEET),
        "demand": read(planner.DEMAND_SHEET),
        "inventory": read(planner.INVENTORY_SHEET),
        "target": read(planner.TARGET_SHEET, required=False),
    }


def target_value(target_df, key, default):
    if target_df.empty or not {"Columns", "Value"}.issubset(target_df.columns):
        return default
    matches = target_df[target_df["Columns"].astype(str).str.strip().str.lower() == key.lower()]
    if matches.empty:
        return default
    value = pd.to_numeric(matches.iloc[0]["Value"], errors="coerce")
    return default if pd.isna(value) else value


def build_input_workbook(flow_df, product_df, demand_df, inventory_df, target_reach, tester_number):
    target_df = pd.DataFrame({
        "Columns": ["Target Reach Level", "Tester Number"],
        "Value": [target_reach, tester_number],
    })

    temp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    temp_path = Path(temp.name)
    temp.close()

    with pd.ExcelWriter(temp_path, engine="openpyxl") as writer:
        flow_df.to_excel(writer, sheet_name=planner.FLOW_SHEET, index=False)
        product_df.to_excel(writer, sheet_name=planner.PRODUCT_SHEET, index=False)
        demand_df.to_excel(writer, sheet_name=planner.DEMAND_SHEET, index=False)
        inventory_df.to_excel(writer, sheet_name=planner.INVENTORY_SHEET, index=False)
        target_df.to_excel(writer, sheet_name=planner.TARGET_SHEET, index=False)

    return temp_path


def build_table_output(monthly_summary_df):
    table_df = monthly_summary_df.pivot_table(
        index=["Basic_Type", "Product_Key"],
        columns="Month",
        values="WaferStart",
        aggfunc="sum",
        fill_value=0,
        margins=True,
        margins_name="Grand Total",
    ).reset_index()

    table_df = table_df.rename(columns={"Basic_Type": "Basic Type", "Product_Key": "Row Labels"})
    return table_df


def build_graph_output(monthly_summary_df):
    graph_df = monthly_summary_df.pivot_table(
        index="Month",
        columns="Product_Key",
        values="Max_TesterUsed",
        aggfunc="sum",
        fill_value=0,
        margins=True,
        margins_name="Grand Total",
    ).reset_index()

    return graph_df.rename(columns={"Month": "Row Labels"})


def make_output_bytes(graph_df, table_df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        graph_df.to_excel(writer, sheet_name="Graph", index=False)
        table_df.to_excel(writer, sheet_name="Table", index=False)
    output.seek(0)
    return output.getvalue()


def draw_graph(graph_df):
    chart_df = graph_df[graph_df["Row Labels"].astype(str) != "Grand Total"].copy()
    chart_df = chart_df.melt(
        id_vars="Row Labels",
        var_name="Product",
        value_name="Tester Used",
    )
    chart_df = chart_df[chart_df["Product"] != "Grand Total"]

    tester_fig = px.bar(
        chart_df,
        x="Row Labels",
        y="Tester Used",
        color="Product",
        title="Monthly Tester Used",
    )
    tester_fig.update_layout(barmode="stack", xaxis_title="Month", yaxis_title="Tester Used")

    st.subheader("Graph")
    st.plotly_chart(tester_fig, use_container_width=True)


uploaded = st.sidebar.file_uploader("Upload input Excel", type=["xlsx"])
file_bytes = uploaded.getvalue() if uploaded else None

with st.sidebar:
    st.header("Optimization")
    planner.OPTIMIZER_MAXITER = st.number_input("Max iterations", min_value=1, max_value=500, value=int(planner.OPTIMIZER_MAXITER), step=10)
    planner.TESTER_LEVEL_WEIGHT = st.number_input("Tester level weight", min_value=0.0, value=float(planner.TESTER_LEVEL_WEIGHT), step=0.01)
    planner.TESTER_CHANGE_WEIGHT = st.number_input("Tester change weight", min_value=0.0, value=float(planner.TESTER_CHANGE_WEIGHT), step=0.1)
    planner.TESTER_JUMP_WEIGHT = st.number_input("Tester jump weight", min_value=0.0, value=float(planner.TESTER_JUMP_WEIGHT), step=1.0)
    planner.WAFER_CHANGE_WEIGHT = st.number_input("Wafer change weight", min_value=0.0, value=float(planner.WAFER_CHANGE_WEIGHT), step=1.0)
    planner.WAFER_CHANGE_LIMIT = st.number_input("Wafer change limit", min_value=0.0, max_value=1.0, value=float(planner.WAFER_CHANGE_LIMIT), step=0.05)


try:
    tables = load_tables(file_bytes)
except Exception as exc:
    st.error(str(exc))
    st.stop()

if tables is None:
    st.info("Please upload an input Excel file from the sidebar, or add input_sample.xlsx to the GitHub repo.")
    st.stop()

target_reach_default = float(target_value(tables["target"], "Target Reach Level", planner.DEFAULT_TESTER_CONFIG["target_REACH"]))
tester_number_default = int(target_value(tables["target"], "Tester Number", planner.DEFAULT_TESTER_CONFIG["available"]))

top_cols = st.columns(2)
with top_cols[0]:
    target_reach = st.number_input("Target Reach Level", min_value=0.0, value=target_reach_default, step=0.5)
with top_cols[1]:
    tester_number = st.number_input("Tester Number", min_value=1, value=tester_number_default, step=1)

tabs = st.tabs(["Flow", "Product", "Demand", "Inventory"])
with tabs[0]:
    flow_df = st.data_editor(tables["flow"], use_container_width=True, num_rows="dynamic", key="flow")
with tabs[1]:
    product_df = st.data_editor(tables["product"], use_container_width=True, num_rows="dynamic", key="product")
with tabs[2]:
    demand_df = st.data_editor(tables["demand"], use_container_width=True, num_rows="dynamic", key="demand")
with tabs[3]:
    inventory_df = st.data_editor(tables["inventory"], use_container_width=True, num_rows="dynamic", key="inventory")


if st.button("Run plan", type="primary"):
    with st.spinner("Running optimization..."):
        try:
            input_path = build_input_workbook(flow_df, product_df, demand_df, inventory_df, target_reach, tester_number)
            stages, products, weekly_demand_map, month_label_map, tester_config, inventory_map = planner.load_all_inputs(input_path)
            results = planner.run_all_products(stages, products, weekly_demand_map, month_label_map, tester_config, inventory_map)
        except Exception as exc:
            st.error(str(exc))
            st.stop()

    all_plan_df, summary_df, monthly_summary_df, skipped_df = results
    graph_df = build_graph_output(monthly_summary_df)
    table_df = build_table_output(monthly_summary_df)
    output_bytes = make_output_bytes(graph_df, table_df)

    st.success("Plan generated")

    draw_graph(graph_df)

    st.subheader("Table")
    st.dataframe(table_df, use_container_width=True)

    st.download_button(
        "Download Graph and Table Excel",
        data=output_bytes,
        file_name="graph_table_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
