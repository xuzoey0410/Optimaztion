from io import BytesIO
from pathlib import Path
import tempfile

import pandas as pd
import plotly.express as px
import streamlit as st
from openpyxl.chart import AreaChart, BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter

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


def build_month_product_table(monthly_summary_df, value_col, aggfunc="sum"):
    output_df = monthly_summary_df.pivot_table(
        index="Month",
        columns="Product_Key",
        values=value_col,
        aggfunc=aggfunc,
        fill_value=0,
        margins=True,
        margins_name="Grand Total",
    ).reset_index()

    output_df = output_df.rename(columns={"Month": "Row Labels"})
    if aggfunc == "mean":
        value_columns = [col for col in output_df.columns if col != "Row Labels"]
        output_df[value_columns] = output_df[value_columns].round(2)
    return output_df


def build_wafer_start_table(monthly_summary_df):
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


def build_bump_sort_dps_table(monthly_summary_df):
    long_df = monthly_summary_df.melt(
        id_vars=["Basic_Type", "Product_Key", "Month"],
        value_vars=["Bump_Wafer", "Sort_Wafer", "DPS_Chip"],
        var_name="Metric",
        value_name="Value",
    )
    long_df["Metric"] = long_df["Metric"].replace({
        "Bump_Wafer": "Bump",
        "Sort_Wafer": "Sort",
        "DPS_Chip": "DPS",
    })

    table_df = long_df.pivot_table(
        index=["Basic_Type", "Product_Key", "Metric"],
        columns="Month",
        values="Value",
        aggfunc="sum",
        fill_value=0,
        margins=True,
        margins_name="Grand Total",
    ).reset_index()

    return table_df.rename(columns={"Basic_Type": "Basic Type", "Product_Key": "Row Labels"})


def build_graph_outputs(monthly_summary_df):
    stock_df = build_month_product_table(monthly_summary_df, "End_Stock", "sum")
    demand_df = build_month_product_table(monthly_summary_df, "Demand", "sum")
    month_rows = demand_df["Row Labels"].astype(str) != "Grand Total"
    stock_df["Next Month Demand"] = 0
    stock_df.loc[month_rows, "Next Month Demand"] = demand_df.loc[month_rows, "Grand Total"].shift(-1).fillna(0).values

    return [
        ("Tester Used", build_month_product_table(monthly_summary_df, "Max_TesterUsed", "sum"), "bar"),
        ("vRFN Demand", build_month_product_table(monthly_summary_df, "Demand", "sum"), "bar"),
        ("Reach Level", build_month_product_table(monthly_summary_df, "Avg_REACH", "mean"), "line"),
        ("Stock", stock_df, "stock"),
    ]


def style_excel_range(worksheet, header_row, last_row, last_col):
    header_fill = PatternFill("solid", fgColor="1F4E78")
    total_fill = PatternFill("solid", fgColor="D9EAF7")
    metric_fill = PatternFill("solid", fgColor="EAF3F8")
    white_font = Font(color="FFFFFF", bold=True)
    bold_font = Font(bold=True)
    thin_side = Side(style="thin", color="D0D7DE")
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    for cell in worksheet[header_row]:
        if cell.column <= last_col:
            cell.fill = header_fill
            cell.font = white_font
            cell.alignment = Alignment(horizontal="center")
            cell.border = border

    for row in worksheet.iter_rows(min_row=header_row + 1, max_row=last_row, max_col=last_col):
        is_total_row = any(str(cell.value) == "Grand Total" for cell in row[:3])
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(horizontal="center")
            if is_total_row:
                cell.fill = total_fill
                cell.font = bold_font
            elif cell.column <= 3:
                cell.fill = metric_fill

    worksheet.freeze_panes = worksheet.cell(row=header_row + 1, column=2)
    for column_index in range(1, last_col + 1):
        max_len = max(
            len(str(worksheet.cell(row=row_index, column=column_index).value or ""))
            for row_index in range(header_row, last_row + 1)
        )
        worksheet.column_dimensions[get_column_letter(column_index)].width = min(max(max_len + 2, 11), 18)


def add_table_block(worksheet, title, output_df, start_row):
    worksheet.cell(row=start_row, column=1, value=title)
    worksheet.cell(row=start_row, column=1).font = Font(bold=True, size=14, color="1F4E78")
    header_row = start_row + 1

    for row_index, row_values in enumerate(dataframe_to_rows(output_df, index=False, header=True), header_row):
        for column_index, value in enumerate(row_values, 1):
            worksheet.cell(row=row_index, column=column_index, value=value)

    style_excel_range(worksheet, header_row, header_row + len(output_df), len(output_df.columns))
    return header_row, header_row + len(output_df), len(output_df.columns)


def add_excel_chart(worksheet, title, chart_kind, header_row, last_row, last_col, anchor):
    headers = [worksheet.cell(row=header_row, column=col).value for col in range(1, last_col + 1)]
    overlay_col = headers.index("Next Month Demand") + 1 if "Next Month Demand" in headers else None
    chart_last_col = last_col

    if overlay_col:
        chart_last_col = overlay_col - 1
    if worksheet.cell(row=header_row, column=chart_last_col).value == "Grand Total":
        chart_last_col -= 1
    if chart_last_col < 2 or last_row <= header_row:
        return

    chart_map = {"bar": BarChart, "line": LineChart, "area": AreaChart, "stock": BarChart}
    chart = chart_map[chart_kind]()
    chart.title = title
    chart.height = 10
    chart.width = 20

    data = Reference(worksheet, min_col=2, max_col=chart_last_col, min_row=header_row, max_row=last_row)
    categories = Reference(worksheet, min_col=1, min_row=header_row + 1, max_row=last_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)

    if chart_kind in ["bar", "stock"]:
        chart.type = "col"
        chart.style = 10
    elif chart_kind == "area":
        chart.grouping = "stacked"

    if chart_kind == "stock" and overlay_col:
        line_chart = LineChart()
        line_data = Reference(worksheet, min_col=overlay_col, max_col=overlay_col, min_row=header_row, max_row=last_row)
        line_chart.add_data(line_data, titles_from_data=True)
        line_chart.set_categories(categories)
        line_chart.y_axis.axId = 200
        line_chart.y_axis.title = "Next Month Demand"
        chart += line_chart

    worksheet.add_chart(chart, anchor)


def style_excel_sheet(worksheet):
    style_excel_range(worksheet, 1, worksheet.max_row, worksheet.max_column)


def make_output_bytes(graph_outputs, wafer_start_table, bump_sort_dps_table):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        wafer_start_table.to_excel(writer, sheet_name="Wafer_Start", index=False)
        bump_sort_dps_table.to_excel(writer, sheet_name="Bump_Sort_DPS", index=False)
        style_excel_sheet(writer.book["Wafer_Start"])
        style_excel_sheet(writer.book["Bump_Sort_DPS"])

        graph_sheet = writer.book.create_sheet("Graph", 0)
        for start_row, (title, graph_df, chart_kind) in zip([1, 30, 59, 88], graph_outputs):
            header_row, last_row, last_col = add_table_block(graph_sheet, title, graph_df, start_row)
            add_excel_chart(graph_sheet, title, chart_kind, header_row, last_row, last_col, f"L{start_row}")
    output.seek(0)
    return output.getvalue()


def style_display_table(output_df):
    numeric_cols = output_df.select_dtypes(include="number").columns.tolist()

    def heatmap_column(column):
        values = pd.to_numeric(column, errors="coerce")
        min_value = values.min()
        max_value = values.max()
        if pd.isna(min_value) or pd.isna(max_value) or min_value == max_value:
            return ["background-color: #f7fbff" for _ in column]

        styles = []
        for value in values:
            if pd.isna(value):
                styles.append("")
                continue
            strength = (value - min_value) / (max_value - min_value)
            blue = int(245 - 55 * strength)
            green = int(250 - 80 * strength)
            red = int(255 - 120 * strength)
            styles.append(f"background-color: rgb({red}, {green}, {blue})")
        return styles

    def highlight_total(row):
        is_total = any(str(value) == "Grand Total" for value in row.values)
        return ["background-color: #d9eaf7; font-weight: 700" if is_total else "" for _ in row]

    styler = (
        output_df.style
        .format({col: "{:,.0f}" for col in numeric_cols})
        .set_table_styles([
            {"selector": "th", "props": [("background-color", "#1f4e78"), ("color", "white"), ("font-weight", "700")]},
            {"selector": "td", "props": [("border-color", "#d0d7de")]},
        ])
    )

    if numeric_cols:
        styler = styler.apply(heatmap_column, subset=numeric_cols, axis=0)

    return styler.apply(highlight_total, axis=1)


def draw_graphs(graph_outputs):
    st.subheader("Graph")
    for title, graph_df, chart_kind in graph_outputs:
        chart_df = graph_df[graph_df["Row Labels"].astype(str) != "Grand Total"].copy()

        if chart_kind == "stock":
            demand_line = chart_df[["Row Labels", "Next Month Demand"]].copy()
            product_cols = [col for col in chart_df.columns if col not in ["Row Labels", "Grand Total", "Next Month Demand"]]
            chart_df = chart_df.melt(
                id_vars="Row Labels",
                value_vars=product_cols,
                var_name="Product",
                value_name=title,
            )
            chart = px.bar(chart_df, x="Row Labels", y=title, color="Product", title=title)
            chart.update_layout(barmode="stack")
            chart.add_scatter(
                x=demand_line["Row Labels"],
                y=demand_line["Next Month Demand"],
                mode="lines+markers",
                name="Next Month Demand",
                line={"color": "#d62728", "width": 3},
                marker={"size": 8},
            )
            chart.update_layout(xaxis_title="Month", yaxis_title="Stock / Next Month Demand")
            st.plotly_chart(chart, use_container_width=True)
            continue

        chart_df = chart_df.melt(id_vars="Row Labels", var_name="Product", value_name=title)
        chart_df = chart_df[chart_df["Product"] != "Grand Total"]

        if chart_kind == "line":
            chart = px.line(chart_df, x="Row Labels", y=title, color="Product", markers=True, title=title)
        elif chart_kind == "area":
            chart = px.area(chart_df, x="Row Labels", y=title, color="Product", title=title)
        else:
            chart = px.bar(chart_df, x="Row Labels", y=title, color="Product", title=title)
            chart.update_layout(barmode="stack")

        chart.update_layout(xaxis_title="Month", yaxis_title=title)
        st.plotly_chart(chart, use_container_width=True)


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
    graph_outputs = build_graph_outputs(monthly_summary_df)
    wafer_start_table = build_wafer_start_table(monthly_summary_df)
    bump_sort_dps_table = build_bump_sort_dps_table(monthly_summary_df)
    output_bytes = make_output_bytes(graph_outputs, wafer_start_table, bump_sort_dps_table)

    st.success("Plan generated")

    draw_graphs(graph_outputs)

    st.subheader("Wafer Start")
    st.dataframe(style_display_table(wafer_start_table), use_container_width=True)

    st.subheader("Bump Sort DPS")
    st.dataframe(style_display_table(bump_sort_dps_table), use_container_width=True)

    st.download_button(
        "Download Final Output Excel",
        data=output_bytes,
        file_name="final_output_graphs_tables.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
