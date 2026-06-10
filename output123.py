from pathlib import Path
import re
import numpy as np
import pandas as pd
from scipy.optimize import minimize

INPUT_EXCEL = Path(r"C:\Users\xuzi\Desktop\Demand Project\All_Output1.xlsx")
OUTPUT_EXCEL = Path(r"C:\Users\xuzi\Desktop\Demand Project\multi_product_plan_output.xlsx")

FLOW_SHEET = "Parameter_Flow"
PRODUCT_SHEET = "Parameter_Product"
DEMAND_SHEET = "Demand_Input"
INVENTORY_SHEET = "Invent_Ex_Input"
TARGET_SHEET = "Target and Constraints"
OPTIMIZER_MAXITER = 100
TESTER_LEVEL_WEIGHT = 0.05
TESTER_CHANGE_WEIGHT = 0.5
TESTER_JUMP_LIMIT = 3
TESTER_JUMP_WEIGHT = 20
WAFER_CHANGE_LIMIT = 0.2
WAFER_CHANGE_WEIGHT = 20
FORECAST_WEEKS = 8
FORECAST_LOOKBACK_WEEKS = 24

DEFAULT_TESTER_CONFIG = {
    "available": 19,
    "target_REACH": 4,
    "reach_window": 4,
}

def normalize_key(x):
    if pd.isna(x):
        return ""
    return re.sub(r"[\s\u3000]+", "", str(x).upper())


def clean_columns(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def clean_name(x):
    return re.sub(r"[\s/]+", "_", str(x).strip().lower())


def find_sheet(excel_path: Path, target_name: str) -> str:
    xl = pd.ExcelFile(excel_path)
    sheet_map = {s.strip().lower(): s for s in xl.sheet_names}

    key = target_name.strip().lower()

    if key not in sheet_map:
        raise ValueError(
            f"找不到 Sheet: {target_name}\n"
            f"当前 Excel Sheets: {xl.sheet_names}"
        )

    return sheet_map[key]


def find_col(df, candidates, required=True):
    cleaned_map = {clean_name(c): c for c in df.columns}

    for c in candidates:
        key = clean_name(c)
        if key in cleaned_map:
            return cleaned_map[key]

    for col in df.columns:
        col_clean = clean_name(col)
        for cand in candidates:
            if clean_name(cand) in col_clean:
                return col

    if required:
        raise ValueError(
            f"找不到列: {candidates}\n"
            f"当前列名: {list(df.columns)}"
        )

    return None


def normalize_yield(v):
    if pd.isna(v):
        return 1.0

    if isinstance(v, str):
        v = v.strip().replace("%", "")
        if v == "":
            return 1.0

    v = float(v)
    return v / 100 if v > 1 else v


def is_month_col(c):
    s = str(c).strip()

    if s.lower().startswith("unnamed"):
        return False

    if re.search(r"^\d{4}[-/]\d{1,2}", s):
        return True

    if re.search(r"^[A-Za-z]{3}[-']?\d{2,4}", s):
        return True

    try:
        dt = pd.to_datetime(s, errors="coerce")
        return not pd.isna(dt)
    except Exception:
        return False


def month_label(x):
    if isinstance(x, pd.Timestamp):
        return x.strftime("%Y-%m")

    s = str(x).strip()

    try:
        dt = pd.to_datetime(s, errors="coerce")
        if not pd.isna(dt):
            return dt.strftime("%Y-%m")
    except Exception:
        pass

    return s


def month_to_weeks(month_value):
    if isinstance(month_value, pd.Timestamp):
        month_num = month_value.month
    else:
        s = str(month_value).strip()

        try:
            dt = pd.to_datetime(s, errors="coerce")
            if not pd.isna(dt):
                month_num = dt.month
            else:
                month_name_map = {
                    "jan": 1,
                    "feb": 2,
                    "mar": 3,
                    "apr": 4,
                    "may": 5,
                    "jun": 6,
                    "jul": 7,
                    "aug": 8,
                    "sep": 9,
                    "oct": 10,
                    "nov": 11,
                    "dec": 12,
                }

                m = re.search(r"[A-Za-z]{3}", s)
                if not m:
                    raise ValueError(f"无法识别月份: {month_value}")

                month_num = month_name_map[m.group(0).lower()]
        except Exception:
            raise ValueError(f"无法识别月份: {month_value}")

    return 5 if month_num in [3, 6, 9, 12] else 4


def load_flow_from_excel(excel_path: Path, sheet_name: str = FLOW_SHEET):
    real_sheet = find_sheet(excel_path, sheet_name)
    df = pd.read_excel(excel_path, sheet_name=real_sheet)
    df = clean_columns(df)
    df = df.dropna(how="all")

    stage_col = find_col(df, ["Stages", "Stage", "Process", "Flow"])
    week_col = find_col(df, ["Time/Week", "Time Week", "Cycle_Time_Week", "Cycle Time Week"])
    yield_col = find_col(df, ["Yield", "Yield Rate"], required=False)

    stages = {}
    current_stage = None

    for _, row in df.iterrows():
        raw_stage = row[stage_col]

        if pd.isna(raw_stage):
            continue

        stage_name = str(raw_stage).strip()
        stage_lower = stage_name.lower()

        time_week = 0 if pd.isna(row[week_col]) else int(round(float(row[week_col])))

        if "transit" in stage_lower or "transist" in stage_lower:
            if current_stage is not None:
                stages[current_stage]["transit"] += time_week
            continue

        y = normalize_yield(row[yield_col]) if yield_col else 1.0

        stages[stage_name] = {"cycle_time": time_week, "transit": 0, "yield": y}

        current_stage = stage_name

    if not stages:
        raise ValueError("Parameter_Flow 没有读到任何 stage。")

    return stages

def load_products_from_excel(excel_path: Path, sheet_name: str = PRODUCT_SHEET):
    real_sheet = find_sheet(excel_path, sheet_name)

    df = pd.read_excel(excel_path, sheet_name=real_sheet)
    df = clean_columns(df)
    df = df.dropna(how="all")

    product_col = find_col(df, ["Product"])
    basic_col = find_col(df, ["Basic Type", "Basic_Type", "Type"], required=False)
    cpw_col = find_col(df, ["CPW"])
    weekly_output_col = find_col(
        df,
        [
            "Wafer / tester / week",
            "Wafer/tester/week",
            "Weekly Output",
            "weekly_output",
            "Tester Output",
            "Output Per Week",
            "UPH",
        ],
    )
    products = {}

    for _, row in df.iterrows():
        product_key = normalize_key(row[product_col])

        if product_key == "":
            continue

        cpw = pd.to_numeric(row[cpw_col], errors="coerce")
        weekly_output = pd.to_numeric(row[weekly_output_col], errors="coerce")

        if pd.isna(cpw) or pd.isna(weekly_output):
            continue

        if product_key in products:
            continue

        products[product_key] = {
            "Product_Key": product_key,
            "Original_Product": str(row[product_col]).strip(),
            "Basic_Type": str(row[basic_col]).strip() if basic_col and not pd.isna(row[basic_col]) else "",
            "CPW": float(cpw),
            "weekly_output": float(weekly_output),
        }

    if not products:
        raise ValueError("Parameter_Product 没有读到任何有效 product。")

    return products

def load_demand_from_excel(excel_path: Path, sheet_name: str = DEMAND_SHEET):
    real_sheet = find_sheet(excel_path, sheet_name)

    df = pd.read_excel(excel_path, sheet_name=real_sheet)
    df = clean_columns(df)
    df = df.dropna(how="all")

    product_col = find_col(df, ["Product"])
    basic_col = find_col(df, ["Basic Type", "Basic_Type", "Type"], required=False)

    id_cols = [c for c in [basic_col, product_col] if c is not None]
    month_cols = [c for c in df.columns if c not in id_cols and is_month_col(c)]

    if not month_cols:
        raise ValueError(
            "Demand_Input 没有找到月份列。\n"
            f"当前列名: {list(df.columns)}"
        )

    df["_Product_Key"] = df[product_col].apply(normalize_key)

    for c in month_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    grouped = (
        df[df["_Product_Key"] != ""]
        .groupby("_Product_Key", as_index=False)[month_cols]
        .sum()
    )

    weekly_demand_map = {}
    month_label_map = {}

    for _, row in grouped.iterrows():
        product_key = str(row["_Product_Key"]).strip()

        weekly_values = []
        month_labels = []

        for m in month_cols:
            total_demand = float(row[m])
            weeks = month_to_weeks(m)
            weekly_demand = total_demand / weeks
            label = month_label(m)

            weekly_values.extend([weekly_demand] * weeks)
            month_labels.extend([label] * weeks)

        weekly_demand_map[product_key] = pd.Series(weekly_values, dtype=float)
        month_label_map[product_key] = month_labels

    if not weekly_demand_map:
        raise ValueError("Demand_Input 没有读到任何 demand。")

    return weekly_demand_map, month_label_map

def load_tester_config_from_excel(excel_path: Path):
    xl = pd.ExcelFile(excel_path)
    sheet_lookup = {s.strip().lower(): s for s in xl.sheet_names}
    config = DEFAULT_TESTER_CONFIG.copy()

    if TARGET_SHEET.strip().lower() not in sheet_lookup:
        return config

    df = clean_columns(pd.read_excel(excel_path, sheet_name=sheet_lookup[TARGET_SHEET.strip().lower()]))
    key_col = find_col(df, ["Columns", "Column", "Parameter", "Key"])
    value_col = find_col(df, ["Value", "Val"])

    for _, row in df.iterrows():
        key = clean_name(row[key_col])
        val = pd.to_numeric(row[value_col], errors="coerce")

        if pd.isna(val):
            continue
        if key in ["target_reach_level", "target_reach"]:
            config["target_REACH"] = float(val)
        elif key in ["tester_number", "tester", "tester_capacity", "available"]:
            config["available"] = int(val)

    return config


def load_inventory_from_excel(excel_path: Path):
    real_sheet = find_sheet(excel_path, INVENTORY_SHEET)
    df = clean_columns(pd.read_excel(excel_path, sheet_name=real_sheet)).dropna(how="all")
    product_col = find_col(df, ["Product"])
    inventory_col = find_col(
        df,
        [
            "Existing DC inventory",
            "Existing DC inventories",
            "Exisitng DC inventories",
            "Inventory",
            "Stock",
        ],
    )

    df["_Product_Key"] = df[product_col].apply(normalize_key)
    df["_Inventory"] = pd.to_numeric(df[inventory_col], errors="coerce").fillna(0)

    return df[df["_Product_Key"] != ""].groupby("_Product_Key")["_Inventory"].sum().to_dict()


def load_all_inputs(excel_path: Path):
    if not excel_path.exists():
        raise FileNotFoundError(f"找不到 Excel 文件: {excel_path}")
    return (
        load_flow_from_excel(excel_path),
        load_products_from_excel(excel_path),
        *load_demand_from_excel(excel_path),
        load_tester_config_from_excel(excel_path),
        load_inventory_from_excel(excel_path),
    )

def get_total_params(stages, products, product):
    total_yield = 1.0
    total_lt = 0
    bump_lag = sort_lag = dps_lag = None
    dps_yield = None

    for stage, params in stages.items():
        stage_lower = stage.lower()

        if bump_lag is None and "bump" in stage_lower:
            bump_lag = total_lt
        if sort_lag is None and ("sort" in stage_lower or "test" in stage_lower):
            sort_lag = total_lt

        total_yield *= params["yield"]
        total_lt += params["cycle_time"] + params["transit"]

        if dps_lag is None and "dps" in stage_lower:
            dps_lag = total_lt
            dps_yield = total_yield

    return {
        "total_yield": total_yield,
        "total_lt": int(total_lt),
        "output_lag": int(total_lt),
        "bump_lag": int(bump_lag or 0),
        "sort_lag": int(sort_lag or 0),
        "dps_lag": int(dps_lag or total_lt),
        "dps_yield": dps_yield or total_yield,
        "cpw": products[product]["CPW"],
        "weekly_out": products[product]["weekly_output"],
    }


def build_planning_horizon(weekly_demand, months, pre_weeks):
    known_demand = weekly_demand.reset_index(drop=True)
    forecast_weeks = max(FORECAST_WEEKS, pre_weeks)
    forecast_base = known_demand.tail(min(FORECAST_LOOKBACK_WEEKS, len(known_demand))).mean()
    forecast_demand = pd.Series([forecast_base] * forecast_weeks, dtype=float)
    pre_zero = pd.Series(np.zeros(pre_weeks), dtype=float)

    demand_opt = pd.concat(
        [pre_zero, known_demand, forecast_demand],
        ignore_index=True,
    )

    months_ext = ["Pre-Demand"] * pre_weeks + list(months) + ["Forecast"] * forecast_weeks
    known_start_idx = pre_weeks
    forecast_start_idx = pre_weeks + len(known_demand)

    return demand_opt, months_ext, known_start_idx, forecast_start_idx


def calc_wafer_start(demand, p):
    n = len(demand)
    wafer = np.zeros(n)

    for i in range(n):
        future = i + p["total_lt"]
        d = demand.iloc[future] if future < n else demand.iloc[-1]
        wafer[i] = d / p["cpw"] / p["total_yield"]

    return pd.Series(wafer)


def calc_tester(wafer, p):
    return wafer / p["weekly_out"]


def calc_dps_out(wafer, p, integer=False):
    n = len(wafer)
    lag = int(p["output_lag"])
    out = np.zeros(n, dtype=int if integer else float)

    for i in range(n):
        if i >= lag:
            val = wafer.iloc[i - lag] * p["cpw"] * p["total_yield"]
        else:
            val = 0

        out[i] = int(round(val)) if integer else val

    return pd.Series(out)


def calc_shifted_output(wafer, lag, multiplier=1.0, integer=False):
    n = len(wafer)
    out = np.zeros(n, dtype=int if integer else float)

    for i in range(n):
        val = wafer.iloc[i - lag] * multiplier if i >= lag else 0
        out[i] = int(round(val)) if integer else val

    return pd.Series(out)


def calc_stock(wafer, demand, p, init_stock=0, integer=False):
    n = len(demand)
    stock = np.zeros(n, dtype=int if integer else float)
    dps_out = calc_dps_out(wafer, p, integer=integer)

    for i in range(n):
        prev = stock[i - 1] if i > 0 else init_stock
        dem = 0 if pd.isna(demand.iloc[i]) else demand.iloc[i]

        if integer:
            dem = int(round(dem))

        current = prev + dps_out.iloc[i] - dem
        stock[i] = max(0, int(round(current)) if integer else current)

    return pd.Series(stock)


def calc_reach_4week_avg(stock, demand, window=4, start_idx=0):
    n = len(demand)
    reach = np.full(n, np.nan, dtype=float)

    for i in range(start_idx, n):
        end_idx = min(i + window, n)
        future_demand = demand.iloc[i:end_idx]
        avg_demand = future_demand.mean()

        if avg_demand > 0:
            reach[i] = stock.iloc[i] / avg_demand

    return pd.Series(reach)


def optimize(demand, p, tester_config, init_stock=0):
    n = len(demand)
    x0 = calc_wafer_start(demand, p).values

    max_tester = tester_config["available"]
    target_reach = tester_config["target_REACH"]
    window = tester_config["reach_window"]
    start_idx = p["output_lag"]

    def objective(x):
        ws = pd.Series(x)
        stk = calc_stock(ws, demand, p, init_stock=init_stock)
        reach = calc_reach_4week_avg(
            stk,
            demand,
            window=window,
            start_idx=start_idx,
        )

        valid = ~np.isnan(reach.values)

        if not np.any(valid):
            return 1e12

        tester = calc_tester(ws, p).values
        reach_error = np.sum((reach.values[valid] - target_reach) ** 2)
        tester_level_error = np.sum((tester - tester.mean()) ** 2)
        tester_change_error = np.sum(np.diff(tester) ** 2)
        tester_jump_error = np.sum(np.maximum(np.abs(np.diff(tester)) - TESTER_JUMP_LIMIT, 0) ** 2)
        wafer_prev = np.maximum(ws.values[:-1], 1)
        wafer_change = np.abs(np.diff(ws.values)) / wafer_prev
        wafer_change_error = np.sum(np.maximum(wafer_change - WAFER_CHANGE_LIMIT, 0) ** 2)

        return float(
            reach_error
            + TESTER_LEVEL_WEIGHT * tester_level_error
            + TESTER_CHANGE_WEIGHT * tester_change_error
            + TESTER_JUMP_WEIGHT * tester_jump_error
            + WAFER_CHANGE_WEIGHT * wafer_change_error
        )

    constraints = [
        {
            "type": "ineq",
            "fun": lambda x: max_tester - calc_tester(pd.Series(x), p).values,
        },
        {
            "type": "ineq",
            "fun": lambda x: calc_stock(pd.Series(x), demand, p, init_stock=init_stock).values,
        },
    ]

    result = minimize(
        objective,
        x0,
        method="SLSQP",
        bounds=[(0, None)] * n,
        constraints=constraints,
        options={
            "maxiter": OPTIMIZER_MAXITER,
            "ftol": 1e-10,
            "disp": False,
        },
    )

    print(f"Optimizer success: {result.success}, iterations: {result.nit}, error: {result.fun:.4f}")
    return pd.Series(result.x)

def run_one_product(product, stages, products, weekly_demand_map, month_label_map, tester_config, inventory_map):
    print(f"Running product: {product}")
    weekly_demand = weekly_demand_map[product]
    months = month_label_map[product]
    init_stock = float(inventory_map.get(product, 0))

    p = get_total_params(stages, products, product)

    pre_weeks = p["total_lt"]

    demand_opt, months_ext, known_start_idx, forecast_start_idx = build_planning_horizon(
        weekly_demand,
        months,
        pre_weeks,
    )

    wafer_opt = optimize(demand_opt, p, tester_config, init_stock=init_stock)
    wafer_exec = pd.Series(np.rint(wafer_opt.values).astype(int))

    tester_used = pd.Series(
        np.ceil(calc_tester(wafer_exec, p).values).astype(int)
    )

    stock = calc_stock(
        wafer_exec,
        demand_opt,
        p,
        init_stock=init_stock,
        integer=True,
    )

    reach = calc_reach_4week_avg(
        stock.astype(float),
        demand_opt,
        window=tester_config["reach_window"],
        start_idx=known_start_idx,
    ).round(2)

    dps_out = calc_dps_out(wafer_exec, p, integer=True)
    bump_wafer = calc_shifted_output(wafer_exec, p["bump_lag"], integer=True)
    sort_wafer = calc_shifted_output(wafer_exec, p["sort_lag"], integer=True)
    dps_chip = calc_shifted_output(
        wafer_exec,
        p["dps_lag"],
        multiplier=p["cpw"] * p["dps_yield"],
        integer=True,
    )

    demand_disp_int = demand_opt.copy()
    demand_disp_int.iloc[:known_start_idx] = np.nan
    demand_disp_int = demand_disp_int.round().astype("Int64")

    plan = pd.DataFrame({
        "Product_Key": product,
        "Basic_Type": products[product].get("Basic_Type", ""),
        "Product": products[product].get("Original_Product", product),
        "Week": np.arange(1, len(demand_opt) + 1, dtype=int),
        "Month": months_ext,
        "Demand": demand_disp_int,
        "WaferStart": wafer_exec.astype(int),
        "Bump_Wafer": bump_wafer.astype(int),
        "Sort_Wafer": sort_wafer.astype(int),
        "DPS_Chip": dps_chip.astype(int),
        "TesterUsed": tester_used.astype(int),
        "DPS_Out_to_DC": dps_out.astype(int),
        "Stock": stock.astype(int),
        "REACH_4W_Avg": reach.values,
    })

    plan_output = plan.iloc[:forecast_start_idx].copy()
    plan_known = plan_output.iloc[known_start_idx:].copy()
    reach_valid = plan_known["REACH_4W_Avg"].dropna()

    summary = {
        "Product_Key": product,
        "Basic_Type": products[product].get("Basic_Type", ""),
        "Product": products[product].get("Original_Product", product),
        "Total_Weeks": len(plan_output),
        "Known_Demand_Weeks": len(plan_known),
        "Initial_Stock": int(round(init_stock)),
        "Total_Demand": int(plan_known["Demand"].fillna(0).sum()),
        "Total_WaferStart": int(plan_output["WaferStart"].sum()),
        "Max_TesterUsed": int(plan_output["TesterUsed"].max()),
        "Tester_Capacity": tester_config["available"],
        "Tester_OK": bool((plan_output["TesterUsed"] <= tester_config["available"]).all()),
        "Min_Stock": int(plan_output["Stock"].min()),
        "End_Stock": int(plan_output["Stock"].iloc[-1]),
        "REACH_Mean": float(reach_valid.mean()) if len(reach_valid) else np.nan,
        "REACH_Min": float(reach_valid.min()) if len(reach_valid) else np.nan,
        "REACH_Max": float(reach_valid.max()) if len(reach_valid) else np.nan,
    }

    monthly_summary = (
        plan_known
        .groupby(["Product_Key", "Basic_Type", "Product", "Month"], sort=False)
        .agg(
            Demand=("Demand", "sum"),
            Bump_Wafer=("Bump_Wafer", "sum"),
            Sort_Wafer=("Sort_Wafer", "sum"),
            DPS_Chip=("DPS_Chip", "sum"),
            Output=("DPS_Out_to_DC", "sum"),
            WaferStart=("WaferStart", "sum"),
            Max_TesterUsed=("TesterUsed", "max"),
            End_Stock=("Stock", "last"),
            Avg_REACH=("REACH_4W_Avg", "mean"),
        )
        .reset_index()
    )

    return plan_output, summary, monthly_summary


def run_all_products(stages, products, weekly_demand_map, month_label_map, tester_config, inventory_map):
    plans = []
    summaries = []
    monthly_summaries = []
    skipped_products = []

    product_keys = sorted(set(products.keys()) | set(weekly_demand_map.keys()))
    for product in product_keys:
        if product not in products:
            skipped_products.append({
                "Product_Key": product,
                "Reason": "Demand_Input 有需求，但 Parameter_Product 没有有效参数",
            })
            continue

        if product not in weekly_demand_map:
            skipped_products.append({
                "Product_Key": product,
                "Reason": "Parameter_Product 有参数，但 Demand_Input 无需求",
            })
            continue

        plan, summary, monthly_summary = run_one_product(
            product,
            stages,
            products,
            weekly_demand_map,
            month_label_map,
            tester_config,
            inventory_map,
        )

        plans.append(plan)
        summaries.append(summary)
        monthly_summaries.append(monthly_summary)

    if not plans:
        raise ValueError("没有任何 product 成功运行。")

    all_plan_df = pd.concat(plans, ignore_index=True)
    summary_df = pd.DataFrame(summaries)
    monthly_summary_df = pd.concat(monthly_summaries, ignore_index=True)
    skipped_df = pd.DataFrame(skipped_products)

    return all_plan_df, summary_df, monthly_summary_df, skipped_df

def write_output_excel(
    output_path: Path,
    all_plan_df: pd.DataFrame,
    summary_df: pd.DataFrame,
    monthly_summary_df: pd.DataFrame,
    skipped_df: pd.DataFrame,
    stages: dict,
    products: dict,
    tester_config: dict,
    inventory_map: dict,
):
    stage_df = pd.DataFrame([
        {
            "Stage": stage,
            "Cycle_Time_Week": params["cycle_time"],
            "Transit_Week": params["transit"],
            "Yield": params["yield"],
        }
        for stage, params in stages.items()
    ])

    product_df = pd.DataFrame([
        {
            "Product_Key": product,
            **params,
        }
        for product, params in products.items()
    ])

    tester_df = pd.DataFrame([
        {"Key": k, "Value": v}
        for k, v in tester_config.items()
    ])

    inventory_df = pd.DataFrame([
        {"Product_Key": product, "Initial_Stock": stock}
        for product, stock in sorted(inventory_map.items())
    ])

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        monthly_summary_df.to_excel(writer, sheet_name="Monthly_Summary", index=False)
        all_plan_df.to_excel(writer, sheet_name="All_Product_Plan", index=False)

        stage_df.to_excel(writer, sheet_name="Loaded_Flow", index=False)
        product_df.to_excel(writer, sheet_name="Loaded_Product", index=False)
        tester_df.to_excel(writer, sheet_name="Loaded_Tester_Config", index=False)
        inventory_df.to_excel(writer, sheet_name="Loaded_Inventory", index=False)

        if not skipped_df.empty:
            skipped_df.to_excel(writer, sheet_name="Skipped_Products", index=False)

    print(f"\nOutput saved to: {output_path}")

def main():
    print("Input Excel :", INPUT_EXCEL)
    print("Output Excel:", OUTPUT_EXCEL)

    stages, products, weekly_demand_map, month_label_map, tester_config, inventory_map = load_all_inputs(INPUT_EXCEL)

    all_plan_df, summary_df, monthly_summary_df, skipped_df = run_all_products(
        stages,
        products,
        weekly_demand_map,
        month_label_map,
        tester_config,
        inventory_map,
    )

    write_output_excel(
        OUTPUT_EXCEL,
        all_plan_df,
        summary_df,
        monthly_summary_df,
        skipped_df,
        stages,
        products,
        tester_config,
        inventory_map,
    )

    print("\n=== Summary ===")
    print(summary_df.to_string(index=False))

    if not skipped_df.empty:
        print("\n=== Skipped Products ===")
        print(skipped_df.to_string(index=False))


if __name__ == "__main__":
    main()