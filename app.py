import io
import re
import datetime

import pandas as pd
import streamlit as st
from openpyxl.styles import PatternFill, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="账号月度对比",
    page_icon="📊",
    layout="wide",
)

# 匹配表可提供的列（按顺序用于替换主表）
MATCH_COLS = ["客户姓名", "实际电话", "省份", "地级市", "区县", "地址"]

# 匹配表列名可能出现的别名
MATCH_COL_CANDIDATES = {
    "客户姓名": ["客户姓名", "姓名"],
    "实际电话": ["实际电话", "联系电话", "电话"],
    "省份": ["省份", "省"],
    "地级市": ["地级市", "市", "城市"],
    "区县": ["区县", "区", "县"],
    "地址": ["地址", "客户地址"],
}

# ---------------- 数据读取与处理 ----------------

def parse_year_month(filename: str, default_year: int):
    """从文件名识别年月"""
    name = filename
    m = re.search(r"(20\d{2})\s*[年\-_/]?\s*(\d{1,2})\s*月?", name)
    if m:
        y, mo = int(m.group(1)), int(m.group(2))
        if 1 <= mo <= 12:
            return y, mo

    m = re.search(r"(\d{1,2})\s*月", name)
    if m:
        mo = int(m.group(1))
        if 1 <= mo <= 12:
            return default_year, mo
    raise ValueError(f"无法从文件名识别月份：{filename}。请把文件名改成类似“6月万商.xls”")

def read_one_file(uploaded_file, default_year: int):
    """读取单个文件，提取基础信息和金额"""
    suffix = uploaded_file.name.lower().split(".")[-1]
    engine = "xlrd" if suffix == "xls" else "openpyxl"
    year, month = parse_year_month(uploaded_file.name, default_year)
    period = pd.Period(year=year, month=month, freq="M")

    sheets = pd.read_excel(uploaded_file, sheet_name=None, engine=engine)
    frames = []

    for sheet_name, df in sheets.items():
        if df is None or df.empty:
            continue

        df = df.copy()
        df.columns = [str(c).strip().replace("\ufeff", "") for c in df.columns]

        account_col = next((c for c in ["下单帐号", "下单账号"] if c in df.columns), None)
        if account_col is None or "订单金额" not in df.columns:
            continue

        name_col = next((c for c in ["客户姓名", "姓名"] if c in df.columns), None)
        phone_col = next((c for c in ["实际电话", "联系电话", "电话"] if c in df.columns), None)
        addr_col = next((c for c in ["地址", "客户地址"] if c in df.columns), None)

        out = pd.DataFrame()
        out["下单账号"] = df[account_col].astype(str).str.strip()
        out["订单金额"] = pd.to_numeric(
            df["订单金额"].astype(str).str.replace(",", "", regex=False).str.strip(),
            errors="coerce",
        )
        out["客户姓名"] = df[name_col].astype(str).str.strip() if name_col else ""
        out["实际电话"] = df[phone_col].astype(str).str.strip() if phone_col else ""
        out["地址"] = df[addr_col].astype(str).str.strip() if addr_col else ""

        mask = out["下单账号"].notna() & (out["下单账号"] != "") & (out["下单账号"].str.lower() != "nan") & out["订单金额"].notna()
        out = out[mask]

        if out.empty:
            continue

        out["月份"] = period
        frames.append(out)

    if not frames:
        raise ValueError(f"文件 {uploaded_file.name} 中未找到有效数据。")
    return pd.concat(frames, ignore_index=True)

def load_match_map(uploaded_file):
    """
    读取匹配表，返回 {下单账号: {客户姓名, 实际电话, 省份, 地级市, 区县, 地址}} 的字典。
    规则：若【实际电话】内容含有 * 或 ＊（全角），则替换为空值。
    """
    if uploaded_file is None:
        return {}

    suffix = uploaded_file.name.lower().split(".")[-1]
    engine = "xlrd" if suffix == "xls" else "openpyxl"

    try:
        sheets = pd.read_excel(uploaded_file, sheet_name=None, engine=engine)
    except Exception:
        uploaded_file.seek(0)
        only = pd.read_excel(uploaded_file, engine=engine)
        sheets = {"Sheet1": only}

    match_map = {}
    for _, df in sheets.items():
        if df is None or df.empty:
            continue

        df = df.copy()
        df.columns = [str(c).strip().replace("\ufeff", "") for c in df.columns]

        account_col = next((c for c in ["下单帐号", "下单账号"] if c in df.columns), None)
        if account_col is None:
            continue

        # 解析每个目标列在匹配表中的实际列名
        resolved = {}
        for target, candidates in MATCH_COL_CANDIDATES.items():
            resolved[target] = next((c for c in candidates if c in df.columns), None)

        # 至少要有一列可用信息
        if all(v is None for v in resolved.values()):
            continue

        for _, row in df.iterrows():
            acc = row[account_col]
            if pd.isna(acc):
                continue
            acc = str(acc).strip()
            if not acc or acc.lower() == "nan":
                continue

            info = {}
            for target, src in resolved.items():
                if src is None:
                    continue
                v = row[src]
                if pd.isna(v):
                    v = ""
                else:
                    v = str(v).strip()
                    if v.lower() == "nan":
                        v = ""
                info[target] = v

            # 电话中含 * 或 ＊ 一律置空
            if "实际电话" in info and re.search(r"[*＊]", info["实际电话"]):
                info["实际电话"] = ""

            match_map[acc] = info

    return match_map

def build_final_table(raw: pd.DataFrame):
    """生成最终的账号月度对比单表，细化状态"""
    raw = raw.sort_values("月份")
    base_info = raw.drop_duplicates("下单账号", keep="last")[["下单账号", "客户姓名", "实际电话", "地址"]]

    # 1. 各月金额宽表
    monthly_sum = raw.groupby(["月份", "下单账号"])["订单金额"].sum().reset_index()
    pivot = monthly_sum.pivot(index="下单账号", columns="月份", values="订单金额").fillna(0)
    pivot.columns = [f"{col.month}月金额" for col in pivot.columns]
    month_cols = sorted([c for c in pivot.columns if "月金额" in c], key=lambda x: int(x.replace("月金额", "")))
    pivot = pivot[month_cols].reset_index()

    # 2. 判断最新月份状态
    available_periods = set(monthly_sum["月份"].unique())
    latest_period = max(available_periods)
    prev_period = latest_period - 1

    curr_df = monthly_sum[monthly_sum["月份"] == latest_period][["下单账号", "订单金额"]].rename(columns={"订单金额": "本月金额"})
    prev_df = monthly_sum[monthly_sum["月份"] == prev_period][["下单账号", "订单金额"]].rename(columns={"订单金额": "上月金额"})

    past_periods = [p for p in available_periods if p <= prev_period]
    past_df = monthly_sum[monthly_sum["月份"].isin(past_periods)][["下单账号"]].drop_duplicates()
    past_df["历史有单"] = True

    comp_df = base_info.merge(curr_df, on="下单账号", how="left")
    comp_df = comp_df.merge(prev_df, on="下单账号", how="left")
    comp_df = comp_df.merge(past_df, on="下单账号", how="left")

    comp_df["本月金额"] = comp_df["本月金额"].fillna(0)
    comp_df["上月金额"] = comp_df["上月金额"].fillna(0)
    comp_df["历史有单"] = comp_df["历史有单"].fillna(False)

    comp_df["变化金额"] = comp_df["本月金额"] - comp_df["上月金额"]

    def get_status(row):
        curr = row["本月金额"]
        prev = row["上月金额"]

        if prev_period not in available_periods:
            return "无上月数据"

        if prev == 0:
            if curr > 0:
                return "老客回流" if row["历史有单"] else "全新账户"
            else:
                return "无单"
        else:
            if curr == 0:
                return "流失"
            diff = curr - prev
            if diff > 0: return "增加"
            if diff < 0: return "减少"
            return "持平"

    comp_df["状态"] = comp_df.apply(get_status, axis=1)

    # 3. 合并宽表
    final_df = pd.merge(comp_df, pivot, on="下单账号", how="outer")

    final_df = final_df.rename(columns={"上月金额": "上月下单金额", "本月金额": "本月下单金额"})

    latest_month_col = f"{latest_period.month}月金额"
    if latest_month_col in final_df.columns:
        final_df = final_df.drop(columns=["本月下单金额"])

    if latest_month_col in final_df.columns:
        final_df = final_df.sort_values(by=[latest_month_col, "上月下单金额"], ascending=[False, False])

    # 确保匹配列存在于最终表（无则补空）
    for c in ["省份", "地级市", "区县"]:
        if c not in final_df.columns:
            final_df[c] = ""

    ordered_cols = (
        ["下单账号", "客户姓名", "实际电话", "省份", "地级市", "区县", "地址"]
        + month_cols
        + ["上月下单金额", "变化金额", "状态"]
    )
    final_df = final_df[[c for c in ordered_cols if c in final_df.columns]]

    final_df = final_df.reset_index(drop=True)

    # ⭐️ 全局兜底清理：强制把主表中带有 * 或 ＊ 的电话替换为空
    if "实际电话" in final_df.columns:
        mask_star = final_df["实际电话"].astype(str).str.contains(r"[*＊]", regex=True, na=False)
        final_df.loc[mask_star, "实际电话"] = ""

    return final_df, month_cols, latest_period

def style_final_row(row):
    status = row["状态"]
    color = ""
    if status == "全新账户": color = "#FFD966"
    elif status == "老客回流": color = "#FFF2CC"
    elif status == "增加": color = "#E2EFDA"
    elif status == "减少": color = "#FCE4D6"
    elif status == "流失": color = "#F4B084"
    elif status == "无上月数据": color = "#F2F2F2"
    return [f"background-color: {color}" if color else "" for _ in row]

def to_excel_bytes(final_df):
    """导出为单张 Excel 表格，并附带颜色格式"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        final_df.to_excel(writer, sheet_name="账号月度对比", index=False)
        wb = writer.book
        ws = wb["账号月度对比"]

        headers = {cell.value: cell.column for cell in ws[1]}
        status_col = headers["状态"]

        fills = {
            "全新账户": PatternFill(fill_type="solid", start_color="FFD966", end_color="FFD966"),
            "老客回流": PatternFill(fill_type="solid", start_color="FFF2CC", end_color="FFF2CC"),
            "增加": PatternFill(fill_type="solid", start_color="E2EFDA", end_color="E2EFDA"),
            "减少": PatternFill(fill_type="solid", start_color="FCE4D6", end_color="FCE4D6"),
            "流失": PatternFill(fill_type="solid", start_color="F4B084", end_color="F4B084"),
            "无上月数据": PatternFill(fill_type="solid", start_color="F2F2F2", end_color="F2F2F2"),
        }

        for row in range(2, ws.max_row + 1):
            status = ws.cell(row=row, column=status_col).value
            fill = fills.get(status)
            if fill:
                for col in range(1, ws.max_column + 1):
                    ws.cell(row=row, column=col).fill = fill

        # 列宽优化
        width_map = {
            "下单账号": 22,
            "客户姓名": 12,
            "实际电话": 16,
            "省份": 10,
            "地级市": 12,
            "区县": 12,
            "地址": 32,
            "状态": 12,
        }
        for col_name, col_idx in headers.items():
            ws.column_dimensions[get_column_letter(col_idx)].width = width_map.get(str(col_name), 15)

        ws.freeze_panes = "E2"

        for col_name, col_idx in headers.items():
            if "金额" in str(col_name):
                for row in range(2, ws.max_row + 1):
                    ws.cell(row=row, column=col_idx).number_format = "#,##0.00"

    output.seek(0)
    return output.getvalue()

# ---------------- UI 界面 ----------------

st.title("📊 账号月度对比工具")
st.caption("横向展示各月下单金额，并在右侧精细化对比状态（区分新老账户与金额涨跌）。")

with st.sidebar:
    st.header("参数设置")
    default_year = st.number_input(
        "文件名只有“X月”时，使用的年份", min_value=2000, max_value=2100, value=datetime.date.today().year, step=1
    )

    st.markdown("---")
    st.markdown("### 颜色图例说明")
    st.markdown(
        """
        <div style='background-color:#FFD966; padding:5px; margin-bottom:5px; border-radius:4px; color:black;'><b>深黄色</b>：全新账户（历史从未下过单）</div>
        <div style='background-color:#FFF2CC; padding:5px; margin-bottom:5px; border-radius:4px; color:black;'><b>浅黄色</b>：老客回流（上月断单，本月重新下单）</div>
        <div style='background-color:#E2EFDA; padding:5px; margin-bottom:5px; border-radius:4px; color:black;'><b>浅绿色</b>：正常增加</div>
        <div style='background-color:#FCE4D6; padding:5px; margin-bottom:5px; border-radius:4px; color:black;'><b>浅红色</b>：正常减少</div>
        <div style='background-color:#F4B084; padding:5px; margin-bottom:5px; border-radius:4px; color:black;'><b>橙红色</b>：流失（上月有单，本月归零）</div>
        """,
        unsafe_allow_html=True,
    )

uploaded_files = st.file_uploader(
    "上传月度 Excel 文件，可一次选择多个月份", type=["xls", "xlsx"], accept_multiple_files=True
)
match_file = st.file_uploader(
    "（可选）上传【信息匹配表】，需包含【下单账号】列，可包含"
    "【客户姓名】【实际电话】【省份】【地级市】【区县】【地址】列；"
    "匹配表中存在的列将按账号覆盖主表；【实际电话】含 * 或 ＊ 的会自动置空。",
    type=["xls", "xlsx"],
)

if uploaded_files:
    all_rows = []
    for f in uploaded_files:
        try:
            df = read_one_file(f, int(default_year))
            all_rows.append(df)
        except Exception as e:
            st.error(f"{f.name} 读取失败：{e}")

    if all_rows:
        raw = pd.concat(all_rows, ignore_index=True)
        final_df, month_cols, latest_period = build_final_table(raw)

        # ---------- 应用匹配表 ----------
        if match_file is not None:
            try:
                match_map = load_match_map(match_file)
            except Exception as e:
                st.warning(f"匹配表读取失败：{e}")
                match_map = {}

            if match_map:
                acc_series = final_df["下单账号"].astype(str).str.strip()
                matched_acc = 0
                field_hits = {c: 0 for c in MATCH_COLS}

                for acc, info in match_map.items():
                    mask = acc_series == acc
                    if not mask.any():
                        continue
                    matched_acc += int(mask.sum())
                    for col, val in info.items():
                        if col in final_df.columns:
                            final_df.loc[mask, col] = val
                            field_hits[col] += int(mask.sum())

                if matched_acc > 0:
                    hit_desc = "、".join(f"{c}×{n}" for c, n in field_hits.items() if n > 0) or "无字段"
                    st.success(
                        f"✅ 已按匹配表更新 {matched_acc} 个账号的字段（{hit_desc}），含 * 的电话已置空。"
                    )
                else:
                    st.info("匹配表中未找到匹配的【下单账号】。")
            else:
                st.info("匹配表未读取到有效数据（需含【下单账号】列且至少有一列可用信息）。")

        # ⭐️ 再次兜底：防止匹配表的内容里混入星号（双重保险）
        if "实际电话" in final_df.columns:
            mask_star = final_df["实际电话"].astype(str).str.contains(r"[*＊]", regex=True, na=False)
            final_df.loc[mask_star, "实际电话"] = ""

        st.subheader(f"账号月度对比 (对比月份: {latest_period})")

        format_dict = {col: "{:,.2f}" for col in month_cols}
        format_dict.update({"上月下单金额": "{:,.2f}", "变化金额": "{:,.2f}"})

        styled = final_df.style.apply(style_final_row, axis=1).format(format_dict)
        st.dataframe(styled, use_container_width=True, height=600)

        excel_bytes = to_excel_bytes(final_df)
        st.download_button(
            label="⬇️ 下载【账号月度对比】Excel",
            data=excel_bytes,
            file_name="账号月度对比_精细版.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("请先上传 Excel 文件。")
