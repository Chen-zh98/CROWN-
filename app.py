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

    # 提取最新月和上个月的数据
    curr_df = monthly_sum[monthly_sum["月份"] == latest_period][["下单账号", "订单金额"]].rename(columns={"订单金额": "本月金额"})
    prev_df = monthly_sum[monthly_sum["月份"] == prev_period][["下单账号", "订单金额"]].rename(columns={"订单金额": "上月金额"})
    
    # 找出在“上个月及以前”曾经下过单的所有历史账号（用于区分纯新和老客）
    past_periods = [p for p in available_periods if p <= prev_period]
    past_df = monthly_sum[monthly_sum["月份"].isin(past_periods)][["下单账号"]].drop_duplicates()
    past_df["历史有单"] = True

    # 合并基础信息与对比数据（以基础信息为主表，确保流失的老客户也能显示出来）
    comp_df = base_info.merge(curr_df, on="下单账号", how="left")
    comp_df = comp_df.merge(prev_df, on="下单账号", how="left")
    comp_df = comp_df.merge(past_df, on="下单账号", how="left")

    comp_df["本月金额"] = comp_df["本月金额"].fillna(0)
    comp_df["上月金额"] = comp_df["上月金额"].fillna(0)
    comp_df["历史有单"] = comp_df["历史有单"].fillna(False)

    # 计算变化金额
    comp_df["变化金额"] = comp_df["本月金额"] - comp_df["上月金额"]

    # 状态判定逻辑
    def get_status(row):
        curr = row["本月金额"]
        prev = row["上月金额"]
        
        # 如果没有上传上个月的数据，则无法对比
        if prev_period not in available_periods:
            return "无上月数据"
            
        if prev == 0:
            if curr > 0:
                # 以前有单，但上月为0 -> 老客回流（新增金额）
                # 以前从来没单，本月有单 -> 全新账户（新增账户）
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
    
    # 重命名和整理列
    final_df = final_df.rename(columns={"上月金额": "上月下单金额", "本月金额": "本月下单金额"})
    
    # 把本月金额替换为对应的X月金额列名，避免冗余
    latest_month_col = f"{latest_period.month}月金额"
    if latest_month_col in final_df.columns:
        final_df = final_df.drop(columns=["本月下单金额"])

    # 排序：有最新单的按金额降序，没单的排后面
    if latest_month_col in final_df.columns:
        final_df = final_df.sort_values(by=[latest_month_col, "上月下单金额"], ascending=[False, False])

    # 整理列顺序（基础信息 -> 各月 -> 上月 -> 变化 -> 状态）
    cols = ["下单账号", "客户姓名", "实际电话", "地址"] + month_cols + ["上月下单金额", "变化金额", "状态"]
    final_df = final_df[[c for c in cols if c in final_df.columns]]

    return final_df, month_cols, latest_period

def style_final_row(row):
    """为表格整行上色，细化颜色"""
    status = row["状态"]
    color = ""
    if status == "全新账户": color = "#FFD966"       # 深黄色：纯新账户
    elif status == "老客回流": color = "#FFF2CC"     # 浅黄色：老号回流（新增金额）
    elif status == "增加": color = "#E2EFDA"         # 浅绿色
    elif status == "减少": color = "#FCE4D6"         # 浅红色
    elif status == "流失": color = "#F4B084"         # 橙红色
    elif status == "无上月数据": color = "#F2F2F2"   # 灰色
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
        
        # 整行涂色
        for row in range(2, ws.max_row + 1):
            status = ws.cell(row=row, column=status_col).value
            fill = fills.get(status)
            if fill:
                for col in range(1, ws.max_column + 1):
                    ws.cell(row=row, column=col).fill = fill

        # 列宽设置
        for col in range(1, ws.max_column + 1):
            col_name = get_column_letter(col)
            ws.column_dimensions[col_name].width = 15
            
        # 冻结窗格：冻结前 5 列和第一行
        ws.freeze_panes = "E2"
        
        # 金额格式化
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
        unsafe_allow_html=True
    )

uploaded_files = st.file_uploader("上传 Excel 文件，可一次选择多个月份", type=["xls", "xlsx"], accept_multiple_files=True)

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
