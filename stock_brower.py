# ----------------- 存取密碼保護機制 -----------------
def check_password():
    """驗證通過回傳 True，否則顯示輸入框並阻斷後續程式"""
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False

    if st.session_state.authenticated:
        return True

    st.title("🔒 台股家庭資產管理系統")
    st.caption("此為家庭私人資產管理平台，請輸入授權通行碼。")
    
    pwd_input = st.text_input("請輸入訪問密碼：", type="password", key="login_password_input")
    
    if st.button("確認解鎖", type="primary"):
        expected_pwd = st.secrets.get("app_password", "123456")
        if pwd_input == expected_pwd:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("❌ 密碼錯誤，拒絕存取！")
            
    return False

# 若未通過驗證，立即終止後續所有持股與 Google Sheets 運算
if not check_password():
    st.stop()
import streamlit as st
import pandas as pd
import yfinance as yf
import os
import io
import requests
import time
from datetime import datetime, date

# Google Sheets API 連線套件
import gspread
from google.oauth2.service_account import Credentials

from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

st.set_page_config(page_title="台股家庭資產管理系統", layout="wide")

# 1. 強制置頂樣式
st.markdown("""
<style>
    * {
        font-weight: normal !important;
    }
    header[data-testid="stHeader"] {
        z-index: 1000;
    }
    .stApp > header {
        background-color: #ffffff;
    }
    div[data-testid="stTabs"] > div:first-child {
        position: sticky;
        top: 3.5rem;
        background-color: #ffffff;
        z-index: 999;
        padding: 10px 0;
        border-bottom: 2px solid #1890ff;
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.06);
    }
    [data-testid="stMetricValue"] {
        font-size: clamp(1.15rem, 1.5vw, 1.7rem) !important;
    }
    [data-testid="stMetricLabel"] {
        font-size: clamp(0.8rem, 0.9vw, 0.95rem) !important;
    }
    .stDataFrame {
        font-size: clamp(13px, 0.9vw, 15px) !important;
    }
    div[data-testid="stDataFrame"] div[role="row"]:nth-child(even) {
        background-color: #f8f9fa !important;
    }
    div[data-testid="stDataFrame"] div[role="row"]:nth-child(odd) {
        background-color: #ffffff !important;
    }
    div.stButton > button,
    button[kind="primary"],
    button[kind="secondary"] {
        background-color: #1890ff !important;
        color: #ffffff !important;
        border: 1px solid #1890ff !important;
        border-radius: 6px !important;
    }
    div.stButton > button:hover {
        background-color: #096dd9 !important;
        color: #ffffff !important;
    }
</style>
""", unsafe_allow_html=True)

# ----------------- 日期永遠保持 YYYY/MM/DD (固定 10 碼) -----------------
def normalize_date(val):
    if not val or pd.isna(val):
        return datetime.now().strftime("%Y/%m/%d")
    s = str(val).strip().replace("-", "/")
    parts = s.split("/")
    if len(parts) == 3:
        try:
            y = int(parts[0])
            m = int(parts[1])
            d = int(parts[2])
            return f"{y:04d}/{m:02d}/{d:02d}"
        except Exception:
            pass
    return s

# ----------------- 台股代號精確清洗 (徹底防呆前導0) -----------------
def clean_stock_code(val):
    s = str(val).strip().upper()
    if not s:
        return ""
    if s.isdigit():
        num_val = int(s)
        if num_val in [919, 929, 940, 878, 50, 56, 830]:
            return f"00{num_val}"
        elif len(s) == 3:
            return s.zfill(5)
        elif len(s) == 4 and s.startswith("0") and s in ["0919", "0929", "0940", "0878"]:
            return "0" + s
        elif len(s) < 4:
            return s.zfill(4)
    return s

BASE_STOCKS = {
    "0050": "元大台灣50", "0056": "元大高股息", "00878": "國泰永續高股息",
    "00919": "群益台灣精選高息", "00929": "復華台灣科技優息", "00940": "元大台灣價值高息",
    "00631L": "元大台灣50正2", "00830": "國泰費城半導體", "00725B": "國泰投資級公司債",
    "2417": "圓剛", "2330": "台積電", "2317": "鴻海", "2454": "聯發科",
    "2449": "京元電子", "2337": "旺宏", "6182": "合晶", "2886": "兆豐金",
    "2881": "富邦金", "2882": "國泰金", "2891": "中信金", "2884": "玉山金",
    "2880": "華南金", "5880": "合庫金", "2603": "長榮", "3231": "緯創",
    "2382": "廣達", "2308": "台達電", "3711": "日月光投控"
}

@st.cache_data(ttl=86400)
def query_stock_name(symbol: str) -> str:
    code = clean_stock_code(symbol)
    if code in BASE_STOCKS:
        return BASE_STOCKS[code]
    try:
        url = f"https://mis.twse.com.tw/stock/api/getStockInfo.jsp?ex_ch=tse_{code}.tw|otc_{code}.two"
        res = requests.get(url, timeout=2).json()
        items = res.get("msgArray", [])
        if items and "n" in items[0]:
            name = items[0]["n"].strip()
            if name:
                BASE_STOCKS[code] = name
                return name
    except Exception:
        pass
    return code

# ----------------- Google Sheets API 連線與安全讀寫 -----------------
@st.cache_resource
def get_gspread_client():
    if "gcp_service_account" in st.secrets:
        creds_dict = dict(st.secrets["gcp_service_account"])
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        credentials = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        return gspread.authorize(credentials)
    return None

def get_spreadsheet():
    client = get_gspread_client()
    if client and "spreadsheet_url" in st.secrets:
        try:
            return client.open_by_url(st.secrets["spreadsheet_url"])
        except Exception:
            return None
    return None

@st.cache_data(ttl=5)
def read_sheet(sheet_name: str) -> pd.DataFrame:
    ss = get_spreadsheet()
    if ss:
        try:
            ws = ss.worksheet(sheet_name)
            records = ws.get_all_values()
            if len(records) > 1:
                header = [str(c).strip() for c in records[0]]
                rows = records[1:]
                return pd.DataFrame(rows, columns=header)
            return pd.DataFrame()
        except Exception:
            return pd.DataFrame()
    csv_file = f"{sheet_name}.csv"
    if os.path.exists(csv_file):
        return pd.read_csv(csv_file, dtype=str)
    return pd.DataFrame()

def write_sheet(sheet_name: str, df: pd.DataFrame):
    st.cache_data.clear()
    ss = get_spreadsheet()
    if ss:
        try:
            ws = ss.worksheet(sheet_name)
            ws.clear()
            df_to_write = df.astype(str).fillna("")
            ws.update([df_to_write.columns.values.tolist()] + df_to_write.values.tolist())
            return
        except Exception as e:
            st.error(f"雲端寫入失敗：{e}")
    df.to_csv(f"{sheet_name}.csv", index=False, encoding="utf-8-sig")

# ----------------- 基本檔讀取與補齊欄位防呆 -----------------
def load_brokers():
    df = read_sheet("brokers")
    if not df.empty and "券商名稱" in df.columns:
        return [str(x).strip() for x in df["券商名稱"].dropna().tolist() if str(x).strip()]
    init_b = ["元大證券", "國泰證券", "富邦證券", "永豐金證券", "群益證券", "凱基證券", "玉山證券"]
    write_sheet("brokers", pd.DataFrame({"券商名稱": init_b}))
    return init_b

def save_brokers(b_list):
    write_sheet("brokers", pd.DataFrame({"券商名稱": b_list}))

def load_members():
    df = read_sheet("members")
    if not df.empty and "成員名稱" in df.columns:
        return [str(x).strip() for x in df["成員名稱"].dropna().tolist() if str(x).strip()]
    init_m = ["本人", "成員1", "成員2", "成員3"]
    write_sheet("members", pd.DataFrame({"成員名稱": init_m}))
    return init_m

def save_members(m_list):
    write_sheet("members", pd.DataFrame({"成員名稱": m_list}))

def load_stock_settings():
    df = read_sheet("stock_settings")
    defaults = {
        "0050": "半年配", "0056": "季配", "00878": "季配", "00919": "季配", 
        "00929": "月配", "00940": "月配", "00725B": "季配", "2330": "季配", 
        "2417": "年配", "2884": "年配"
    }
    if not df.empty and "代號" in df.columns:
        mapping = {}
        for _, r in df.iterrows():
            c = clean_stock_code(r["代號"])
            mapping[c] = str(r["頻率"]).strip()
        return mapping
    write_sheet("stock_settings", pd.DataFrame(list(defaults.items()), columns=["代號", "頻率"]))
    return defaults

def save_stock_settings(settings):
    df = pd.DataFrame(list(settings.items()), columns=["代號", "頻率"])
    df["代號"] = df["代號"].apply(clean_stock_code)
    write_sheet("stock_settings", df)

# 交易資料讀取：確保包含所有必備欄位
def load_transactions():
    df = read_sheet("transactions")
    req_cols = ["id", "日期", "成員", "券商", "代號", "名稱", "類別", "股數", "價格", "頻率"]
    if df.empty:
        return pd.DataFrame(columns=req_cols)

    # 欄位防呆補齊
    for col in req_cols:
        if col not in df.columns:
            if col == "成員":
                df[col] = "本人"
            elif col == "券商":
                df[col] = "元大證券"
            elif col == "頻率":
                df[col] = "季配"
            elif col == "類別":
                df[col] = "買進"
            else:
                df[col] = ""

    df["代號"] = df["代號"].apply(clean_stock_code)
    df["日期"] = df["日期"].apply(normalize_date)
    if "id" not in df.columns or df["id"].isnull().all():
        df["id"] = range(1, len(df) + 1)

    for idx, r in df.iterrows():
        c = r["代號"]
        n = str(r["名稱"]).strip()
        if not n or n.isupper() or n == c or c in BASE_STOCKS:
            df.at[idx, "名稱"] = query_stock_name(c)
    return df

def save_all_transactions(df):
    df["代號"] = df["代號"].apply(clean_stock_code)
    df["日期"] = df["日期"].apply(normalize_date)
    write_sheet("transactions", df)

# 收益資料讀取：核心防呆（徹底避免 KeyError: '成員'）
def load_income_records():
    df = read_sheet("income_records")
    req_cols = ["id", "日期", "年度", "成員", "券商", "代號", "名稱", "項目", "金額", "獲配股數", "備註"]
    if df.empty:
        return pd.DataFrame(columns=req_cols)

    # 若工作表缺少必要欄位，自動防呆補上
    for col in req_cols:
        if col not in df.columns:
            if col == "成員":
                df[col] = "本人"
            elif col == "券商":
                df[col] = "元大證券"
            elif col == "項目":
                df[col] = "現金股利"
            elif col in ["金額", "獲配股數"]:
                df[col] = "0"
            else:
                df[col] = ""

    df["代號"] = df["代號"].apply(clean_stock_code)
    df["日期"] = df["日期"].apply(normalize_date)
    if "id" not in df.columns or df["id"].isnull().all():
        df["id"] = range(1, len(df) + 1)
    return df

def save_income_records(df):
    df["代號"] = df["代號"].apply(clean_stock_code)
    df["日期"] = df["日期"].apply(normalize_date)
    write_sheet("income_records", df)

# ----------------- 市價精確查詢 -----------------
@st.cache_data(ttl=60)
def fetch_stock_market_data(symbol: str):
    code = clean_stock_code(symbol)
    candidates = [f"{code}.TW", f"{code}.TWO"]
    price = 0.0
    last_dividend = 0.0

    for target in candidates:
        try:
            ticker = yf.Ticker(target)
            try:
                fi = ticker.fast_info
                p = getattr(fi, "last_price", None) or fi.get("lastPrice", None)
                if p and float(p) > 0:
                    price = float(p)
            except Exception:
                pass

            if price <= 0:
                try:
                    intraday = ticker.history(period="1d", interval="1m")
                    if not intraday.empty:
                        price = float(intraday["Close"].dropna().iloc[-1])
                except Exception:
                    pass

            if price <= 0:
                try:
                    hist = ticker.history(period="5d")
                    if not hist.empty:
                        price = float(hist["Close"].dropna().iloc[-1])
                except Exception:
                    pass

            try:
                divs = ticker.dividends
                if not divs.empty:
                    last_dividend = float(divs.dropna().iloc[-1])
            except Exception:
                pass

            if price > 0:
                break
        except Exception:
            continue

    return round(price, 2), round(last_dividend, 2)

FREQ_MULTIPLIER = {"月配": 12, "季配": 4, "半年配": 2, "年配": 1, "不配息": 0}

def calculate_portfolio(trans_df, stock_settings):
    holdings = {}
    realized_records = []

    if trans_df.empty:
        return pd.DataFrame(), pd.DataFrame()

    t_df = trans_df.copy()
    t_df["代號"] = t_df["代號"].apply(clean_stock_code)
    trans_sorted = t_df.sort_values(by=["日期", "id"]).to_dict("records")

    for t in trans_sorted:
        code = clean_stock_code(t["代號"])
        shares = float(pd.to_numeric(t["股數"], errors="coerce") or 0)
        price = float(pd.to_numeric(t["價格"], errors="coerce") or 0)
        action = str(t["類別"]).strip()
        name = str(t.get("名稱", "")).strip() or query_stock_name(code)
        freq = t.get("頻率", stock_settings.get(code, "季配"))
        broker = t.get("券商", "元大證券")
        member = t.get("成員", "本人")

        if code not in holdings:
            holdings[code] = {
                "名稱": name,
                "庫存股數": 0.0,
                "總成本": 0.0,
                "平均成本": 0.0,
                "頻率": freq
            }
        else:
            if name and name != code:
                holdings[code]["名稱"] = name
            holdings[code]["頻率"] = freq

        if action == "買進":
            holdings[code]["總成本"] += shares * price
            holdings[code]["庫存股數"] += shares
            if holdings[code]["庫存股數"] > 0:
                holdings[code]["平均成本"] = holdings[code]["總成本"] / holdings[code]["庫存股數"]
        elif action == "配股":
            holdings[code]["庫存股數"] += shares
            if holdings[code]["庫存股數"] > 0:
                holdings[code]["平均成本"] = holdings[code]["總成本"] / holdings[code]["庫存股數"]
        elif action == "賣出":
            avg_cost = holdings[code]["平均成本"]
            cost_of_sold = shares * avg_cost
            revenue = shares * price
            realized_pnl = revenue - cost_of_sold

            realized_records.append({
                "日期": normalize_date(t["日期"]),
                "成員": member,
                "券商": broker,
                "代號": code,
                "名稱": holdings[code]["名稱"],
                "類別": "賣出",
                "賣出股數": shares,
                "賣出價格": price,
                "成本單價": round(avg_cost, 2),
                "總成本": round(cost_of_sold, 0),
                "實現損益": round(realized_pnl, 0)
            })

            holdings[code]["庫存股數"] -= shares
            holdings[code]["總成本"] -= cost_of_sold
            if holdings[code]["庫存股數"] <= 0:
                holdings[code]["庫存股數"] = 0.0
                holdings[code]["總成本"] = 0.0
                holdings[code]["平均成本"] = 0.0

    portfolio_list = []
    for code, data in holdings.items():
        if data["庫存股數"] > 0:
            portfolio_list.append({
                "代號": code,
                "名稱": data["名稱"],
                "庫存股數": data["庫存股數"],
                "平均成本": round(data["平均成本"], 2),
                "總成本": round(data["總成本"], 0),
                "頻率": stock_settings.get(code, data["頻率"])
            })

    return pd.DataFrame(portfolio_list), pd.DataFrame(realized_records)

def style_pnl_value(val):
    try:
        num = float(val)
        if num > 0:
            return 'color: #ff4d4f;'
        elif num < 0:
            return 'color: #52c41a;'
    except Exception:
        pass
    return ''

def style_action_cell(val):
    s_val = str(val).strip()
    if s_val == "賣出":
        return 'background-color: #f6ffed; color: #389e0d;'
    elif s_val in ["配股", "股票股利"]:
        return 'background-color: #fff7e6; color: #d4380d;'
    return ''

# ==================== 主頁面控制區 ====================
member_list = load_members()
broker_list = load_brokers()
raw_trans_df = load_transactions()
stock_settings = load_stock_settings()
raw_income_df = load_income_records()

col_head1, col_head2 = st.columns([6, 4])
with col_head1:
    st.title("💼 台股家庭資產管理系統")
    if get_spreadsheet():
        st.caption("☁️ 雲端狀態：已連線至 Google Sheets 即時同步 (安全直連)")
    else:
        st.caption("💻 雲端狀態：使用本機 CSV 運行")

with col_head2:
    all_member_options = ["全部家庭成員 (合併總覽)"] + member_list
    current_selected_member = st.selectbox(
        "👤 目前檢視成員帳戶：",
        options=all_member_options,
        index=0,
        key="active_member_selector"
    )

if current_selected_member == "全部家庭成員 (合併總覽)":
    trans_df = raw_trans_df.copy()
    income_df = raw_income_df.copy()
else:
    trans_df = raw_trans_df[raw_trans_df["成員"] == current_selected_member].copy()
    income_df = raw_income_df[raw_income_df["成員"] == current_selected_member].copy()

portfolio_df, realized_df = calculate_portfolio(trans_df, stock_settings)

# 動態建立乾淨的代號候選庫
dynamic_stock_pool = {}
for code, name in BASE_STOCKS.items():
    dynamic_stock_pool[clean_stock_code(code)] = name

for _, r in raw_trans_df.iterrows():
    c_code = clean_stock_code(r["代號"])
    c_name = str(r["名稱"]).strip()
    if c_code and c_code not in dynamic_stock_pool:
        dynamic_stock_pool[c_code] = c_name

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📈 現有持股與個股穿透檢視", 
    "📝 交易記帳與明細編輯", 
    "💰 已實現賣出損益", 
    "🎁 實際配息與配股登記", 
    "📋 自訂報表匯出",
    "👥 基本資料設定與維護"
])

# ----------------- TAB 1: 持股與穿透檢視 -----------------
with tab1:
    if current_selected_member != "全部家庭成員 (合併總覽)":
        st.info(f"📌 目前正在檢視：【{current_selected_member}】 的專屬投資組合")

    if portfolio_df.empty:
        st.info(f"【{current_selected_member}】目前無在庫持股，請至「交易記帳」登錄買進交易。")
    else:
        with st.spinner("同步市場即時現價中..."):
            mkt_prices, recent_divs, annual_divs, mkt_values = [], [], [], []
            unrealized_pnls, unrealized_pnl_rates, div_yields = [], [], []

            for _, row in portfolio_df.iterrows():
                code = clean_stock_code(row["代號"])
                price, last_div = fetch_stock_market_data(code)
                current_price = price if price > 0 else row["平均成本"]

                freq = row["頻率"]
                multiplier = FREQ_MULTIPLIER.get(freq, 4)
                est_annual_per_share = last_div * multiplier
                est_total_dividend = est_annual_per_share * row["庫存股數"]

                mkt_val = current_price * row["庫存股數"]
                pnl = mkt_val - row["總成本"]
                pnl_rate = (pnl / row["總成本"] * 100) if row["總成本"] > 0 else 0.0
                div_yield = (est_annual_per_share / current_price * 100) if current_price > 0 else 0.0

                mkt_prices.append(current_price)
                recent_divs.append(last_div)
                annual_divs.append(est_total_dividend)
                mkt_values.append(mkt_val)
                unrealized_pnls.append(pnl)
                unrealized_pnl_rates.append(pnl_rate)
                div_yields.append(div_yield)

            portfolio_df["目前現價"] = mkt_prices
            portfolio_df["最近配息"] = recent_divs
            portfolio_df["目前市值"] = mkt_values
            portfolio_df["未實現損益"] = unrealized_pnls
            portfolio_df["報酬率(%)"] = unrealized_pnl_rates
            portfolio_df["預估年股息"] = annual_divs
            portfolio_df["預估殖利率(%)"] = div_yields

        tot_mkt = portfolio_df["目前市值"].sum()
        tot_cost = portfolio_df["總成本"].sum()
        tot_pnl = tot_mkt - tot_cost
        tot_pnl_pct = (tot_pnl / tot_cost * 100) if tot_cost > 0 else 0.0
        tot_div = portfolio_df["預估年股息"].sum()
        tot_yield = (tot_div / tot_mkt * 100) if tot_mkt > 0 else 0.0

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("持股總市值", f"NT$ {tot_mkt:,.0f}")
        c2.metric("總投入成本", f"NT$ {tot_cost:,.0f}")
        c3.metric("未實現損益", f"NT$ {tot_pnl:,.0f}", delta=f"{tot_pnl_pct:+.2f}%")
        c4.metric("預估年度總股息", f"NT$ {tot_div:,.0f}", delta=f"殖利率 {tot_yield:.2f}%")

        st.divider()
        st.subheader("目前持股總覽")

        disp_p = portfolio_df[[
            "代號", "名稱", "庫存股數", "平均成本", "總成本", "目前現價", "目前市值",
            "未實現損益", "報酬率(%)", "頻率", "最近配息", "預估年股息", "預估殖利率(%)"
        ]].copy()
        disp_p.insert(0, "序號", range(1, len(disp_p) + 1))

        styled_p = disp_p.style.format({
            "序號": "{:d}",
            "庫存股數": "{:,.0f}",
            "平均成本": "NT$ {:,.2f}",
            "總成本": "NT$ {:,.0f}",
            "目前現價": "NT$ {:,.2f}",
            "目前市值": "NT$ {:,.0f}",
            "未實現損益": "NT$ {:+,.0f}",
            "報酬率(%)": "{:+.2f}%",
            "最近配息": "{:,.2f}",
            "預估年股息": "NT$ {:,.0f}",
            "預估殖利率(%)": "{:.2f}%"
        }).map(style_pnl_value, subset=["未實現損益", "報酬率(%)"])

        st.dataframe(styled_p, use_container_width=True, hide_index=True)

        st.divider()
        st.subheader("🔍 個股履歷穿透檢視")

        stock_options = [f"{clean_stock_code(r['代號'])} {r['名稱']}" for _, r in portfolio_df.iterrows()]
        selected_stock_str = st.selectbox(
            "請選擇要檢視履歷的股票：",
            options=stock_options,
            index=0,
            key="drill_select_fixed"
        )

        target_drill_code = clean_stock_code(selected_stock_str.split()[0])
        target_drill_name = selected_stock_str.split()[1]
        cur_stock_row = portfolio_df[portfolio_df["代號"] == target_drill_code].iloc[0]

        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("現有庫存", f"{cur_stock_row['庫存股數']:,} 股")
        k2.metric("平均成本", f"NT$ {cur_stock_row['平均成本']:,.2f}")
        k3.metric("目前現價", f"NT$ {cur_stock_row['目前現價']:,.2f}")
        k4.metric("目前市值", f"NT$ {cur_stock_row['目前市值']:,.0f}")
        k5.metric("未實現損益", f"NT$ {cur_stock_row['未實現損益']:+,.0f}", delta=f"{cur_stock_row['報酬率(%)']:+.2f}%")

        stock_txs = trans_df[trans_df["代號"] == target_drill_code].sort_values(by="日期", ascending=False).copy()
        if not stock_txs.empty:
            stock_txs.insert(0, "序號", range(1, len(stock_txs) + 1))

        stock_incs = income_df[income_df["代號"] == target_drill_code].sort_values(by="日期", ascending=False).copy()
        if not stock_incs.empty:
            stock_incs.insert(0, "序號", range(1, len(stock_incs) + 1))

        col_sub_tx, col_sub_inc = st.columns(2)
        with col_sub_tx:
            st.write("📜 **買賣與除權流水帳**")
            if stock_txs.empty:
                st.info("無交易明細")
            else:
                disp_stk_tx = stock_txs[["序號", "日期", "成員", "券商", "類別", "股數", "價格"]].copy()
                disp_stk_tx["日期"] = disp_stk_tx["日期"].apply(normalize_date)
                disp_stk_tx["股數"] = pd.to_numeric(disp_stk_tx["股數"], errors="coerce").fillna(0)
                disp_stk_tx["價格"] = pd.to_numeric(disp_stk_tx["價格"], errors="coerce").fillna(0)
                styled_stk_tx = disp_stk_tx.style.format({
                    "序號": "{:d}",
                    "股數": lambda v: f"{float(v):,.0f}",
                    "價格": lambda v: f"NT$ {float(v):,.2f}"
                }).map(style_action_cell, subset=["類別"])
                st.dataframe(styled_stk_tx, use_container_width=True, hide_index=True)

        with col_sub_inc:
            st.write("💰 **實際收益明細**")
            if stock_incs.empty:
                st.info("尚無收息紀錄")
            else:
                # 解決 KeyError 的安全取欄位邏輯
                sub_cols = ["序號", "日期", "項目", "金額", "獲配股數"]
                if "成員" in stock_incs.columns:
                    sub_cols.insert(2, "成員")
                if "券商" in stock_incs.columns:
                    sub_cols.insert(3, "券商")
                if "備註" in stock_incs.columns:
                    sub_cols.append("備註")

                disp_stk_inc = stock_incs[sub_cols].copy()
                disp_stk_inc["日期"] = disp_stk_inc["日期"].apply(normalize_date)
                disp_stk_inc["金額"] = pd.to_numeric(disp_stk_inc["金額"], errors="coerce").fillna(0)
                disp_stk_inc["獲配股數"] = pd.to_numeric(disp_stk_inc["獲配股數"], errors="coerce").fillna(0)
                styled_stk_inc = disp_stk_inc.style.format({
                    "序號": "{:d}",
                    "金額": lambda v: f"NT$ {float(v):,.0f}",
                    "獲配股數": lambda v: f"{float(v):,.0f} 股"
                }).map(style_action_cell, subset=["項目"])
                st.dataframe(styled_stk_inc, use_container_width=True, hide_index=True)

# ----------------- TAB 2: 交易記帳 -----------------
with tab2:
    st.subheader("登錄新交易")

    candidates = sorted(list(set([f"{clean_stock_code(c)} {n}" for c, n in dynamic_stock_pool.items()])))

    c_search, c_member, c_broker = st.columns([3, 2, 2])
    with c_search:
        sel_candidate = st.selectbox("輸入關鍵字快速選取：", options=["自定義手動輸入"] + candidates, index=0)
    with c_member:
        trade_member = st.selectbox("歸屬成員", member_list, index=0)
    with c_broker:
        trade_broker = st.selectbox("券商帳戶", broker_list, index=0 if broker_list else None)

    if sel_candidate != "自定義手動輸入":
        preset_code = clean_stock_code(sel_candidate.split()[0])
        preset_name = sel_candidate.split()[1]
    else:
        preset_code = ""
        preset_name = ""

    c_in1, c_in2, c_in3 = st.columns([2, 3, 2])
    with c_in1:
        trade_code = st.text_input("確認代號", value=preset_code, placeholder="例: 00919 或 2417")
    with c_in2:
        trade_name = st.text_input("確認股名 (中文)", value=preset_name, placeholder="例: 群益台灣精選高息")

    target_code = clean_stock_code(trade_code)
    last_used_freq = "季配"
    freq_options = ["月配", "季配", "半年配", "年配", "不配息"]

    if target_code in stock_settings:
        last_used_freq = stock_settings[target_code]

    default_freq_idx = freq_options.index(last_used_freq) if last_used_freq in freq_options else 1

    with c_in3:
        trade_freq = st.selectbox("配息頻率", freq_options, index=default_freq_idx)

    st.write("---")
    col_t1, col_t2, col_t3, col_t4 = st.columns(4)
    t_date = col_t1.date_input("交易日期", datetime.now())
    t_action = col_t2.selectbox("交易類別", ["買進", "賣出"])
    t_shares = col_t3.number_input("股數", min_value=1, step=100, value=1000)
    t_price = col_t4.number_input("成交單價 (元)", min_value=0.01, step=0.5, value=38.0)

    save_btn = st.button("確認新增此筆交易", type="primary")

    if save_btn:
        f_code = clean_stock_code(trade_code)
        f_name = trade_name.strip()
        if not f_code:
            st.error("請填寫股票代號！")
        else:
            if not f_name or f_name == f_code:
                f_name = query_stock_name(f_code)

            next_id = int(pd.to_numeric(raw_trans_df["id"], errors="coerce").max() + 1) if not raw_trans_df.empty else 1
            new_r = pd.DataFrame([{
                "id": next_id,
                "日期": t_date.strftime("%Y/%m/%d"),  # 永遠固定 YYYY/MM/DD
                "成員": trade_member,
                "券商": trade_broker,
                "代號": f_code,
                "名稱": f_name,
                "類別": t_action,
                "股數": int(t_shares),
                "價格": float(t_price),
                "頻率": trade_freq
            }])
            
            updated_trans = pd.concat([raw_trans_df, new_r], ignore_index=True)
            save_all_transactions(updated_trans)
            
            stock_settings[f_code] = trade_freq
            save_stock_settings(stock_settings)

            st.success(f"✅ 成功寫入 Google Sheets！已登錄：【{trade_member}】[{trade_broker}] {t_action} [{f_code}] {f_name} {t_shares:,} 股")
            time.sleep(1)
            st.rerun()

    st.divider()
    st.subheader("歷史交易流水帳")

    disp_trans = trans_df.sort_values(by=["日期", "id"], ascending=[False, False]).copy()
    disp_trans.insert(0, "序號", range(1, len(disp_trans) + 1))
    disp_trans["日期"] = disp_trans["日期"].apply(normalize_date)
    disp_trans["股數"] = pd.to_numeric(disp_trans["股數"], errors="coerce").fillna(0)
    disp_trans["價格"] = pd.to_numeric(disp_trans["價格"], errors="coerce").fillna(0)

    styled_trans = disp_trans.style.format({
        "序號": "{:d}",
        "股數": lambda v: f"{float(v):,.0f}",
        "價格": lambda v: f"NT$ {float(v):,.2f}"
    }).map(style_action_cell, subset=["類別"])
    st.dataframe(styled_trans, use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("編輯與管理交易清單")
    edit_df = raw_trans_df.copy()
    edit_df.insert(0, "刪除", False)
    edit_df["日期"] = edit_df["日期"].apply(normalize_date)
    edit_df["股數"] = pd.to_numeric(edit_df["股數"], errors="coerce").fillna(0).astype(int)
    edit_df["價格"] = pd.to_numeric(edit_df["價格"], errors="coerce").fillna(0).astype(float)

    edited_table = st.data_editor(
        edit_df,
        column_config={
            "刪除": st.column_config.CheckboxColumn("勾選刪除?", default=False),
            "id": st.column_config.NumberColumn("流水號", disabled=True),
            "日期": st.column_config.TextColumn("交易日期 (YYYY/MM/DD)"),
            "成員": st.column_config.SelectboxColumn("成員", options=member_list),
            "券商": st.column_config.SelectboxColumn("券商", options=broker_list),
            "代號": st.column_config.TextColumn("代號"),
            "名稱": st.column_config.TextColumn("名稱"),
            "類別": st.column_config.SelectboxColumn("類別", options=["買進", "賣出", "配股"]),
            "股數": st.column_config.NumberColumn("股數", min_value=1, step=100, format="%d"),
            "價格": st.column_config.NumberColumn("單價", min_value=0.0, step=0.1, format="%.2f"),
            "頻率": st.column_config.SelectboxColumn("配息頻率", options=["月配", "季配", "半年配", "年配", "不配息"]),
        },
        disabled=["id"],
        hide_index=True,
        use_container_width=True,
        key="tx_editor_v16"
    )

    delete_targets = edited_table[edited_table["刪除"] == True]
    num_to_delete = len(delete_targets)

    if st.button("儲存修改內容", type="primary"):
        clean_save_df = edited_table[edited_table["刪除"] == False].drop(columns=["刪除"])
        for idx, r in clean_save_df.iterrows():
            c = clean_stock_code(r["代號"])
            clean_save_df.at[idx, "代號"] = c
            clean_save_df.at[idx, "日期"] = normalize_date(r["日期"])
            real_name = query_stock_name(c)
            if real_name != c or not r["名稱"]:
                clean_save_df.at[idx, "名稱"] = real_name

        save_all_transactions(clean_save_df)
        st.success("所有修改已成功同步回 Google 試算表！")
        time.sleep(1)
        st.rerun()

    if num_to_delete > 0:
        st.warning(f"⚠️ 警告：目前已勾選 {num_to_delete} 筆交易準備刪除！")
        if st.checkbox("我已確認要永久刪除上述勾選的交易紀錄", value=False):
            if st.button("確認永久刪除勾選項目", type="secondary"):
                remain_after_del = edited_table[edited_table["刪除"] == False].drop(columns=["刪除"])
                save_all_transactions(remain_after_del)
                st.success(f"已成功刪除 {num_to_delete} 筆交易！")
                time.sleep(1)
                st.rerun()

# ----------------- TAB 3: 已實現賣出損益 -----------------
with tab3:
    st.subheader("已實現賣出損益明細")
    if realized_df.empty:
        st.info("目前尚未有賣出紀錄。")
    else:
        tot_realized = realized_df["實現損益"].sum()
        pnl_pct = (tot_realized / realized_df["總成本"].sum() * 100) if realized_df["總成本"].sum() > 0 else 0.0

        st.metric("累計已實現損益總計", f"NT$ {tot_realized:,.0f}", delta=f"{pnl_pct:+.2f}%")

        disp_real = realized_df.sort_values(by="日期", ascending=False)[[
            "日期", "成員", "券商", "代號", "名稱", "類別", "賣出股數", "賣出價格", "成本單價", "總成本", "實現損益"
        ]].copy()
        disp_real.insert(0, "序號", range(1, len(disp_real) + 1))
        disp_real["日期"] = disp_real["日期"].apply(normalize_date)
        disp_real["賣出股數"] = pd.to_numeric(disp_real["賣出股數"], errors="coerce").fillna(0)
        disp_real["賣出價格"] = pd.to_numeric(disp_real["賣出價格"], errors="coerce").fillna(0)
        disp_real["成本單價"] = pd.to_numeric(disp_real["成本單價"], errors="coerce").fillna(0)
        disp_real["總成本"] = pd.to_numeric(disp_real["總成本"], errors="coerce").fillna(0)
        disp_real["實現損益"] = pd.to_numeric(disp_real["實現損益"], errors="coerce").fillna(0)

        styled_real = disp_real.style.format({
            "序號": "{:d}",
            "賣出股數": lambda v: f"{float(v):,.0f}",
            "賣出價格": lambda v: f"NT$ {float(v):,.2f}",
            "成本單價": lambda v: f"NT$ {float(v):,.2f}",
            "總成本": lambda v: f"NT$ {float(v):,.0f}",
            "實現損益": lambda v: f"NT$ {float(v):+,.0f}"
        }).map(style_pnl_value, subset=["實現損益"]).map(style_action_cell, subset=["類別"])

        st.dataframe(styled_real, use_container_width=True, hide_index=True)

# ----------------- TAB 4: 實際配息與配股登記 -----------------
with tab4:
    st.subheader("實際收息、配股與借券收入登記")

    def on_inc_stock_change():
        chosen = st.session_state.inc_stock_sel_key
        if chosen != "自定義輸入":
            st.session_state.inc_code_val_key = clean_stock_code(chosen.split()[0])
            st.session_state.inc_name_val_key = chosen.split()[1]
        else:
            st.session_state.inc_code_val_key = ""
            st.session_state.inc_name_val_key = ""

    if "inc_code_val_key" not in st.session_state:
        st.session_state.inc_code_val_key = ""
    if "inc_name_val_key" not in st.session_state:
        st.session_state.inc_name_val_key = ""

    with st.expander("登記新入帳", expanded=True):
        col_inc1, col_inc2, col_inc_mem, col_inc_brk = st.columns(4)
        inc_date = col_inc1.date_input("入帳日期", datetime.now(), key="inc_date_input")
        inc_type = col_inc2.selectbox("收入類別", ["現金股利", "借券收入", "股票股利 (配股)"], key="inc_type_input")
        inc_member = col_inc_mem.selectbox("歸屬成員", member_list, index=0, key="inc_member_input")
        inc_broker = col_inc_brk.selectbox("入帳券商", broker_list, index=0 if broker_list else None, key="inc_broker_input")

        known_stocks = sorted(list(set([f"{clean_stock_code(c)} {n}" for c, n in dynamic_stock_pool.items()])))
        st.selectbox(
            "入帳標的",
            ["自定義輸入"] + known_stocks,
            key="inc_stock_sel_key",
            on_change=on_inc_stock_change
        )

        col_inc4, col_inc5, col_inc6 = st.columns(3)
        inc_code_input = col_inc4.text_input("確認代號", key="inc_code_val_key")
        inc_name_input = col_inc5.text_input("確認股名", key="inc_name_val_key")

        if inc_type == "股票股利 (配股)":
            inc_shares_get = col_inc6.number_input("獲配股票股數", min_value=1, step=10, value=100)
            inc_amount_get = 0
            auto_sync_holdings = True
            st.caption("✅ 系統已自動設定：同步寫入交易清單")
        else:
            inc_amount_get = col_inc6.number_input("實收現金金額 (元)", min_value=1, step=500, value=5000)
            inc_shares_get = 0
            auto_sync_holdings = False

        inc_note = st.text_input("備註說明", value="", key="inc_note_input")

        if st.button("確認登記入帳", type="primary"):
            final_inc_code = clean_stock_code(inc_code_input)
            final_inc_name = inc_name_input.strip() or query_stock_name(final_inc_code)
            year_val = int(inc_date.strftime("%Y"))

            if not final_inc_code:
                st.error("請輸入或選擇入帳標的代號！")
            else:
                next_inc_id = int(pd.to_numeric(raw_income_df["id"], errors="coerce").max() + 1) if not raw_income_df.empty else 1
                new_inc_row = pd.DataFrame([{
                    "id": next_inc_id,
                    "日期": inc_date.strftime("%Y/%m/%d"),  # 永遠固定 YYYY/MM/DD
                    "年度": year_val,
                    "成員": inc_member,
                    "券商": inc_broker,
                    "代號": final_inc_code,
                    "名稱": final_inc_name,
                    "項目": "股票股利" if "股票股利" in inc_type else inc_type,
                    "金額": int(inc_amount_get),
                    "獲配股數": int(inc_shares_get),
                    "備註": inc_note
                }])
                
                updated_inc = pd.concat([raw_income_df, new_inc_row], ignore_index=True)
                save_income_records(updated_inc)

                if auto_sync_holdings and inc_shares_get > 0:
                    next_tx_id = int(pd.to_numeric(raw_trans_df["id"], errors="coerce").max() + 1) if not raw_trans_df.empty else 1
                    tx_sync_row = pd.DataFrame([{
                        "id": next_tx_id,
                        "日期": inc_date.strftime("%Y/%m/%d"),
                        "成員": inc_member,
                        "券商": inc_broker,
                        "代號": final_inc_code,
                        "名稱": final_inc_name,
                        "類別": "配股",
                        "股數": int(inc_shares_get),
                        "價格": 0.0,
                        "頻率": stock_settings.get(final_inc_code, "季配")
                    }])
                    updated_tx = pd.concat([raw_trans_df, tx_sync_row], ignore_index=True)
                    save_all_transactions(updated_tx)

                st.success(f"✅ 成功登記入帳：【{inc_member}】[{inc_broker}] {year_val}年 {final_inc_name} {inc_type}")
                time.sleep(1)
                st.rerun()

    st.divider()
    st.subheader("🎁 歷年股票股利獨立清單")
    stock_divs_only = income_df[income_df["項目"] == "股票股利"].copy()

    if stock_divs_only.empty:
        st.info("目前尚無任何股票股利入帳紀錄。")
    else:
        tot_shares_awarded = pd.to_numeric(stock_divs_only["獲配股數"], errors="coerce").sum()
        st.caption(f"全年度累積配股總數：{tot_shares_awarded:,.0f} 股")

        cols_s = ["日期", "年度", "代號", "名稱", "項目", "獲配股數"]
        if "成員" in stock_divs_only.columns:
            cols_s.insert(2, "成員")
        if "券商" in stock_divs_only.columns:
            cols_s.insert(3, "券商")
        if "備註" in stock_divs_only.columns:
            cols_s.append("備註")

        disp_stock_divs = stock_divs_only.sort_values(by="日期", ascending=False)[cols_s].copy()
        disp_stock_divs.insert(0, "序號", range(1, len(disp_stock_divs) + 1))
        disp_stock_divs["日期"] = disp_stock_divs["日期"].apply(normalize_date)
        disp_stock_divs["獲配股數"] = pd.to_numeric(disp_stock_divs["獲配股數"], errors="coerce").fillna(0)

        styled_stock_divs = disp_stock_divs.style.format({
            "序號": "{:d}",
            "獲配股數": lambda v: f"{float(v):,.0f} 股"
        }).map(style_action_cell, subset=["項目"])

        st.dataframe(styled_stock_divs, use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("分年度收益彙整看板")

    if not income_df.empty:
        inc_copy = income_df.copy()
        inc_copy["金額"] = pd.to_numeric(inc_copy["金額"], errors="coerce").fillna(0)
        inc_copy["獲配股數"] = pd.to_numeric(inc_copy["獲配股數"], errors="coerce").fillna(0)
        inc_copy["年度"] = pd.to_numeric(inc_copy["年度"], errors="coerce").fillna(0).astype(int)

        cash_df = inc_copy[inc_copy["項目"].isin(["現金股利", "借券收入"])]
        yearly_cash = cash_df.pivot_table(
            index="年度",
            columns="項目",
            values="金額",
            aggfunc="sum",
            fill_value=0
        ).reset_index()

        if "現金股利" not in yearly_cash.columns:
            yearly_cash["現金股利"] = 0
        if "借券收入" not in yearly_cash.columns:
            yearly_cash["借券收入"] = 0

        yearly_cash["年度現金總計"] = yearly_cash["現金股利"] + yearly_cash["借券收入"]

        stock_div_df = inc_copy[inc_copy["項目"] == "股票股利"]
        yearly_stock = stock_div_df.groupby("年度")["獲配股數"].sum().reset_index()

        yearly_summary = pd.merge(yearly_cash, yearly_stock, on="年度", how="outer").fillna(0)
        yearly_summary = yearly_summary.sort_values(by="年度", ascending=False).reset_index(drop=True)

        all_time_cash_div = yearly_summary["現金股利"].sum()
        all_time_lend_inc = yearly_summary["借券收入"].sum()
        all_time_total_cash = yearly_summary["年度現金總計"].sum()
        all_time_stock_shares = yearly_summary["獲配股數"].sum()

        my1, my2, my3, my4 = st.columns(4)
        my1.metric("歷年累計現金總收益", f"NT$ {all_time_total_cash:,.0f}")
        my2.metric("歷年累計現金股利", f"NT$ {all_time_cash_div:,.0f}")
        my3.metric("歷年累計借券收入", f"NT$ {all_time_lend_inc:,.0f}")
        my4.metric("歷年累計獲配股票", f"{all_time_stock_shares:,.0f} 股")
        st.write("")

        yearly_summary.insert(0, "序號", range(1, len(yearly_summary) + 1))
        yearly_total_row = {
            "序號": "合計",
            "年度": "-",
            "現金股利": all_time_cash_div,
            "借券收入": all_time_lend_inc,
            "年度現金總計": all_time_total_cash,
            "獲配股數": all_time_stock_shares
        }
        yearly_summary_disp = pd.concat([yearly_summary, pd.DataFrame([yearly_total_row])], ignore_index=True)

        styled_summary = yearly_summary_disp.style.format({
            "序號": "{}",
            "年度": lambda v: f"{int(v)}" if isinstance(v, (int, float)) and v > 0 else str(v),
            "現金股利": lambda v: f"NT$ {float(v):,.0f}" if isinstance(v, (int, float)) else str(v),
            "借券收入": lambda v: f"NT$ {float(v):,.0f}" if isinstance(v, (int, float)) else str(v),
            "年度現金總計": lambda v: f"NT$ {float(v):,.0f}" if isinstance(v, (int, float)) else str(v),
            "獲配股數": lambda v: f"{float(v):,.0f} 股" if isinstance(v, (int, float)) else str(v)
        })
        st.dataframe(styled_summary, use_container_width=True, hide_index=True)

# ----------------- TAB 5: 報表匯出 -----------------
with tab5:
    st.subheader("📋 自訂條件統計報表")

    col_rep1, col_rep_sd, col_rep_ed = st.columns(3)
    with col_rep1:
        report_type = st.selectbox(
            "選擇要產出的報表類型：",
            [
                "買賣交易流水帳報表",
                "實際收益 (股利/借券/配股) 報表",
                "已實現賣出損益報表"
            ]
        )

    with col_rep_sd:
        start_date_val = st.date_input("開始日期 (起)", value=date(2025, 1, 1), key="rep_start_date")
    with col_rep_ed:
        end_date_val = st.date_input("結束日期 (迄)", value=datetime.now().date(), key="rep_end_date")

    col_rep_brk, col_rep_stk = st.columns(2)
    with col_rep_brk:
        filter_brokers = st.multiselect("篩選券商帳戶 (留空代表全部)：", options=broker_list, default=[])
    with col_rep_stk:
        all_clean_codes = sorted(list(dynamic_stock_pool.keys()))
        filter_stocks = st.multiselect("篩選特定股票代號 (留空代表全部)：", options=all_clean_codes, default=[])

    start_d = normalize_date(str(start_date_val))
    end_d = normalize_date(str(end_date_val))
    report_member_target = current_selected_member

    conditions_summary = {
        "檢視帳戶": report_member_target,
        "報表類型": report_type,
        "日期區間": f"{start_d} ~ {end_d}",
        "券商帳戶": ", ".join(filter_brokers) if filter_brokers else "全部券商",
        "股票標的": ", ".join(filter_stocks) if filter_stocks else "全部標的"
    }

    filtered_report_df = pd.DataFrame()

    if report_type == "買賣交易流水帳報表":
        df_src = raw_trans_df.copy()
        df_src["日期"] = df_src["日期"].apply(normalize_date)
        df_src = df_src[(df_src["日期"] >= start_d) & (df_src["日期"] <= end_d)]
        if report_member_target != "全部家庭成員 (合併總覽)":
            df_src = df_src[df_src["成員"] == report_member_target]
        if filter_stocks:
            df_src = df_src[df_src["代號"].isin(filter_stocks)]
        if filter_brokers:
            df_src = df_src[df_src["券商"].isin(filter_brokers)]
        filtered_report_df = df_src.sort_values(by="日期", ascending=False)

    elif report_type == "實際收益 (股利/借券/配股) 報表":
        df_src = raw_income_df.copy()
        df_src["日期"] = df_src["日期"].apply(normalize_date)
        df_src = df_src[(df_src["日期"] >= start_d) & (df_src["日期"] <= end_d)]
        if report_member_target != "全部家庭成員 (合併總覽)":
            df_src = df_src[df_src["成員"] == report_member_target]
        if filter_stocks:
            df_src = df_src[df_src["代號"].isin(filter_stocks)]
        if filter_brokers:
            df_src = df_src[df_src["券商"].isin(filter_brokers)]
        filtered_report_df = df_src.sort_values(by="日期", ascending=False)

    elif report_type == "已實現賣出損益報表":
        _, all_realized_df = calculate_portfolio(raw_trans_df, stock_settings)
        if not all_realized_df.empty:
            df_src = all_realized_df.copy()
            df_src["日期"] = df_src["日期"].apply(normalize_date)
            df_src = df_src[(df_src["日期"] >= start_d) & (df_src["日期"] <= end_d)]
            if report_member_target != "全部家庭成員 (合併總覽)":
                df_src = df_src[df_src["成員"] == report_member_target]
            if filter_stocks:
                df_src = df_src[df_src["代號"].isin(filter_stocks)]
            if filter_brokers:
                df_src = df_src[df_src["券商"].isin(filter_brokers)]
            filtered_report_df = df_src.sort_values(by="日期", ascending=False)

    if not filtered_report_df.empty:
        st.write("---")
        cond_text_disp = " ｜ ".join([f"**{k}**: {v}" for k, v in conditions_summary.items()])
        st.info(f"📌 {cond_text_disp} ｜ **資料筆數: 共 {len(filtered_report_df)} 筆**")

        df_with_no = filtered_report_df.copy()
        df_with_no.insert(0, "序號", range(1, len(df_with_no) + 1))
        df_with_no["日期"] = df_with_no["日期"].apply(normalize_date)

        if report_type == "買賣交易流水帳報表":
            df_with_no["股數"] = pd.to_numeric(df_with_no["股數"], errors="coerce").fillna(0)
            df_with_no["價格"] = pd.to_numeric(df_with_no["價格"], errors="coerce").fillna(0)

            b_mask = df_with_no["類別"] == "買進"
            s_mask = df_with_no["類別"] == "賣出"
            tot_b_amt = (df_with_no.loc[b_mask, "股數"] * df_with_no.loc[b_mask, "價格"]).sum()
            tot_s_amt = (df_with_no.loc[s_mask, "股數"] * df_with_no.loc[s_mask, "價格"]).sum()
            tot_shares = df_with_no["股數"].sum()

            m1, m2, m3 = st.columns(3)
            m1.metric("交易總股數", f"{tot_shares:,.0f} 股")
            m2.metric("買進總金額", f"NT$ {tot_b_amt:,.0f}")
            m3.metric("賣出總金額", f"NT$ {tot_s_amt:,.0f}")

            sum_row = {c: "" for c in df_with_no.columns}
            sum_row["序號"] = "合計"
            sum_row["股數"] = tot_shares
            df_with_no_disp = pd.concat([df_with_no, pd.DataFrame([sum_row])], ignore_index=True)

            styled_report = df_with_no_disp.style.format({
                "序號": "{}",
                "股數": lambda v: f"{float(v):,.0f}" if isinstance(v, (int, float)) else str(v),
                "價格": lambda v: f"NT$ {float(v):,.2f}" if isinstance(v, (int, float)) and v > 0 else str(v)
            }).map(style_action_cell, subset=["類別"])
            st.dataframe(styled_report, use_container_width=True, hide_index=True)

        col_down1, _ = st.columns([1, 1])
        file_stem = f"股票報表_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        with col_down1:
            csv_data = df_with_no.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
            st.download_button(
                label="📥 匯出此份報表為 CSV",
                data=csv_data,
                file_name=f"{file_stem}.csv",
                mime="text/csv",
                key="download_csv_btn"
            )
    else:
        st.warning("⚠️ 所選條件區間內無任何符合的資料！")

# ----------------- TAB 6: 基本資料設定與安全檢查 -----------------
with tab6:
    st.subheader("👥 基本資料維護與防呆安全檢查")
    st.caption("為維護資產紀錄完整性，若成員、券商尚有進出紀錄或庫存，系統將自動封鎖刪除功能並提示警示。")

    sub_m1, sub_m2 = st.columns(2)

    with sub_m1:
        st.write("#### 👤 家庭成員管理")
        col_add_m1, col_add_m2 = st.columns([3, 1])
        new_m_val = col_add_m1.text_input("新增成員名稱", placeholder="例如：女兒、兒子", key="add_m_input")
        if col_add_m2.button("➕ 新增", key="btn_add_m"):
            m_clean = new_m_val.strip()
            if m_clean and m_clean not in member_list:
                member_list.append(m_clean)
                save_members(member_list)
                st.success(f"已新增成員：{m_clean}")
                time.sleep(1)
                st.rerun()

        st.write("---")
        del_target_m = st.selectbox("選擇欲刪除之成員", options=["-- 請選擇 --"] + member_list, key="sel_del_m")
        if st.button("🗑️ 刪除所選成員", key="btn_del_m"):
            if del_target_m != "-- 請選擇 --":
                if len(member_list) <= 1:
                    st.error("⚠️ 至少需保留一位家庭成員！")
                else:
                    tx_count = len(raw_trans_df[raw_trans_df["成員"] == del_target_m]) if "成員" in raw_trans_df.columns else 0
                    inc_count = len(raw_income_df[raw_income_df["成員"] == del_target_m]) if "成員" in raw_income_df.columns else 0
                    if tx_count + inc_count > 0:
                        st.error(f"⛔ 無法刪除【{del_target_m}】！名下尚有 {tx_count} 筆交易與 {inc_count} 筆收益紀錄。")
                    else:
                        member_list = [m for m in member_list if m != del_target_m]
                        save_members(member_list)
                        st.success(f"已成功刪除成員：{del_target_m}")
                        time.sleep(1)
                        st.rerun()

    with sub_m2:
        st.write("#### 🏢 券商帳戶管理")
        col_add_b1, col_add_b2 = st.columns([3, 1])
        new_b_val = col_add_b1.text_input("新增券商名稱", placeholder="例如：新光證券、口袋證券", key="add_b_input")
        if col_add_b2.button("➕ 新增", key="btn_add_b"):
            b_clean = new_b_val.strip()
            if b_clean and b_clean not in broker_list:
                broker_list.append(b_clean)
                save_brokers(broker_list)
                st.success(f"已新增券商：{b_clean}")
                time.sleep(1)
                st.rerun()

        st.write("---")
        del_target_b = st.selectbox("選擇欲刪除之券商", options=["-- 請選擇 --"] + broker_list, key="sel_del_b")
        if st.button("🗑️ 刪除所選券商", key="btn_del_b"):
            if del_target_b != "-- 請選擇 --":
                if len(broker_list) <= 1:
                    st.error("⚠️ 至少需保留一家券商帳戶！")
                else:
                    b_tx_cnt = len(raw_trans_df[raw_trans_df["券商"] == del_target_b]) if "券商" in raw_trans_df.columns else 0
                    b_inc_cnt = len(raw_income_df[raw_income_df["券商"] == del_target_b]) if "券商" in raw_income_df.columns else 0
                    if b_tx_cnt + b_inc_cnt > 0:
                        st.error(f"⛔ 無法刪除【{del_target_b}】！該券商尚有 {b_tx_cnt} 筆交易與 {b_inc_cnt} 筆收益關聯資料。")
                    else:
                        broker_list = [b for b in broker_list if b != del_target_b]
                        save_brokers(broker_list)
                        st.success(f"已成功刪除券商：{del_target_b}")
                        time.sleep(1)
                        st.rerun()
