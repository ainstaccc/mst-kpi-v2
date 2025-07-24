import streamlit as st
import requests
import json
from urllib.parse import urlencode
import pandas as pd
from io import BytesIO

# -------------------- 頁面設定 --------------------
st.set_page_config(page_title="米斯特績效考核查詢", page_icon="📊")

# -------------------- 登入驗證區塊 --------------------
GOOGLE_CLIENT_ID = st.secrets["google_oauth"]["client_id"]
GOOGLE_CLIENT_SECRET = st.secrets["google_oauth"]["client_secret"]
REDIRECT_URI = st.secrets["google_oauth"]["redirect_uri"]
ALLOWED_USERS = [email.lower() for email in st.secrets["google_oauth"]["allowed_users"]]

def get_login_url():
    params = {
        "client_id": GOOGLE_CLIENT_ID,
        "redirect_uri": REDIRECT_URI,
        "response_type": "code",
        "scope": "https://www.googleapis.com/auth/userinfo.email",
        "access_type": "offline",
        "prompt": "select_account"
    }
    return f"https://accounts.google.com/o/oauth2/v2/auth?{urlencode(params)}"

def get_token(code):
    data = {
        "code": code,
        "client_id": GOOGLE_CLIENT_ID,
        "client_secret": GOOGLE_CLIENT_SECRET,
        "redirect_uri": REDIRECT_URI,
        "grant_type": "authorization_code"
    }
    response = requests.post("https://oauth2.googleapis.com/token", data=data)
    return response.json()

def get_user_info(access_token):
    headers = {"Authorization": f"Bearer {access_token}"}
    response = requests.get("https://www.googleapis.com/oauth2/v2/userinfo", headers=headers)
    return response.json()

# -------------------- 登入驗證流程 --------------------
query_params = st.experimental_get_query_params()
code = query_params.get("code", [None])[0]

if "user_email" not in st.session_state:
    if code:
        token_response = get_token(code)
        access_token = token_response.get("access_token")
        if access_token:
            user_info = get_user_info(access_token)
            email = user_info.get("email", "").lower()
            if email in ALLOWED_USERS:
                st.session_state.user_email = email
                st.success(f"歡迎，{email}！")
            else:
                st.error("❌ 此帳號未授權存取此應用程式。")
                st.stop()
        else:
            st.error("⚠️ 無法取得 access token。請重新登入。")
            st.stop()
    else:
        login_url = get_login_url()
        st.markdown(f"[Hello，米斯特的夥伴! 請登入 Google帳號，驗證後開始查詢考核成績 📊 ]({login_url})")
        st.stop()
else:
    st.write(f"✅ 你已登入：**{st.session_state.user_email}**")

# -------------------- 資料載入區 --------------------
FILE_URL = "https://raw.githubusercontent.com/ainstaccc/mst-kpi-v2/main/v2-2025.06_MST-PA.xlsx"

@st.cache_data(ttl=3600)
def load_data():
    try:
        xls = pd.ExcelFile(FILE_URL, engine="openpyxl")

        def check_columns(df, name):
            if "區主管" not in df.columns or "部門編號" not in df.columns:
                st.error(f"❌ 【{name}】中缺少必要欄位：區主管 或 部門編號")
                st.stop()

        df_summary = xls.parse("門店 考核總表", header=1)
        check_columns(df_summary, "門店 考核總表")

        df_eff = xls.parse("人效分析", header=1)
        check_columns(df_eff, "人效分析")

        df_mgr = xls.parse("店長副店 考核明細", header=1)
        check_columns(df_mgr, "店長副店 考核明細")

        df_staff = xls.parse("店員儲備 考核明細", header=1)
        check_columns(df_staff, "店員儲備 考核明細")

        df_dist = xls.parse("等級分布", header=None, nrows=15, usecols="A:N")

        try:
            summary_month = xls.parse("門店 考核總表", header=None, nrows=1).iloc[0, 0]
        except Exception:
            summary_month = "未知月份"

        return df_summary, df_eff, df_mgr, df_staff, df_dist, summary_month
    except Exception as e:
        st.error(f"❌ 資料載入失敗：{e}")
        return None, None, None, None, None, None

def format_eff(df):
    if df.empty:
        return df
    df = df.copy()
    # 模糊找出含換行的欄位
    rename_map = {}
    for col in df.columns:
        if "品牌" in col and "客單" in col:
            rename_map[col] = "品牌 客單價"
        elif "個人" in col and "客單" in col:
            rename_map[col] = "個人 客單價"
        elif "客單" in col and "相對" in col:
            rename_map[col] = "客單 相對績效"
        elif "品牌" in col and "會員率" in col:
            rename_map[col] = "品牌 結帳會員率"
        elif "個人" in col and "會員率" in col:
            rename_map[col] = "個人 結帳會員率"
        elif "會員" in col and "相對" in col:
            rename_map[col] = "會員 相對績效"
    df = df.rename(columns=rename_map)

    for col in ["個績目標", "個績貢獻", "品牌 客單價", "個人 客單價"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').round(1)
    for col in ["個績達成%", "客單 相對績效", "品牌 結帳會員率", "個人 結帳會員率", "會員 相對績效"]:
        if col in df.columns:
            df[col] = df[col].apply(lambda x: f"{x}%" if pd.notnull(x) else x)
    return df

def run_main_ui():
    st.markdown("<h3>📊 米斯特 門市 工作績效月考核查詢系統</h3>", unsafe_allow_html=True)

    df_summary, df_eff, df_mgr, df_staff, df_dist, summary_month = load_data()
    if df_summary is None:
        st.stop()

    with st.expander("🔍 查詢條件", expanded=True):
        st.markdown("**🔺查詢條件任一欄即可，避免多重條件造成查詢錯誤。**")
        col1, col2 = st.columns(2)
        area = col1.selectbox("區域/區主管", options=[""] + sorted(df_summary["區主管"].dropna().unique()))
        dept_code = col2.text_input("部門編號/門店編號")
        st.markdown(f"查詢月份：**{summary_month}**")

    if st.button("🔎 查詢", type="primary"):
        try:
            df_result = df_summary.copy()
            if area:
                df_result = df_result[df_result["區主管"] == area]
            if dept_code:
                df_result = df_result[df_result["部門編號"] == dept_code]

            df_eff_result = df_eff.copy()
            if area:
                df_eff_result = df_eff_result[df_eff_result["區主管"] == area]
            if dept_code:
                df_eff_result = df_eff_result[df_eff_result["部門編號"] == dept_code]

            df_mgr_result = df_mgr.copy()
            if area:
                df_mgr_result = df_mgr_result[df_mgr_result["區主管"] == area]
            if dept_code:
                df_mgr_result = df_mgr_result[df_mgr_result["部門編號"] == dept_code]

            df_staff_result = df_staff.copy()
            if area:
                df_staff_result = df_staff_result[df_staff_result["區主管"] == area]
            if dept_code:
                df_staff_result = df_staff_result[df_staff_result["部門編號"] == dept_code]

            st.image("https://raw.githubusercontent.com/ainstaccc/kpi-checker/main/2025.06_grade.jpg", use_column_width=True)

            st.markdown("## 🧾 門店考核總表")
            st.markdown(f"共查得：{len(df_result)} 筆")
            st.dataframe(df_result.iloc[:, 2:11] if not df_result.empty else pd.DataFrame(["查無資料"]), use_container_width=True)

            st.markdown("## 👥 人效分析")
            df_eff_fmt = format_eff(df_eff_result)
            st.markdown(f"共查得：{len(df_eff_fmt)} 筆")
            st.dataframe(df_eff_fmt if not df_eff_fmt.empty else pd.DataFrame(["查無資料"]), use_container_width=True)

            st.markdown("## 👔 店長/副店 考核明細")
            df_mgr_display = pd.concat([df_mgr_result.iloc[:, 1:7], df_mgr_result.iloc[:, 11:28]], axis=1) if not df_mgr_result.empty else pd.DataFrame(["查無資料"])
            st.markdown(f"共查得：{len(df_mgr_result)} 筆")
            st.dataframe(df_mgr_display, use_container_width=True)

            st.markdown("## 👟 店員/儲備 考核明細")
            df_staff_display = pd.concat([df_staff_result.iloc[:, 1:7], df_staff_result.iloc[:, 11:28]], axis=1) if not df_staff_result.empty else pd.DataFrame(["查無資料"])
            st.markdown(f"共查得：{len(df_staff_result)} 筆")
            st.dataframe(df_staff_display, use_container_width=True)

            st.markdown("<p style='color:red;font-weight:bold;font-size:16px;'>※如對分數有疑問，請洽區主管/品牌經理說明。</p>", unsafe_allow_html=True)

        except Exception as e:
            st.error(f"❌ 查詢過程發生錯誤：{e}")

# -------------------- 執行主流程 --------------------
if __name__ == "__main__":
    run_main_ui()
