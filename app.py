import streamlit as st
import requests
import json
from urllib.parse import urlencode
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO

# -------------------- 登入驗證區塊 --------------------
st.set_page_config(page_title="登入驗證", page_icon="🔐")

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

# 處理登入驗證流程
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

# -------------------- 系統主功能區 --------------------

FILE_URL = "https://raw.githubusercontent.com/ainstaccc/mst-kpi-v2/main/v2-2025.06_MST-PA.xlsx"

@st.cache_data(ttl=3600)
def load_data():
    def parse_sheet(sheet_name, default_header=1):
        for hdr in [default_header, 2]:
            try:
                df = pd.read_excel(FILE_URL, sheet_name=sheet_name, header=hdr, engine="openpyxl")
                if "區主管" in df.columns:
                    return df  # 成功找到區主管欄，視為正確 header
            except Exception as e:
                print(f"讀取 {sheet_name} 發生錯誤（header={hdr}）：{e}")
        # 如果都失敗，就回傳空表
        return pd.DataFrame()


    df_summary = parse_sheet("門店 考核總表", default_header=1)
    df_eff     = parse_sheet("人效分析", default_header=1)
    df_mgr     = parse_sheet("店長副店 考核明細", default_header=2)  # ✅ 預設直接抓 header=2
    df_staff   = parse_sheet("店員儲備 考核明細", default_header=2)  # ✅ 預設直接抓 header=2
    df_dist    = pd.read_excel(FILE_URL, sheet_name="等級分布", header=None, nrows=15, usecols="A:N", engine="openpyxl")
    summary_month = pd.read_excel(FILE_URL, sheet_name="門店 考核總表", nrows=1, engine="openpyxl").columns[0]

    return df_summary, df_eff, df_mgr, df_staff, df_dist, summary_month

def format_eff(df):
    if df.empty:
        return df
    df = df.copy()
    for col in ["個績目標", "個績貢獻", "品牌 客單價", "個人 客單價"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').round(1)
    for col in ["個績達成%", "客單 相對績效", "品牌 結帳會員率", "個人 結帳會員率", "會員 相對績效"]:
        if col in df.columns:
            df[col] = df[col].apply(lambda x: f"{x}%" if pd.notnull(x) else x)
    return df

def main():
    st.markdown("<h3>📊 米斯特 門市 工作績效月考核查詢系統</h3>", unsafe_allow_html=True)
    df_summary, df_eff, df_mgr, df_staff, df_dist, summary_month = load_data()

    with st.expander("🔍 查詢條件", expanded=True):
        st.markdown("**🔺查詢條件任一欄即可，避免多重條件造成查詢錯誤。**")
        col1, col2 = st.columns(2)
        area = col1.selectbox("區域/區主管", options=[""] + list(df_summary["區主管"].dropna().unique()))
        dept_code = col2.text_input("部門編號/門店編號")
        month = st.selectbox("查詢月份", options=["2025/06"])

    if st.button("🔎 查詢", type="primary"):
        mask = pd.Series(True, index=df_summary.index)
        if area:
            mask &= df_summary["區主管"] == area
        if dept_code:
            mask &= df_summary["部門編號"] == dept_code
        df_result = df_summary[mask]

        eff_mask = df_eff["區主管"].eq(area) if area else pd.Series(True, index=df_eff.index)
        if dept_code:
            eff_mask &= df_eff["部門編號"] == dept_code
        df_eff_result = df_eff[eff_mask]

        mgr_mask = df_mgr["區主管"].eq(area) if area else pd.Series(True, index=df_mgr.index)
        if dept_code:
            mgr_mask &= df_mgr["部門編號"] == dept_code
        df_mgr_result = df_mgr[mgr_mask]

        staff_mask = df_staff["區主管"].eq(area) if area else pd.Series(True, index=df_staff.index)
        if dept_code:
            staff_mask &= df_staff["部門編號"] == dept_code
        df_staff_result = df_staff[staff_mask]

        st.markdown("## 🧾 門店考核總表")
        st.dataframe(df_result.iloc[:, 2:11], use_container_width=True)

        st.markdown("## 👥 人效分析")
        df_eff_result_fmt = format_eff(df_eff_result)
        st.dataframe(df_eff_result_fmt, use_container_width=True)

        st.markdown("## 👔 店長/副店 考核明細")
        df_mgr_display = pd.concat([df_mgr_result.iloc[:, 1:7], df_mgr_result.iloc[:, 11:28]], axis=1)
        st.dataframe(df_mgr_display, use_container_width=True)

        st.markdown("## 👟 店員/儲備 考核明細")
        df_staff_display = pd.concat([df_staff_result.iloc[:, 1:7], df_staff_result.iloc[:, 11:28]], axis=1)
        st.dataframe(df_staff_display, use_container_width=True)

        output_excel = BytesIO()
        with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
            df_result.iloc[:, 2:11].to_excel(writer, sheet_name="門店考核總表", index=False)
            df_eff_result_fmt.to_excel(writer, sheet_name="人效分析", index=False)
            df_mgr_display.to_excel(writer, sheet_name="店長副店 考核明細", index=False)
            df_staff_display.to_excel(writer, sheet_name="店員儲備 考核明細", index=False)
        output_excel.seek(0)

        st.download_button("📥 下載查詢結果 Excel", data=output_excel, file_name="考核查詢結果.xlsx")

        st.markdown("<p style='color:red;font-weight:bold;font-size:16px;'>※如對分數有疑問，請洽區主管/品牌經理說明。</p>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
