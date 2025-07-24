import streamlit as st
import requests
import json
from urllib.parse import urlencode
import pandas as pd
from io import BytesIO

# 📌 Email 對應使用者中文姓名
EMAIL_TO_NAME = {
    "fp71612@gmail.com": "蕭中",
    "fabio89608@gmail.com": "李政勳",
    "124453221s@gmail.com": "鄧思思",
    "yolu902@gmail.com": "林宥儒",
    "a6108568@gmail.com": "羅婉心",
    "wmksue12976@gmail.com": "王建樹",
    "aqianyu8@gmail.com": "楊茜聿",
    "happy0623091@gmail.com": "陳宥蓉",
    "cvcv0897@gmail.com": "吳岱侑",
    "minkatieweng@gmail.com": "翁聖閔",
    "a0956505289@gmail.com": "黃啟周",
    "noncks@gmail.com": "栗晉屏",
    "vicecolife0969@gmail.com": "王瑞辰",
    "life8ray@gmail.com": "RAY",
    "inthing123@gmail.com": "IVEN",
    "leslie641230@gmail.com": "廣安",
    "ainstaccc@gmail.com": "沛瑜",
    "life8x35@gmail.com": "ALL WEARS台中麗寶店(AKW)",
    "allwears04@gmail.com": "ALL WEARS台中三井店(AKW)",
    "life8x33@gmail.com": "ALL WEARS台中新時代店(AKW)",
    "allwears05@gmail.com": "ALL WEARS台中大遠百店(AKW)",
    "allwears08@gmail.com": "ALL WEARS新北中和環球店(AKW)",
    "allwears10@gmail.com": "ALL WEARS桃園台茂店(AK)",
    "allwears16@gmail.com": "ALL WEARS台北京站(AKW)",
    "allwears20@gmail.com": "ALL WEARS新北新店裕隆城(AKW)",
    "allwears07@gmail.com": "ALL WEARS高雄夢時代店(AKW)",
    "allwears15@gmail.com": "ALL WEARS台南三井店(AKW)",
    "boylondonx02@gmail.com": "BOYLONDON台中三井店(BN)",
    "boylondonx03@gmail.com": "BOYLONDON台中LaLaport店",
    "boylondonx01@gmail.com": "BOYLONDON林口三井店(BN)",
    "boylondonx04@gmail.com": "BOYLONDON台南三井店(BN)",
    "life8x17@gmail.com": "Life8台中麗寶店(LAKWE)",
    "life8x23@gmail.com": "Life8台中文心秀泰店",
    "life8x30@gmail.com": "Life8台中SOGO店(LAKWE)",
    "life8x38@gmail.com": "Life8台中三井店(LE)",
    "life8x18@gmail.com": "Life8台中新時代店(LW)",
    "life8x63@gmail.com": "Life8台中松竹旗艦店",
    "life8x46@gmail.com": "Life8台中LLP店(LAKWE)",
    "life8x29@gmail.com": "Life8台中老虎城店(LAK)",
    "life8x56@gmail.com": "Life8台中誠品480店(LAK)",
    "life8x16@gmail.com": "Life8苗栗尚順店(LAKE)",
    "lc001@life8.com.tw": "Life8台中勤美旗艦店",
    "life8x34@gmail.com": "Life8台北西門誠品店(LAKW)",
    "life8x45@gmail.com": "Life8新北宏匯廣場店(LAK)",
    "lm040@life8.com.tw": "Life8台北武昌誠品店(LAKWE)",
    "life8x25@gmail.com": "Life8台北信義A11店(L)",
    "life8x36@gmail.com": "Life8宜蘭新月店(LW)",
    "lm038@life8.com.tw": "Life8新北板橋誠品(LW)",
    "life8x60@gmail.com": "Life8新北永和比漾(LW)",
    "life8x64@gmail.com": "Life8台北信義ATT",
    "life8x65@gmail.com": "Life8南港LaLaPort",
    "life8x31@gmail.com": "Life8桃園統領店",
    "life8x41@gmail.com": "Life8新北樹林秀泰店",
    "life8x40@gmail.com": "Life8基隆摩亞時尚店(LAKW)",
    "life8x48@gmail.com": "Life8新北新店裕隆城(LW)",
    "life8x13@gmail.com": "Life8桃園華泰店(LE)",
    "life8x19@gmail.com": "Life8新北中和環球店(LW)",
    "life8x39@gmail.com": "Life8台北南港CITYLI(LW)",
    "life8x43@gmail.com": "Life8新北板橋環球店(LAKW)",
    "life8x57@gmail.com": "Life8台北美麗華店(LAKWE)",
    "life8x58@gmail.com": "Life8新北汐止iFG遠雄廣場(LW)",
    "lm039@life8.com.tw": "Life8新店小碧潭站店(LAKWE)",
    "life8x61@gmail.com": "Life8桃園台茂(LW)",
    "life8x66@gmail.com": "Life8台北遠企",
    "life8x28@gmail.com": "Life8嘉義秀泰店(LAKEM)",
    "life8x20@gmail.com": "Life8高雄夢時代店(LKWEM)",
    "life8x14@gmail.com": "Life8高雄SKM店(LAKE)",
    "life8x21@gmail.com": "Life8屏東環球店(LAK)",
    "life8x42@gmail.com": "Life8高雄岡山樂購店",
    "life8x47@gmail.com": "Life8高雄左營新光(LW)",
    "lm037@life8.com.tw": "Life8高雄義享天地(LAKW)",
    "life8x69@gmail.com": "Life8高雄大遠百",
    "life8x12@gmail.com": "Life8高雄義大A區店(LAKW)",
    "life8x24@gmail.com": "Life8高雄義大C區店(LW)",
    "life8x03@gmail.com": "Life8台南小西門店(LAKW)",
    "life8x32@gmail.com": "Life8台南Focus店(LW)",
    "life8x62@gmail.com": "Life8嘉義耐斯",
    "life8x67@gmail.com": "Life8台南碳佐店(LAKW)",
    "mollifix08@gmail.com": "Mollifix台中LLP(ME)",
    "mollifix01@gmail.com": "Mollifix台北復興SOGO(M)",
    "mollifix03@gmail.com": "Mollifix台北南西新光(ME)",
    "mollifix04@gmail.com": "Mollifix台北京站(ME)",
    "mollifix06@gmail.com": "Mollifix桃園大江(ME)",
    "mollifix07@gmail.com": "Mollifix台北南港City(ME)",
    "mollifix09@gmail.com": "Mollifix高雄漢神巨蛋(ME)",
    "mollifix10@gmail.com": "Mollifix台南小西門(ME)",
    "nonspace02@gmail.com": "NonSpace南港LLP(NZ)",
    "nonspace03@gmail.com": "NonSpace中和環球(NZ)",
    "wildmeet02@gmail.com": "WILDMEET台中麗寶店(WAKN)",
    "wildmeet10@gmail.com": "WILDMEET台中老虎城店(WN)",
    "wildmeet11@gmail.com": "WILDMEET-WM苗栗尚順(WN)",
    "wildmeet08@gmail.com": "WILDMEET桃園華泰店(WAK)",
    "wildmeet09@gmail.com": "WILDMEET台北京站(WN)",
    "wildmeet07@gmail.com": "WILDMEET新光南西(WN)",
    "wildmeet13@gmail.com": "WILDMEET林口三井店(WAKN)",
    "wildmeet14@gmail.com": "WILDMEET新北宏匯(WN)",
    "wildmeet15@gmail.com": "WILDMEET桃園大江購物(WAKN)",
    "wildmeet03@gmail.com": "WILDMEET高雄SKM店(WN)",
    "wildmeet04@gmail.com": "WILDMEET嘉義秀泰店(WN)",
    "wildmeet06@gmail.com": "WILDMEET屏東環球店(WN)",
    "wildmeet12@gmail.com": "WILDMEET新光台南中山(WAK)"
    # 👉 其他帳號可依需要繼續補上
}


# -------------------- 頁面設定 --------------------
st.set_page_config(page_title="米斯特績效考核查詢", page_icon="📊")

# -------------------- Google 登入驗證 --------------------
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

# -------------------- 登入流程處理 --------------------
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
                user_name = EMAIL_TO_NAME.get(email, email)
                st.success(f"👋 Hi {user_name}，歡迎使用查詢系統！")
            else:
                st.error("❌ 此帳號未授權存取此應用程式。")
                st.stop()
        else:
            st.error("⚠️ 無法取得 access token。請重新登入。")
            st.stop()
    else:
        login_url = get_login_url()
        st.markdown("<h3>📊 米斯特 門市 工作績效月考核查詢系統</h3>", unsafe_allow_html=True)
        
        st.markdown(f"""
        ### 👋 Hello，米斯特夥伴!  
        請先登入 Google 帳號完成驗證，即可查詢每月考核成績：  
        👉 [立即登入]({login_url})
        
        ---
        
        🔺 **門市**：請使用門市 Gmail 帳號登入  
        🔺 **營運主管**：請使用個人 Gmail 帳號登入  
        🔺 每月考核結果會在 **每月5日更新**，如有疑慮，請於 **當月8日（含）前**完成申覆，逾期未提出者，視為認同本次考核結果。
        
        📌 **例：**  
        8月5日更新「7月考核」 → 須於 8月8日 前完成申覆。
        """, unsafe_allow_html=True)
        
        st.stop()


# -------------------- 📂 資料來源設定 --------------------
FILE_URL = "https://raw.githubusercontent.com/ainstaccc/kpi-checker/main/2025.06_MST-PA.xlsx"

# -------------------- 📥 載入 Excel 各工作表 --------------------
@st.cache_data(ttl=3600)
def load_data():
    try:
        # 嘗試讀取 Excel 檔案
        xls = pd.ExcelFile(FILE_URL, engine="openpyxl")
        sheet_names = xls.sheet_names

        # 檢查是否有缺少必要的工作表
        required_sheets = ["門店 考核總表", "人效分析", "店長副店 考核明細", "店員儲備 考核明細", "等級分布"]
        for name in required_sheets:
            if name not in sheet_names:
                raise ValueError(f"❌ 缺少必要工作表：{name}")

        # 各工作表轉成 dataframe（帶標題列）
        df_summary = xls.parse("門店 考核總表", header=1)
        df_eff = xls.parse("人效分析", header=1)
        df_mgr = xls.parse("店長副店 考核明細", header=1)
        df_staff = xls.parse("店員儲備 考核明細", header=1)
        df_dist = xls.parse("等級分布", header=None, nrows=15, usecols="A:N")

        # 嘗試讀取月份資訊（從第一列第一欄抓取）
        try:
            summary_month = xls.parse("門店 考核總表", header=None, nrows=1).iloc[0, 0]
        except Exception:
            summary_month = "未知月份"

        return df_summary, df_eff, df_mgr, df_staff, df_dist, summary_month

    except Exception as e:
        st.error(f"❌ 資料載入失敗：{e}")
        return None, None, None, None, None, None

def format_staff_id(df):
    if "員編" in df.columns:
        df["員編"] = df["員編"].apply(lambda x: str(int(float(x))).zfill(8) if pd.notnull(x) else "")
    return df

# -------------------- 🧽 格式化人效分析欄位 --------------------
def format_eff(df):
    if df is None or df.empty:
        return pd.DataFrame()
    df = df.copy()
    # 🧽 清洗欄位名稱（移除多餘空白或換行）
    df.columns = df.columns.str.replace(r"\s+", " ", regex=True).str.strip()

    # 員編：補足8碼 + 移除小數點，確保是字串顯示
    if "員編" in df.columns:
        df["員編"] = df["員編"].apply(lambda x: str(int(float(x))).zfill(8) if pd.notnull(x) else "")


    # 金額欄位：轉為千分位整數（純顯示用的文字格式）
    for col in ["個績目標", "個績貢獻"]:
        if col in df.columns:
            df[col] = (
                df[col].astype(str)
                .str.replace(",", "")
                .astype(float)
                .round(0)
                .map(lambda x: f"{int(x):,}" if pd.notnull(x) else "")
            )

    # 個績達成%：統一 xx.x% 格式
    if "個績達成%" in df.columns:
        df["個績達成%"] = df["個績達成%"].apply(
            lambda x: f"{float(str(x).replace('%', '')):.1f}%" if pd.notnull(x) else x
        )

    # 客單價欄位：顯示為整數
    for col in ["品牌 客單價", "個人 客單價"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").round(0).astype("Int64")

    # 客單相對績效：轉為百分比整數（加%）
    if "客單 相對績效" in df.columns:
        df["客單 相對績效"] = df["客單 相對績效"].apply(
            lambda x: f"{round(x * 100)}%" if pd.notnull(x) else x
        )

    # 結帳會員率 / 會員相對績效：轉為百分比整數（加%）
    for col in ["品牌 結帳會員率", "個人 結帳會員率", "會員 相對績效"]:
        if col in df.columns:
            df[col] = df[col].apply(
                lambda x: f"{round(x * 100)}%" if pd.notnull(x) else x
            )

    return df


# -------------------- 🚀 主程式入口 --------------------
def main():
    # 網頁標題
    st.markdown("<h3>📊 米斯特 門市 工作績效月考核查詢系統</h3>", unsafe_allow_html=True)

    # 載入資料
    df_summary, df_eff, df_mgr, df_staff, df_dist, summary_month = load_data()
    if df_summary is None:
        st.stop()

    # ---------------- 查詢條件區塊 ----------------
    with st.expander("🔍 查詢條件", expanded=True):
        st.markdown("**🔺查詢條件任一欄即可，避免多重條件造成查詢錯誤。**")
        col1, col2 = st.columns(2)

        # 選擇區域主管、部門編號、月份
        area = col1.selectbox("區域/區主管", options=[
            "", "李政勳", "林宥儒", "羅婉心", "王建樹", "楊茜聿",
            "陳宥蓉", "吳岱侑", "翁聖閔", "黃啓周", "栗晉屏", "王瑞辰"
        ])
        dept_code = col2.text_input("部門編號/門店編號")
        month = st.selectbox("查詢月份", options=["2025/06"])  # 可預留擴充月份選單

    # ---------------- 查詢按鈕觸發 ----------------
    st.markdown("<br><br>", unsafe_allow_html=True)

    if st.button("🔎 查詢", type="primary"):
        try:
            # 依條件篩選資料（四張表格）
            df_result = df_summary.copy()
            df_eff_result = df_eff.copy()
            df_mgr_result = df_mgr.copy()
            df_staff_result = df_staff.copy()

            if area:
                df_result = df_result[df_result["區主管"] == area]
                df_eff_result = df_eff_result[df_eff_result["區主管"] == area]
                df_mgr_result = df_mgr_result[df_mgr_result["區主管"] == area]
                df_staff_result = df_staff_result[df_staff_result["區主管"] == area]

            if dept_code:
                df_result = df_result[df_result["部門編號"] == dept_code]
                df_eff_result = df_eff_result[df_eff_result["部門編號"] == dept_code]
                df_mgr_result = df_mgr_result[df_mgr_result["部門編號"] == dept_code]
                df_staff_result = df_staff_result[df_staff_result["部門編號"] == dept_code]

            df_result = format_staff_id(df_result)
            df_mgr_result = format_staff_id(df_mgr_result)
            df_staff_result = format_staff_id(df_staff_result)


            # ---------------- 成績參考圖示 ----------------
            st.image("https://raw.githubusercontent.com/ainstaccc/kpi-checker/main/2025.06_grade.jpg", use_column_width=True)

            # ---------------- 顯示表格：門店考核總表 ----------------
            st.markdown("## 🧾 門店考核總表")
            st.markdown("<span style='color:red;'>🔺紅字顯示：考核項目分數＜80、管理項目分數＜25</span>", unsafe_allow_html=True)
            
            df_display = df_result.iloc[:, 2:11].copy()
            
            # 員編：補足8碼 + 去除小數點
            if "員編" in df_display.columns:
                df_display["員編"] = df_display["員編"].apply(lambda x: str(int(float(x))).zfill(8) if pd.notnull(x) else "")
            
            # 考核項目分數：保留1位小數
            if "考核項目分數" in df_display.columns:
                df_display["考核項目分數"] = pd.to_numeric(df_display["考核項目分數"], errors="coerce").round(1)
            
            # 管理項目分數：保留整數
            if "管理項目分數" in df_display.columns:
                df_display["管理項目分數"] = pd.to_numeric(df_display["管理項目分數"], errors="coerce").astype("Int64")
            
            # 紅字樣式條件
            def highlight_scores(val, col):
                try:
                    num = float(val)
                    if col == "考核項目分數" and num < 80:
                        return "color: red;"
                    elif col == "管理項目分數" and num < 25:
                        return "color: red;"
                    else:
                        return ""
                except:
                    return ""
            
            cols_to_highlight = ["考核項目分數", "管理項目分數"]
            styled = df_display.style.apply(
                lambda col: [highlight_scores(v, col.name) for v in col],
                subset=cols_to_highlight
            ).format({
                "考核項目分數": "{:.1f}"  # 顯示到小數1位
            })
            
            st.markdown(f"共查得：{len(df_display)} 筆")
            st.dataframe(styled, use_container_width=True)

            # ---------------- 顯示表格：人效分析 ----------------
            st.markdown("## 👥 人效分析")
            st.markdown("<span style='color:red;'>🔺計分指標：個績達成%、客單相對績效、會員相對績效</span>", unsafe_allow_html=True)
            
            df_eff_fmt = format_eff(df_eff_result)
            st.markdown(f"共查得：{len(df_eff_fmt)} 筆")
            st.dataframe(df_eff_fmt, use_container_width=True)

            # ---------------- 顯示表格：店長/副店考核明細 ----------------
            st.markdown("## 👔 店長/副店 考核明細")
            st.markdown("<span style='color:red;'>🔺分數小計項：總分、業績項目分數、管理分數_人資、財務、商控、服務</span>", unsafe_allow_html=True)
            
            df_mgr_display = pd.concat([
                df_mgr_result.iloc[:, 1:7],
                df_mgr_result.iloc[:, 11:28]
            ], axis=1)
            st.markdown(f"共查得：{len(df_mgr_display)} 筆")
            st.dataframe(df_mgr_display, use_container_width=True)


            # ---------------- 顯示表格：店員/儲備考核明細 ----------------
            st.markdown("## 👟 店員/儲備 考核明細")
            st.markdown("<span style='color:red;'>🔺分數小計項：總分、業績項目分數、管理分數_人資、財務、商控、服務</span>", unsafe_allow_html=True)
            
            df_staff_display = pd.concat([
                df_staff_result.iloc[:, 1:7],
                df_staff_result.iloc[:, 11:28]
            ], axis=1)
            st.markdown(f"共查得：{len(df_staff_display)} 筆")
            st.dataframe(df_staff_display, use_container_width=True)


            # ---------------- 備註提醒 ----------------
            st.markdown("<p style='color:red;font-weight:bold;font-size:16px;'>※如對分數有疑問，請洽區主管/品牌經理說明。</p>", unsafe_allow_html=True)

        except Exception as e:
            st.error(f"❌ 查詢過程發生錯誤：{e}")

# -------------------- 程式啟動 --------------------
if __name__ == "__main__":
    main()

