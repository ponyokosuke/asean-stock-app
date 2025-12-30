import streamlit as st
import pandas as pd
import time
import io
import os
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from dotenv import load_dotenv

# ローカル用（.envがあれば読み込むが、クラウドでは無視される）
load_dotenv()

# Import existing logic
import data_processor

# Page config
st.set_page_config(page_title="ASEAN Stock Analyzer", layout="wide")

# --- 🔑 SECRETS & ENV CONFIG ---
# クラウド(Secrets機能)とローカル(.env)の両方から取得を試みる
def get_secret(key):
    if key in st.secrets:
        return st.secrets[key]
    return os.environ.get(key)

env_password = get_secret("APP_PASSWORD")
env_gemini_key = get_secret("GEMINI_API_KEY")

# --- 🔐 PASSWORD AUTHENTICATION ---
def password_entered():
    if st.session_state["password"] == env_password:
        st.session_state["password_correct"] = True
        del st.session_state["password"]
    else:
        st.session_state["password_correct"] = False

def check_password():
    if not env_password: # パスワード設定がない場合はスキップ（デバッグ用）
        return True
    if "password_correct" not in st.session_state:
        st.text_input("Please enter the password:", type="password", on_change=password_entered, key="password")
        return False
    elif not st.session_state["password_correct"]:
        st.text_input("Please enter the password:", type="password", on_change=password_entered, key="password")
        st.error("😕 Password incorrect")
        return False
    return True

if not check_password():
    st.stop()

# --- 💾 SESSION STATE ---
if "excel_buffer" not in st.session_state:
    st.session_state.excel_buffer = None
if "final_df" not in st.session_state:
    st.session_state.final_df = None

# --- 🛠 HELPERS ---
def robust_clean_columns(df):
    """どんな時でも重複列を完全に消し去る"""
    return df.loc[:, ~df.columns.duplicated()].copy()

# --- MAIN APP ---
st.title("📊 ASEAN Stock Financial & AI Analysis Tool")

with st.sidebar:
    st.header("Settings")
    if env_gemini_key:
        os.environ["GEMINI_API_KEY"] = env_gemini_key
        from google import genai
        data_processor.client = genai.Client(api_key=env_gemini_key)
        st.success("API Key Loaded ✅")
    else:
        api_key = st.text_input("Gemini API Key (Required if not set in Secrets)", type="password")
        if api_key:
            from google import genai
            data_processor.client = genai.Client(api_key=api_key)

uploaded_file = st.file_uploader("Upload Stock List (CSV)", type=["csv"])
use_sample = st.checkbox("Use default list (asean_list.csv) if no file is available")

if st.button("Start Analysis 🚀"):
    target_csv = uploaded_file if uploaded_file else ("asean_list.csv" if use_sample else None)
    
    if target_csv is None:
        st.error("Please upload a CSV file.")
    else:
        try:
            st.session_state.excel_buffer = None
            df_input = pd.read_csv(target_csv, header=None)
            codes = df_input[0].astype(str).tolist()
            
            status_text = st.empty()
            progress_bar = st.progress(0)
            
            all_results = []
            for i, code in enumerate(codes):
                code = code.strip()
                status_text.text(f"Processing ({i+1}/{len(codes)}): {code}...")
                progress_bar.progress((i + 1) / (len(codes) + 1))
                raw_data = data_processor.get_stock_data(code)
                if raw_data:
                    all_results.append(data_processor.extract_data(code, raw_data))
                time.sleep(0.1)
            
            if all_results:
                status_text.text("🤖 Running AI Analysis...")
                all_results = data_processor.batch_analyze_segments(all_results)
                
                # データフレーム化
                df = pd.DataFrame(all_results)
                df = robust_clean_columns(df) # 重複排除
                
                # 整形
                df = data_processor.format_for_excel(df)
                df = robust_clean_columns(df) # 再度排除
                
                # 列の追加
                df["Ref"] = range(1, len(df) + 1)
                empty_cols = ["Taka's comments", "Remarks", "Visited (V) / Meeting Proposal (MP)", "Access", "Last Communications", "Category Classification/\nShareInvestor", "Incorporated\n (IN / Year)", "Category Classification/SGX", "Sector & Industry/ SGX"]
                for col in empty_cols:
                    if col not in df.columns:
                        df[col] = ""
                
                df["Listed 'o' / Non Listed \"x\""] = "o"

                yesterday = datetime.now() - timedelta(days=1)
                yesterday_str = yesterday.strftime("%b %d")
                final_stock_price_col = f"Stock Price ({yesterday_str}, Closing)"
                final_rate_col = f"Exchange Rate (to SGD) ({yesterday_str}, Closing)"
                
                # リネーム
                df = df.rename(columns={"Stock Price": final_stock_price_col, "Exchange Rate": final_rate_col})
                if "Number of Employee" in df.columns:
                    df = df.rename(columns={"Number of Employee": "Number of Employee Current"})

                # 並び替え（エラー対策の要）
                target_order = ["Ref", "Name of Company", "Code", "Listed 'o' / Non Listed \"x\"", "Taka's comments", "Remarks", "Visited (V) / Meeting Proposal (MP)", "Website", "Major Shareholders", "Currency", final_rate_col, "FY", "REVENUE SGD('000)", "Segments", "PROFIT ('000)", "GROSS PROFIT ('000)", "OPERATING PROFIT ('000)", "NET PROFIT (Group) ('000)", "NET PROFIT (Shareholders) ('000)", "Minority Interest ('000)", "Shareholders' Equity ('000)", "Total Equity ('000)", "TOTAL ASSET ('000)", "Debt/Equity(%)", "Loan ('000)", "Loan/Equity (%)", final_stock_price_col, "Shares Outstanding ('000)", "Market Cap ('000)", "Summary of Business", "Chairman / CEO", "Address", "Contact No.", "Access", "Last Communications", "Number of Employee Current", "Category Classification/YahooFin", "Sector & Industry/YahooFin", "Category Classification/\nShareInvestor", "Incorporated\n (IN / Year)", "Category Classification/SGX", "Sector & Industry/ SGX"]
                
                # 重複を最終排除した上で、足りない列を埋める
                df = robust_clean_columns(df)
                for col in target_order:
                    if col not in df.columns:
                        df[col] = ""
                
                # 並び替え実行
                df = df.reindex(columns=target_order)

                # Excel作成
                temp_buffer = io.BytesIO()
                df.to_excel(temp_buffer, index=False)
                temp_buffer.seek(0)
                wb = load_workbook(temp_buffer)
                for cell in wb.active[1]:
                    cell.fill = PatternFill(start_color="fefe99", end_color="fefe99", fill_type="solid")
                    cell.font = Font(bold=True)
                
                final_buffer = io.BytesIO()
                wb.save(final_buffer)
                
                st.session_state.excel_buffer = final_buffer.getvalue()
                st.session_state.final_df = df
                st.session_state.output_filename = f"asean_financial_data_{datetime.today().strftime('%Y-%m-%d')}.xlsx"
                
                progress_bar.progress(100)
                status_text.text("✅ Analysis completed!")

        except Exception as e:
            st.error(f"Error: {e}")

# --- 📥 DOWNLOAD AREA ---
if st.session_state.excel_buffer:
    st.divider()
    st.download_button(
        label="📥 Download Excel File",
        data=st.session_state.excel_buffer,
        file_name=st.session_state.output_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.dataframe(st.session_state.final_df)
