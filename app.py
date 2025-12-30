import streamlit as st
import pandas as pd
import time
import io
import os
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

# Import existing logic
import data_processor

# Page config (Must be the first Streamlit command)
st.set_page_config(page_title="ASEAN Stock Analyzer", layout="wide")

# --- 🔐 PASSWORD AUTHENTICATION START ---
def check_password():
    """Returns `True` if the user had the correct password."""

    def password_entered():
        """Checks whether a password entered by the user is correct."""
        if st.session_state["password"] == st.secrets["APP_PASSWORD"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # don't store password
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        # First run, show input for password.
        st.text_input(
            "Please enter the password to access this app:", 
            type="password", 
            on_change=password_entered, 
            key="password"
        )
        return False
    elif not st.session_state["password_correct"]:
        # Password was incorrect, show input + error.
        st.text_input(
            "Please enter the password to access this app:", 
            type="password", 
            on_change=password_entered, 
            key="password"
        )
        st.error("😕 Password incorrect")
        return False
    else:
        # Password was correct.
        return True

if not check_password():
    st.stop()  # Do not run the rest of the app if password is incorrect
# --- 🔐 PASSWORD AUTHENTICATION END ---


# === ⬇️ MAIN APP CONTENT STARTS HERE ⬇️ ===

st.title("📊 ASEAN Stock Financial & AI Analysis Tool")
st.markdown("Upload a stock list (CSV) to integrate Yahoo Finance data with Gemini AI analysis and export to Excel.")

# --- Sidebar: Settings ---
with st.sidebar:
    st.header("Settings")
    
    # API Key Input
    if "GEMINI_API_KEY" in st.secrets:
        os.environ["GEMINI_API_KEY"] = st.secrets["GEMINI_API_KEY"]
        from google import genai
        data_processor.client = genai.Client(api_key=os.environ["GEMINI_API_KEY"])
        st.success("API Key loaded from Secrets ✅")
    else:
        api_key = st.text_input("Gemini API Key", type="password")
        if api_key:
            os.environ["GEMINI_API_KEY"] = api_key
            from google import genai
            data_processor.client = genai.Client(api_key=api_key)

# --- File Uploader ---
uploaded_file = st.file_uploader("Upload Stock List (CSV)", type=["csv"])

# Sample Data Option
use_sample = st.checkbox("Use default list (asean_list.csv) if no file is available")

if st.button("Start Analysis 🚀"):
    target_csv = None
    
    if uploaded_file is not None:
        target_csv = uploaded_file
    elif use_sample:
        target_csv = "asean_list.csv"
    
    if target_csv is None:
        st.error("Please upload a CSV file or select the default list.")
    else:
        # --- Start Analysis ---
        status_text = st.empty()
        progress_bar = st.progress(0)
        
        try:
            # Read CSV
            if isinstance(target_csv, str):
                df_input = pd.read_csv(target_csv, header=None)
            else:
                df_input = pd.read_csv(target_csv, header=None)
                
            codes = df_input[0].astype(str).tolist()
            total_codes = len(codes)
            
            st.info(f"Retrieving data for {total_codes} stocks...")
            
            all_results = []
            
            # 1. Data Retrieval Loop
            for i, code in enumerate(codes):
                code = code.strip()
                status_text.text(f"Processing ({i+1}/{total_codes}): Retrieving data for {code}...")
                progress_bar.progress((i + 1) / (total_codes + 1))
                
                raw_data = data_processor.get_stock_data(code)
                
                if raw_data:
                    processed_data = data_processor.extract_data(code, raw_data)
                    all_results.append(processed_data)
                
                time.sleep(0.5) 
            
            # 2. AI Analysis
            if all_results:
                status_text.text("🤖 Running AI Segment Analysis... (This may take some time)")
                all_results = data_processor.batch_analyze_segments(all_results)
            
            progress_bar.progress(100)
            status_text.text("✅ All processes completed!")
            
            # 3. Excel Creation
            if all_results:
                df = pd.DataFrame(all_results)
                
                # Format Data
                df = df.loc[:, ~df.columns.duplicated()]
                df = data_processor.format_for_excel(df)
                
                if "Sector /Industry" in df.columns:
                    df = df.drop(columns=["Sector /Industry"])
                
                df["Ref"] = range(1, len(df) + 1)
                
                # Add Empty Columns
                empty_cols = [
                    "Taka's comments", "Remarks", "Visited (V) / Meeting Proposal (MP)",
                    "Access", "Last Communications", "Category Classification/\nShareInvestor", 
                    "Incorporated\n (IN / Year)", "Category Classification/SGX", "Sector & Industry/ SGX"
                ]
                for col in empty_cols:
                    df[col] = ""
                
                df["Listed 'o' / Non Listed \"x\""] = "o"

                # Rename Date Columns (Using Yesterday logic like main.py)
                yesterday = datetime.now() - timedelta(days=1)
                yesterday_str = yesterday.strftime("%b %d")
                
                final_stock_price_col = f"Stock Price ({yesterday_str}, Closing)"
                if "Stock Price" in df.columns:
                    df = df.rename(columns={"Stock Price": final_stock_price_col})
                    
                final_rate_col = f"Exchange Rate (to SGD) ({yesterday_str}, Closing)"
                if "Exchange Rate" in df.columns:
                    df = df.rename(columns={"Exchange Rate": final_rate_col})

                # Reorder Columns
                target_order = [
                    "Ref", "Name of Company", "Code", "Listed 'o' / Non Listed \"x\"",
                    "Taka's comments", "Remarks", "Visited (V) / Meeting Proposal (MP)",
                    "Website", "Major Shareholders", "Currency", 
                    final_rate_col, 
                    "FY", "REVENUE SGD('000)", "Segments", "PROFIT ('000)",
                    "GROSS PROFIT ('000)", "OPERATING PROFIT ('000)",
                    "NET PROFIT (Group) ('000)", "NET PROFIT (Shareholders) ('000)",
                    "Minority Interest ('000)", "Shareholders' Equity ('000)",
                    "Total Equity ('000)", "TOTAL ASSET ('000)", "Debt/Equity(%)",
                    "Loan ('000)", "Loan/Equity (%)",
                    final_stock_price_col, 
                    "Shares Outstanding ('000)", "Market Cap ('000)",
                    "Summary of Business", "Chairman / CEO", "Address", "Contact No.",
                    "Access", "Last Communications", "Number of Employee Current",
                    "Category Classification/YahooFin", "Sector & Industry/YahooFin",
                    "Category Classification/\nShareInvestor", "Incorporated\n (IN / Year)",
                    "Category Classification/SGX", "Sector & Industry/ SGX"
                ]
                
                for col in target_order:
                    if col not in df.columns:
                        df[col] = ""
                
                if "Number of Employee" in df.columns:
                    df = df.rename(columns={"Number of Employee": "Number of Employee Current"})
                
                df = df.loc[:, ~df.columns.duplicated()]
                df = df.reindex(columns=target_order)

                # --- ★ここが修正ポイント: 確実な保存方法 ---
                
                # 1. まずデータだけのExcelを一時メモリ(temp_buffer)に保存
                temp_buffer = io.BytesIO()
                df.to_excel(temp_buffer, index=False)
                temp_buffer.seek(0) # 巻き戻し

                # 2. openpyxlで読み込み
                wb = load_workbook(temp_buffer)
                ws = wb.active

                # 3. スタイルを適用
                header_fill = PatternFill(start_color="fefe99", end_color="fefe99", fill_type="solid")
                header_font = Font(bold=True)
                right_align = Alignment(horizontal='right')
                
                for cell in ws[1]:
                    # ヘッダーの色設定
                    cell.fill = header_fill
                    cell.font = header_font
                    
                    col_name = str(cell.value)
                    col_idx = cell.column
                    
                    number_format = None
                    apply_alignment = False
                    
                    if "('000)" in col_name:
                        number_format = '#,##0;(#,##0)'
                        apply_alignment = True
                    elif "(%)" in col_name or "%" in col_name:
                        number_format = '0.00%'
                        apply_alignment = True
                    elif col_name == "FY":
                        apply_alignment = True
                    elif "Stock Price" in col_name:
                        number_format = '#,##0.000'
                        apply_alignment = True
                    elif "Exchange Rate" in col_name:
                        number_format = '0.0000'
                        apply_alignment = True
                        
                    if number_format or apply_alignment:
                         for row in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
                            for cell_data in row:
                                if apply_alignment:
                                    cell_data.alignment = right_align
                                if number_format:
                                    cell_data.number_format = number_format

                # 4. 完成品を最終メモリ(final_buffer)に保存
                final_buffer = io.BytesIO()
                wb.save(final_buffer)
                final_buffer.seek(0) # 巻き戻し（ダウンロードのために必須）

                # 5. ダウンロードボタン
                file_name = f"asean_financial_data_{datetime.today().strftime('%Y-%m-%d')}.xlsx"
                
                st.success("Analysis Complete! Download the Excel file below.")
                st.download_button(
                    label="📥 Download Excel File",
                    data=final_buffer,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                # Data Preview
                st.subheader("Data Preview")
                st.dataframe(df)

        except Exception as e:
            st.error(f"An error occurred: {e}")
