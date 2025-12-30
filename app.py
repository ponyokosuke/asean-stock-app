import streamlit as st
import pandas as pd
import time
import io
import os
from datetime import datetime
from openpyxl.styles import PatternFill, Font, Alignment

# Import existing logic
import data_processor

# Page config
st.set_page_config(page_title="ASEAN Stock Analyzer", layout="wide")

st.title("📊 ASEAN Stock Financial & AI Analysis Tool")
st.markdown("Upload a stock list (CSV) to integrate Yahoo Finance data with Gemini AI analysis and export to Excel.")

# --- Sidebar: Settings ---
with st.sidebar:
    st.header("Settings")
    
    # API Key Input
    api_key = st.text_input("Gemini API Key", type="password", help="Leave blank if configured in Secrets")
    
    if api_key:
        os.environ["GEMINI_API_KEY"] = api_key
        # Re-initialize client
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

                # Rename Date Columns
                # Calculate YESTERDAY'S date for "Previous Close"
                yesterday = datetime.now() - pd.Timedelta(days=1)
                yesterday_str = yesterday.strftime("%b %d")
                
                final_stock_price_col = f"Stock Price (Prev. Close as of {yesterday_str})"
                if "Stock Price" in df.columns:
                    df = df.rename(columns={"Stock Price": final_stock_price_col})
                    
                final_rate_col = f"Exchange Rate (to SGD) (Prev. Close as of {yesterday_str})"
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

                # Save to Buffer with Styling
                buffer = io.BytesIO()
                
                # Using openpyxl engine to access the workbook for styling
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Sheet1')
                    
                    # Access the worksheet
                    workbook = writer.book
                    worksheet = writer.sheets['Sheet1']
                    
                    # Define Styles
                    header_fill = PatternFill(start_color="fefe99", end_color="fefe99", fill_type="solid")
                    header_font = Font(bold=True)
                    right_align = Alignment(horizontal='right')
                    
                    # Iterate through the header row (Row 1)
                    for cell in worksheet[1]:
                        cell.fill = header_fill
                        cell.font = header_font
                        
                        # Apply Number Formats based on column name
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
                            
                        # Apply format to the data column below the header
                        if number_format or apply_alignment:
                             for row in worksheet.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
                                for cell_data in row:
                                    if apply_alignment:
                                        cell_data.alignment = right_align
                                    if number_format:
                                        cell_data.number_format = number_format

                buffer.seek(0)
                
                # Download Button
                file_name = f"asean_financial_data_{datetime.today().strftime('%Y-%m-%d')}.xlsx"
                
                st.success("Analysis Complete! Download the Excel file below.")
                st.download_button(
                    label="📥 Download Excel File",
                    data=buffer,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                # Data Preview
                st.subheader("Data Preview")
                st.dataframe(df)

        except Exception as e:
            st.error(f"An error occurred: {e}")