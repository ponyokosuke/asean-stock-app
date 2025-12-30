import pandas as pd
import yfinance as yf
from datetime import datetime
import time
import json
import os
from google import genai
from dotenv import load_dotenv

load_dotenv()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

client = None
if GEMINI_API_KEY:
    client = genai.Client(api_key=GEMINI_API_KEY)

# --- 1. AIによるセグメント分析 ---
def batch_analyze_segments(all_results_list):
    if not client:
        print("  ⚠️ APIキー(.env)が見つからないため、AI分析をスキップします")
        return all_results_list

    targets = [item for item in all_results_list if item.get('Summary of Business')]
    
    if not targets:
        return all_results_list

    print(f"\n🤖 Gemini AI分析開始: 対象 {len(targets)} 件をまとめて処理します (バッチ処理)...")
    
    batch_size = 20
    model_name = 'gemini-2.5-flash'
    
    for i in range(0, len(targets), batch_size):
        batch = targets[i : i + batch_size]
        current_count = min(i + batch_size, len(targets))
        print(f"  - バッチ処理中: {i+1}〜{current_count} 件目...")

        input_text = ""
        for item in batch:
            summary_snippet = str(item['Summary of Business'])[:500].replace("\n", " ")
            input_text += f"Code: {item['Code']}\nSummary: {summary_snippet}...\n---\n"

        prompt = f"""
        You are a financial analyst. I will provide business summaries for multiple companies.
        Extract the main 'Business Segments' for EACH company based on the summary.

        # Input Data
        {input_text}
        
        # Output Rules
        - Return ONLY a valid JSON object.
        - The keys must be the stock 'Code'.
        - The values must be the 'Business Segments' (comma separated string, clear and concise).
        - If segments are not clearly stated, summarize the main business areas in 3-4 words.
        - Example JSON Format:
        {{
            "4863.KL": "Telecommunication Services, Digital Solutions",
            "0021.KL": "Payment Services, Solution Services"
        }}
        """

        try:
            response = client.models.generate_content(
                model=model_name,
                contents=prompt
            )
            response_text = response.text.strip()
            
            if "```json" in response_text:
                response_text = response_text.split("```json")[1].split("```")[0].strip()
            elif "```" in response_text:
                response_text = response_text.split("```")[1].split("```")[0].strip()
            
            segments_map = json.loads(response_text)

            for item in batch:
                code = item['Code']
                if code in segments_map:
                    item['Segments'] = segments_map[code]
            
            time.sleep(1)

        except Exception as e:
            print(f"  ⚠️ バッチ処理エラー (このバッチはスキップします): {e}")

    print("✅ AI分析完了\n")
    return all_results_list


# --- 2. データ取得関数 ---
def get_stock_data(code):
    try:
        ticker = yf.Ticker(code)
        try:
            info = ticker.info
        except:
            return None
            
        if not info:
            return None

        raw_data = {
            "info": info,
            "balance_sheet": ticker.balance_sheet,
            "financials": ticker.financials,
            "major_holders": ticker.major_holders,
            "institutional_holders": ticker.institutional_holders
        }
        return raw_data
    except Exception as e:
        print(f"  データ取得エラー: {e}")
        return None

# ★変更: 為替レートも「前日終値」を優先取得
def get_exchange_rate(from_currency):
    """
    指定された通貨からSGDへの為替レート (SGD/外貨) を取得します。
    整合性を保つため、株価と同様に 'previousClose' を優先します。
    """
    if not from_currency or from_currency == "SGD":
        return 1.0
    
    if from_currency == "RMB (CNY)":
        currency_code = "CNY"
    else:
        currency_code = from_currency

    pair = f"{currency_code}SGD=X"
    
    try:
        ticker = yf.Ticker(pair)
        
        # 1. まず info から previousClose (前日終値) を取得
        rate = ticker.info.get('previousClose')
        
        # 2. 取れなければ履歴データの最新終値で代用
        if rate is None:
            hist = ticker.history(period="5d")
            if not hist.empty:
                rate = hist['Close'].iloc[-1]
            else:
                return "N/A"
        
        return rate
    except:
        return "N/A"


# --- 3. データの整形・抽出 ---
def format_shareholders(holders_data, data_type="institutional"):
    if holders_data is None or holders_data.empty:
        return None

    result_lines = []
    try:
        if data_type in ["institutional", "insider"]:
            name_col = None
            pct_col = None
            for col in holders_data.columns:
                col_lower = str(col).lower()
                if "holder" in col_lower or "insider" in col_lower:
                    name_col = col
                if "%" in col_lower or "pct" in col_lower or "out" in col_lower:
                    pct_col = col
            
            if name_col and pct_col:
                for index, row in holders_data.head(10).iterrows():
                    name = str(row[name_col])
                    val = row[pct_col]
                    try:
                        val_float = float(val)
                        if val_float < 1.0: 
                            val_float = val_float * 100
                        line = f"{name}: {val_float:.2f}%"
                    except:
                        line = f"{name}: {val}"
                    result_lines.append(line)

        elif data_type == "major":
            if len(holders_data.columns) == 1:
                for index, row in holders_data.head(5).iterrows():
                    desc = str(index)
                    val = row.iloc[0]
                    try:
                        if isinstance(val, str) and '%' in val:
                             line = f"{desc}: {val}"
                        else:
                             val_float = float(val)
                             if val_float <= 1.0:
                                 line = f"{desc}: {val_float:.2%}"
                             else:
                                 line = f"{desc}: {val}"
                    except:
                        line = f"{desc}: {val}"
                    result_lines.append(line)

            elif len(holders_data.columns) >= 2:
                for index, row in holders_data.head(5).iterrows():
                    val_raw = row.iloc[0]
                    desc_raw = row.iloc[1]
                    if len(str(desc_raw)) > len(str(val_raw)):
                        name = str(desc_raw)
                        val = str(val_raw)
                    else:
                        name = str(val_raw)
                        val = str(desc_raw)
                    if name:
                        line = f"{name}: {val}"
                        result_lines.append(line)
    except Exception as e:
        return f"Error parsing: {e}"

    if not result_lines:
        return None
    return "\n".join(result_lines)


def extract_data(code, raw_data):
    info = raw_data.get("info", {})
    bs = raw_data.get("balance_sheet")
    inc = raw_data.get("financials")
    
    inst_holders = raw_data.get("institutional_holders")
    major_holders = raw_data.get("major_holders")
    
    shareholder_text = "Not Available"
    text_major = format_shareholders(major_holders, "major")
    if text_major:
        shareholder_text = text_major
    else:
        text_inst = format_shareholders(inst_holders, "institutional")
        if text_inst:
            shareholder_text = text_inst

    latest_date = None
    if bs is not None and not bs.empty:
        latest_date = bs.columns[0]
    elif inc is not None and not inc.empty:
        latest_date = inc.columns[0]

    def get_fin_value(df, key):
        if (df is not None and not df.empty and 
            key in df.index and 
            latest_date is not None and 
            latest_date in df.columns):
            return df.loc[key, latest_date]
        return 0

    revenue = get_fin_value(inc, "Total Revenue")
    pretax_income = get_fin_value(inc, "Pretax Income")
    operating_income = get_fin_value(inc, "Operating Income")
    gross_profit = get_fin_value(inc, "Gross Profit")
    profit = pretax_income if pretax_income != 0 else operating_income

    net_profit_owners = get_fin_value(inc, "Net Income")
    if net_profit_owners == 0:
        net_profit_owners = get_fin_value(inc, "Net Income Common Stock")

    net_profit_group = get_fin_value(inc, "Net Income Including Noncontrolling Interests")
    if net_profit_group == 0:
         net_profit_group = get_fin_value(inc, "Net Income Continuous Operations")
    
    minority_interest = get_fin_value(bs, "Minority Interest")
    if net_profit_group == 0 and net_profit_owners != 0:
        net_profit_group = net_profit_owners

    stockholders_equity = get_fin_value(bs, "Stockholders Equity")
    total_assets = get_fin_value(bs, "Total Assets")
    total_equity = get_fin_value(bs, "Total Equity Gross Minority Interest")
    if total_equity == 0 and stockholders_equity != 0:
        total_equity = stockholders_equity + minority_interest

    current_loan = get_fin_value(bs, "Current Debt")
    non_current_loan = get_fin_value(bs, "Long Term Debt")
    loan = get_fin_value(bs, "Total Debt")
    capital_lease = get_fin_value(bs, "Capital Lease Obligations")
    if loan != 0 and capital_lease != 0:
        if loan > capital_lease:
            loan = loan - capital_lease
    if loan == 0:
        loan = current_loan + non_current_loan

    debt_equity_ratio = None
    if total_equity and total_equity != 0 and total_assets and total_assets != 0:
         total_liabilities = total_assets - total_equity
         debt_equity_ratio = (total_liabilities / total_equity)

    loan_equity_ratio = None
    if total_equity and total_equity != 0 and loan is not None:
         loan_equity_ratio = (loan / total_equity)

    fy_date = None
    if info.get('lastFiscalYearEnd'):
        try:
            fy_date = datetime.fromtimestamp(info['lastFiscalYearEnd'])
        except:
            pass

    officers = info.get('companyOfficers', [])
    top_exec = "N/A"
    if officers:
        for officer in officers:
            title = officer.get('title', '').lower()
            if 'ceo' in title or 'chairman' in title:
                top_exec = f"{officer.get('name')} ({officer.get('title')})"
                break
        if top_exec == "N/A" and len(officers) > 0:
            top_exec = officers[0].get('name')

    address_parts = [
        info.get('address1'), info.get('address2'),
        info.get('city'), info.get('country'), info.get('zip')
    ]
    full_address = ", ".join([str(a) for a in address_parts if a])

    sector = info.get('sector')
    industry = info.get('industry')
    
    # 通貨情報の取得と整形
    raw_currency = info.get('financialCurrency')
    if not raw_currency:
        raw_currency = info.get('currency', 'SGD') 
        
    display_currency = raw_currency
    if display_currency == 'CNY':
        display_currency = 'RMB (CNY)'
    
    # 為替レートの取得 (前日終値)
    exchange_rate = get_exchange_rate(display_currency)

    website = info.get('website', '')
    
    market = info.get('exchange', 'Unknown')
    if str(code).endswith('.KL'):
        ticker_clean = str(code).replace('.KL', '')
        if len(ticker_clean) == 5 and ticker_clean.startswith('03'):
            market = "LEAP"
        elif len(ticker_clean) == 4 and ticker_clean.startswith('0'):
            market = "ACE"
        elif len(ticker_clean) == 4 and ticker_clean[0] in '123456789':
            market = "Main"
        else:
            market = "Main/Other"
    
    # ★変更: 株価は「前日終値 (previousClose)」を使用する
    current_price = info.get('previousClose')
    # 取れない場合は現在値で代用
    if current_price is None:
        current_price = info.get('regularMarketPrice')

    shares_outstanding = info.get('sharesOutstanding')
    
    # 時価総額の計算 (前日終値 x 発行済株式数)
    market_cap = None
    if current_price and shares_outstanding:
        market_cap = current_price * shares_outstanding
    else:
        # 計算不能ならYahooの値をそのまま使う
        market_cap = info.get('marketCap')

    stock_price_col_name = "Stock Price"

    result = {
        "Name of Company": info.get('longName'),
        "Code": code,
        "Currency": display_currency,
        "Exchange Rate": exchange_rate,
        "Website": website,
        "Major Shareholders": shareholder_text,
        "FY": fy_date,
        "REVENUE": revenue,
        "Segments": "",
        "PROFIT": profit,
        "GROSS PROFIT": gross_profit,
        "OPERATING PROFIT": operating_income,
        "NET PROFIT (Group)": net_profit_group,
        "NET PROFIT (Shareholders)": net_profit_owners,
        "Minority Interest": minority_interest,
        "Shareholders' Equity": stockholders_equity,
        "Total Equity": total_equity,
        "TOTAL ASSET": total_assets,
        "Debt/Equity(%)": debt_equity_ratio,
        "Loan": loan,
        "Loan/Equity (%)": loan_equity_ratio,
        
        stock_price_col_name: current_price,
        "Shares Outstanding": shares_outstanding,
        "Market Cap": market_cap,
        
        "Summary of Business": info.get('longBusinessSummary', ''),
        "Chairman / CEO": top_exec,
        "Address": full_address,
        "Contact No.": info.get('phone'),
        "Number of Employee": info.get('fullTimeEmployees'),
        "Category Classification/YahooFin": sector if sector else "Not Available",
        "Sector & Industry/YahooFin": industry if industry else "Not Available",
        "Market": market
    }
    return result

# --- 4. Excel出力用整形 ---
def format_for_excel(df):
    print("データを千単位('000)に変換しています...")
    money_cols = [
        'REVENUE', 'PROFIT', 'GROSS PROFIT', 'OPERATING PROFIT', 
        'NET PROFIT (Group)', 'NET PROFIT (Shareholders)',
        'Minority Interest',
        "Shareholders' Equity", 'Total Equity', 'TOTAL ASSET',
        'Loan',
        'Market Cap',
        'Shares Outstanding'
    ]
    
    divisor = 1000.0

    for col in money_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
            df[col] = df[col] / divisor

    pct_cols = ["Debt/Equity(%)", "Loan/Equity (%)"]
    for col in pct_cols:
         if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')

    rename_map = {c: f"{c} ('000)" for c in money_cols}
    df = df.rename(columns=rename_map)
    
    if "REVENUE ('000)" in df.columns:
        df = df.rename(columns={"REVENUE ('000)": "REVENUE SGD('000)"})

    print("日付を 'Month YYYY' 形式に変換しています...")
    date_cols = ['FY']
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%b %Y').fillna('')

    return df