import pandas as pd
import yfinance as yf
from datetime import datetime
import time
import json
import os
import sys
from google import genai
from dotenv import load_dotenv

# 文字エンコードの問題を回避するための設定
load_dotenv()

# APIキーの取得
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

client = None
if GEMINI_API_KEY:
    try:
        client = genai.Client(api_key=GEMINI_API_KEY)
    except Exception as e:
        print(f"⚠️ Gemini Client Init Error: {e}")

# --- 1. AIによるセグメント分析 (エンコードエラー対策を最大化) ---
def batch_analyze_segments(all_results_list):
    if not client:
        return all_results_list

    targets = [item for item in all_results_list if item.get('Summary of Business')]
    if not targets:
        return all_results_list

    print(f"🤖 Gemini AI分析開始: 対象 {len(targets)} 件")
    
    batch_size = 15 # 負荷を減らすため少し小さく
    model_name = 'gemini-2.0-flash'
    
    for i in range(0, len(targets), batch_size):
        batch = targets[i : i + batch_size]
        
        input_text = ""
        for item in batch:
            # ★エンコード対策：特殊文字を除去し、ASCIIで表現可能な形式に一旦落としてから戻す、または確実にUTF-8で扱う
            summary = str(item.get('Summary of Business', ''))[:500]
            # 改行やタブを排除して1行にする
            summary = " ".join(summary.split())
            input_text += f"Code: {item['Code']}\nSummary: {summary}\n---\n"

        prompt = f"""
        Extract the main 'Business Segments' for EACH company based on the summary.
        Return ONLY a JSON object: {{"CODE": "Segments", ...}}
        
        # Input Data
        {input_text}
        """

        try:
            # AIへのリクエスト
            response = client.models.generate_content(
                model=model_name,
                contents=prompt
            )
            
            response_text = response.text.strip()
            
            # JSONの抽出
            if "```json" in response_text:
                response_text = response_text.split("```json")[1].split("```")[0].strip()
            elif "```" in response_text:
                response_text = response_text.split("```")[1].split("```")[0].strip()
            
            # 先頭や末尾にゴミがあれば除去
            start_idx = response_text.find('{')
            end_idx = response_text.rfind('}') + 1
            if start_idx != -1 and end_idx != 0:
                response_text = response_text[start_idx:end_idx]

            segments_map = json.loads(response_text)

            for item in batch:
                code = item['Code']
                if code in segments_map:
                    item['Segments'] = str(segments_map[code])
            
            time.sleep(1.5)

        except Exception as e:
            # エラーメッセージ自体のエンコードエラーも防ぐ
            print(f"⚠️ Batch AI Error for chunk {i}: AI Analysis Skipped")
            continue

    return all_results_list

# --- 2. データ取得関数 (404対策と安定化) ---
def get_stock_data(code):
    try:
        # yfinanceのセッションを安定させるための工夫
        ticker = yf.Ticker(code)
        
        # まずは基本的なinfoが取れるか確認
        info = ticker.info
        
        # クラウド環境では info が空になる場合があるため、historyで補完を試みる
        if not info or len(info) < 5:
            hist = ticker.history(period="1d")
            if hist.empty:
                print(f"Skipped {code}: No data available on Yahoo Finance")
                return None
            # 最小限の情報を偽装してエラーを防ぐ
            info = info if info else {}
            info['symbol'] = code
            info['shortName'] = info.get('shortName', code)

        return {
            "info": info,
            "balance_sheet": ticker.balance_sheet,
            "financials": ticker.financials,
            "major_holders": ticker.major_holders,
            "institutional_holders": ticker.institutional_holders
        }
    except Exception as e:
        print(f"Skipped {code}: {str(e)[:50]}")
        return None

def get_exchange_rate(from_currency):
    if not from_currency or from_currency == "SGD":
        return 1.0
    
    currency_code = "CNY" if from_currency == "RMB (CNY)" else from_currency
    pair = f"{currency_code}SGD=X"
    
    try:
        ticker = yf.Ticker(pair)
        # infoから取れない場合はhistoryを使う
        hist = ticker.history(period="5d")
        if not hist.empty:
            return hist['Close'].iloc[-1]
        return ticker.info.get('previousClose', "N/A")
    except:
        return "N/A"

# --- 3. データの抽出・整形 ---
def format_shareholders(holders_data):
    if holders_data is None or holders_data.empty:
        return "Not Available"
    
    try:
        # 表形式を文字列に変換
        lines = []
        for _, row in holders_data.head(5).iterrows():
            name = str(row.iloc[1]) if len(row) > 1 else str(row.name)
            value = str(row.iloc[0])
            lines.append(f"{name}: {value}")
        return "\n".join(lines)
    except:
        return "Data Parsing Error"

def extract_data(code, raw_data):
    info = raw_data.get("info", {})
    bs = raw_data.get("balance_sheet")
    inc = raw_data.get("financials")
    
    # 日付の特定（最も新しいカラム）
    latest_date = None
    if bs is not None and not bs.empty: latest_date = bs.columns[0]
    elif inc is not None and not inc.empty: latest_date = inc.columns[0]

    def get_val(df, key):
        if df is not None and not df.empty and key in df.index and latest_date in df.columns:
            val = df.loc[key, latest_date]
            return val if pd.notnull(val) else 0
        return 0

    revenue = get_val(inc, "Total Revenue")
    op_income = get_val(inc, "Operating Income")
    net_profit = get_val(inc, "Net Income")
    
    total_assets = get_val(bs, "Total Assets")
    total_equity = get_val(bs, "Total Equity Gross Minority Interest") or get_val(bs, "Stockholders Equity")
    loan = get_val(bs, "Total Debt")

    raw_currency = info.get('financialCurrency', info.get('currency', 'SGD'))
    display_currency = 'RMB (CNY)' if raw_currency == 'CNY' else raw_currency

    return {
        "Name of Company": info.get('longName') or info.get('shortName') or code,
        "Code": code,
        "Currency": display_currency,
        "Exchange Rate": get_exchange_rate(display_currency),
        "Website": info.get('website', ''),
        "Major Shareholders": format_shareholders(raw_data.get("major_holders")),
        "FY": datetime.fromtimestamp(info['lastFiscalYearEnd']) if info.get('lastFiscalYearEnd') else None,
        "REVENUE": revenue,
        "Segments": "",
        "PROFIT": get_val(inc, "Pretax Income") or op_income,
        "GROSS PROFIT": get_val(inc, "Gross Profit"),
        "OPERATING PROFIT": op_income,
        "NET PROFIT (Group)": get_val(inc, "Net Income Including Noncontrolling Interests") or net_profit,
        "NET PROFIT (Shareholders)": net_profit,
        "Minority Interest": get_val(bs, "Minority Interest"),
        "Shareholders' Equity": get_val(bs, "Stockholders Equity"),
        "Total Equity": total_equity,
        "TOTAL ASSET": total_assets,
        "Debt/Equity(%)": (total_assets - total_equity) / total_equity if total_equity else 0,
        "Loan": loan,
        "Loan/Equity (%)": loan / total_equity if total_equity else 0,
        "Stock Price": info.get('previousClose') or info.get('regularMarketPrice') or 0,
        "Shares Outstanding": info.get('sharesOutstanding'),
        "Market Cap": info.get('marketCap'),
        "Summary of Business": info.get('longBusinessSummary', ''),
        "Chairman / CEO": info.get('companyOfficers', [{}])[0].get('name', 'N/A') if info.get('companyOfficers') else 'N/A',
        "Address": f"{info.get('address1', '')}, {info.get('city', '')}",
        "Contact No.": info.get('phone', ''),
        "Number of Employee": info.get('fullTimeEmployees'),
        "Category Classification/YahooFin": info.get('sector', 'N/A'),
        "Sector & Industry/YahooFin": info.get('industry', 'N/A'),
        "Market": info.get('exchange', 'Unknown')
    }

def format_for_excel(df):
    money_cols = [
        'REVENUE', 'PROFIT', 'GROSS PROFIT', 'OPERATING PROFIT', 
        'NET PROFIT (Group)', 'NET PROFIT (Shareholders)', 'Minority Interest',
        "Shareholders' Equity", 'Total Equity', 'TOTAL ASSET', 'Loan',
        'Market Cap', 'Shares Outstanding'
    ]
    for col in money_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0) / 1000.0

    df = df.rename(columns={c: f"{c} ('000)" for c in money_cols})
    if "REVENUE ('000)" in df.columns:
        df = df.rename(columns={"REVENUE ('000)": "REVENUE SGD('000)"})

    if 'FY' in df.columns:
        df['FY'] = pd.to_datetime(df['FY'], errors='coerce').dt.strftime('%b %Y').fillna('')
    return df
