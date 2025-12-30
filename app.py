import pandas as pd
import yfinance as yf
from datetime import datetime
import time
import json
import os
from google import genai
from dotenv import load_dotenv

load_dotenv()

# APIキーの取得（環境変数から）
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

client = None
if GEMINI_API_KEY:
    try:
        client = genai.Client(api_key=GEMINI_API_KEY)
    except Exception as e:
        print(f"⚠️ Gemini Client Initialization Error: {e}")

# --- 1. AIによるセグメント分析 (encodingエラー対策済み) ---
def batch_analyze_segments(all_results_list):
    """
    複数企業のビジネスサマリーをまとめてGeminiに送り、セグメント情報を抽出します。
    """
    if not client:
        print("⚠️ Gemini client is not initialized. Skipping AI analysis.")
        return all_results_list

    # ビジネスサマリーがあるものだけを対象にする
    targets = [item for item in all_results_list if item.get('Summary of Business')]
    
    if not targets:
        return all_results_list

    print(f"🤖 Gemini AI分析開始: 対象 {len(targets)} 件")
    
    batch_size = 20
    model_name = 'gemini-2.0-flash' # 高速・高機能な最新モデルを使用
    
    for i in range(0, len(targets), batch_size):
        batch = targets[i : i + batch_size]
        
        # プロンプトの構築
        input_text = ""
        for item in batch:
            # 文字列を安全に処理（None対策と文字数制限）
            summary = str(item.get('Summary of Business', ''))[:500].replace("\n", " ")
            input_text += f"Code: {item['Code']}\nSummary: {summary}\n---\n"

        prompt = f"""
        You are a financial analyst. Based on the business summaries provided below, 
        extract the main 'Business Segments' for EACH company.

        # Output Rules
        - Return ONLY a valid JSON object.
        - Keys must be the 'Code'.
        - Values must be a concise string of segments (comma separated).
        - If unknown, provide a 3-4 word summary of their industry.

        # Input Data
        {input_text}
        """

        try:
            # AIリクエスト
            response = client.models.generate_content(
                model=model_name,
                contents=prompt
            )
            
            response_text = response.text.strip()
            
            # JSON部分の抽出
            if "```json" in response_text:
                response_text = response_text.split("```json")[1].split("```")[0].strip()
            elif "```" in response_text:
                response_text = response_text.split("```")[1].split("```")[0].strip()
            
            segments_map = json.loads(response_text)

            # 結果を元のリストに反映
            for item in batch:
                code = item['Code']
                if code in segments_map:
                    item['Segments'] = segments_map[code]
            
            time.sleep(1) # レートリミット回避

        except Exception as e:
            # ログ出力時のエンコードエラーを防ぐため、エラーメッセージを安全に表示
            err_msg = str(e).encode('ascii', 'ignore').decode('ascii')
            print(f"⚠️ Batch AI Error: {err_msg}")
            continue

    return all_results_list


# --- 2. データ取得関数 (404エラー対策済み) ---
def get_stock_data(code):
    """
    Yahoo Financeから個別の銘柄データを取得します。
    """
    try:
        ticker = yf.Ticker(code)
        # 取得を試みる（存在しない場合はここでエラーまたは空が返る）
        info = ticker.info
        
        if not info or 'symbol' not in info:
            print(f"Skipped {code}: No data found")
            return None

        return {
            "info": info,
            "balance_sheet": ticker.balance_sheet,
            "financials": ticker.financials,
            "major_holders": ticker.major_holders,
            "institutional_holders": ticker.institutional_holders
        }
    except Exception as e:
        # ログを簡潔に（404エラーなど）
        print(f"Skipped {code}: {str(e)[:50]}")
        return None

def get_exchange_rate(from_currency):
    """
    指定通貨からSGDへの為替レートを取得。
    """
    if not from_currency or from_currency == "SGD":
        return 1.0
    
    currency_code = "CNY" if from_currency == "RMB (CNY)" else from_currency
    pair = f"{currency_code}SGD=X"
    
    try:
        ticker = yf.Ticker(pair)
        rate = ticker.info.get('previousClose')
        if rate is None:
            hist = ticker.history(period="5d")
            rate = hist['Close'].iloc[-1] if not hist.empty else "N/A"
        return rate
    except:
        return "N/A"


# --- 3. データの抽出・整形ロジック ---
def format_shareholders(holders_data, data_type="major"):
    if holders_data is None or holders_data.empty:
        return None
    
    lines = []
    try:
        # 基本的に上位5件程度を抽出して文字列にする
        for index, row in holders_data.head(5).iterrows():
            if len(row) >= 2:
                lines.append(f"{row.iloc[1]}: {row.iloc[0]}")
            else:
                lines.append(f"{index}: {row.iloc[0]}")
    except:
        return "Error parsing"
    
    return "\n".join(lines) if lines else "Not Available"


def extract_data(code, raw_data):
    info = raw_data.get("info", {})
    bs = raw_data.get("balance_sheet")
    inc = raw_data.get("financials")
    
    # 決算期日の特定
    latest_date = bs.columns[0] if bs is not None and not bs.empty else None

    def get_val(df, key):
        if df is not None and not df.empty and key in df.index and latest_date in df.columns:
            val = df.loc[key, latest_date]
            return val if pd.notnull(val) else 0
        return 0

    # 会計数値の抽出
    revenue = get_val(inc, "Total Revenue")
    operating_income = get_val(inc, "Operating Income")
    net_profit_owners = get_val(inc, "Net Income")
    
    total_assets = get_val(bs, "Total Assets")
    total_equity = get_val(bs, "Total Equity Gross Minority Interest")
    if total_equity == 0:
        total_equity = get_val(bs, "Stockholders Equity")

    loan = get_val(bs, "Total Debt")
    
    # 指標計算
    debt_ratio = (total_assets - total_equity) / total_equity if total_equity != 0 else 0
    loan_ratio = loan / total_equity if total_equity != 0 else 0

    # 基本情報
    raw_currency = info.get('financialCurrency', info.get('currency', 'SGD'))
    display_currency = 'RMB (CNY)' if raw_currency == 'CNY' else raw_currency
    
    # マレーシア市場の細分類
    market = info.get('exchange', 'Unknown')
    if str(code).endswith('.KL'):
        clean_code = str(code).replace('.KL', '')
        if clean_code.startswith('03'): market = "LEAP"
        elif clean_code.startswith('0'): market = "ACE"
        else: market = "Main"

    return {
        "Name of Company": info.get('longName'),
        "Code": code,
        "Currency": display_currency,
        "Exchange Rate": get_exchange_rate(display_currency),
        "Website": info.get('website'),
        "Major Shareholders": format_shareholders(raw_data.get("major_holders")),
        "FY": datetime.fromtimestamp(info['lastFiscalYearEnd']) if info.get('lastFiscalYearEnd') else None,
        "REVENUE": revenue,
        "Segments": "",
        "PROFIT": get_val(inc, "Pretax Income") or operating_income,
        "GROSS PROFIT": get_val(inc, "Gross Profit"),
        "OPERATING PROFIT": operating_income,
        "NET PROFIT (Group)": get_val(inc, "Net Income Including Noncontrolling Interests") or net_profit_owners,
        "NET PROFIT (Shareholders)": net_profit_owners,
        "Minority Interest": get_val(bs, "Minority Interest"),
        "Shareholders' Equity": get_val(bs, "Stockholders Equity"),
        "Total Equity": total_equity,
        "TOTAL ASSET": total_assets,
        "Debt/Equity(%)": debt_ratio,
        "Loan": loan,
        "Loan/Equity (%)": loan_ratio,
        "Stock Price": info.get('previousClose') or info.get('regularMarketPrice'),
        "Shares Outstanding": info.get('sharesOutstanding'),
        "Market Cap": info.get('marketCap'),
        "Summary of Business": info.get('longBusinessSummary', ''),
        "Chairman / CEO": info.get('companyOfficers', [{}])[0].get('name', 'N/A'),
        "Address": f"{info.get('address1', '')}, {info.get('city', '')}, {info.get('country', '')}",
        "Contact No.": info.get('phone'),
        "Number of Employee": info.get('fullTimeEmployees'),
        "Category Classification/YahooFin": info.get('sector', 'N/A'),
        "Sector & Industry/YahooFin": info.get('industry', 'N/A'),
        "Market": market
    }

# --- 4. Excel出力用数値整形 ---
def format_for_excel(df):
    """
    数値を1000単位にし、日付を見やすく整形します。
    """
    money_cols = [
        'REVENUE', 'PROFIT', 'GROSS PROFIT', 'OPERATING PROFIT', 
        'NET PROFIT (Group)', 'NET PROFIT (Shareholders)', 'Minority Interest',
        "Shareholders' Equity", 'Total Equity', 'TOTAL ASSET', 'Loan',
        'Market Cap', 'Shares Outstanding'
    ]
    
    for col in money_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce') / 1000.0

    # カラム名に ('000) を付与
    df = df.rename(columns={c: f"{c} ('000)" for c in money_cols})
    if "REVENUE ('000)" in df.columns:
        df = df.rename(columns={"REVENUE ('000)": "REVENUE SGD('000)"})

    # 日付整形
    if 'FY' in df.columns:
        df['FY'] = pd.to_datetime(df['FY'], errors='coerce').dt.strftime('%b %Y').fillna('')

    return df
