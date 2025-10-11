import os
import pandas as pd
from glob import glob

TARGET_ACCOUNTS = [
    "매출액", "영업이익", "당기순이익",  # 손익계산서
    "자산총계", "부채총계", "자본총계"   # 재무상태표
]

def load_and_clean_file(filepath):
    try:
        df = pd.read_csv(filepath)

        # 파일명에서 회사명 추출 (예: '삼성전자_2023.csv' → '삼성전자')
        filename = os.path.basename(filepath)
        corp_name = filename.split('_')[0]

        df = df[df['account_nm'].isin(TARGET_ACCOUNTS)]
        df = df[['bsns_year', 'stock_code', 'account_nm', 'thstrm_amount']]
        df['stock_code'] = df['stock_code'].astype(str).str.zfill(6)
        df = df.rename(columns={
            'bsns_year': 'year',
            'account_nm': 'account',
            'thstrm_amount': 'value'
        })
        df['corp_name'] = corp_name  # 회사명 추가
        df['value'] = df['value'].astype(str).str.replace(',', '').str.replace(' ', '')
        df['value'] = pd.to_numeric(df['value'], errors='coerce')
        return df
    except Exception as e:
        print(f"Failed to process {filepath}: {e}")
        return pd.DataFrame()

def preprocess_all_financials(input_folder="data/raw/dart", output_file="data/processed/financial_summary.csv"):
    all_files = glob(os.path.join(input_folder, "*.csv"))
    all_data = []

    for file in all_files:
        df = load_and_clean_file(file)
        if not df.empty:
            all_data.append(df)

    merged = pd.concat(all_data)
    final = merged.pivot_table(index=['corp_name', 'year', 'stock_code'], columns='account', values='value').reset_index()
    final.to_csv(output_file, index=False, encoding='utf-8-sig')
    print(f"Saved merged financial summary to {output_file}")

if __name__ == "__main__":
    preprocess_all_financials()
