# src/data/fetch_dart.py

import os
import requests
import pandas as pd
import xml.etree.ElementTree as ET
from tqdm import tqdm
from dotenv import load_dotenv
import zipfile
from io import BytesIO
import time

load_dotenv()
API_KEY = os.getenv("DART_API_KEY")

# 기업 고유번호 가져오기 (처음 1번만 실행해도 충분)
def get_corp_codes():
    url = f"https://opendart.fss.or.kr/api/corpCode.xml?crtfc_key={API_KEY}"
    response = requests.get(url)

    # response.content → 메모리 상에서 zip으로 처리
    with zipfile.ZipFile(BytesIO(response.content)) as zip_file:
        zip_file.extractall("data/external/")  # 여기서 CORPCODE.xml 추출됨

    # 이제 CORPCODE.xml을 파싱
    tree = ET.parse("data/external/CORPCODE.xml")
    root = tree.getroot()
    data = []

    for child in root.findall("list"):
        data.append({
            "corp_code": child.find("corp_code").text,
            "corp_name": child.find("corp_name").text,
            "stock_code": child.find("stock_code").text
        })

    df = pd.DataFrame(data)
    df.to_csv("data/external/corp_codes.csv", index=False, encoding="utf-8-sig")
    print("corp_codes.csv 저장 완료")



# 특정 회사 재무제표 가져오기 (연도별)
def fetch_financial_statement(corp_code, corp_name, year, report_code="11011"):
    url = f"https://opendart.fss.or.kr/api/fnlttSinglAcnt.json"
    params = {
        "crtfc_key": API_KEY,
        "corp_code": corp_code,
        "bsns_year": str(year),
        "reprt_code": report_code,  # 11011: 사업보고서 (1년치)
        "fs_div": "CFS"  # CFS: 연결재무제표 / OFS: 개별재무제표
    }

    r = requests.get(url, params=params)
    result = r.json()

    if result['status'] != '013':  # 013: 해당 회사는 재무제표 없음
        df = pd.DataFrame(result['list'])
        return df
    else:
        return None


# 여러 회사 반복 저장
def fetch_and_save_many(years):
    corp_df = pd.read_csv("data/external/corp_codes.csv", dtype=str)
    # 필터링: 종목코드가 존재하는 기업만
    corp_df = corp_df[corp_df['stock_code'].notna() & (corp_df['stock_code'] != ' ')] 
    corp_df = corp_df.drop_duplicates(subset="corp_code")
    print("The length is:" + str(len(corp_df)))

    for _, row in tqdm(corp_df.iterrows(), total=corp_df.shape[0]):
        corp_code = row["corp_code"]
        corp_name = row["corp_name"]
        stock_code = row["stock_code"]

        for year in years:
            try:
            # 여기에 재무제표 수집 함수 호출
                df = fetch_financial_statement(corp_code, corp_name, year)
                if df is not None:
                    folder = "data/raw/dart"
                    os.makedirs(folder, exist_ok=True)

                    # Clean corp_name to avoid invalid filename characters
                    safe_corp_name = "".join(c for c in corp_name if c.isalnum() or c in (" ", "_")).strip()

                    filename = f"{folder}/{safe_corp_name}_{year}.csv"
                    df.to_csv(filename, index=False, encoding="utf-8-sig")

            except Exception as e:
                print(f"{corp_name} ({corp_code}) {year} 수집 실패: {e}")
            time.sleep(0.2)  # API 호출 간격 조정


if __name__ == "__main__":
    # Step 1: 고유번호 저장 (처음 1회만 실행)
    if not os.path.exists("data/external/corp_codes.csv"):
        get_corp_codes()

    # Step 2: 원하는 종목코드와 연도 설정
    years = [2020, 2021, 2022, 2023, 2024]

    # Step 3: 재무제표 수집 및 저장
    fetch_and_save_many(years)
