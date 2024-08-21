# Import libraries
import pandas as pd
import numpy as np
import glob
from funct import process_data

# Define the file paths
std_file = r".\KICPA\VALUESearch_Standard.xlsx"
file_paths = [
    r".\KICPA\VALUESearch1_Liability.xlsx",
    r".\KICPA\VALUESearch2_Goodwill.xlsx",
    r".\KICPA\VALUESearch2_Development.xlsx",
    r".\KICPA\VALUESearch2_Other_Intagibles.xlsx",
    r".\KICPA\VALUESearch2_Total_Asset.xlsx",
    r".\KICPA\VALUESearch4_LiquidityRatio.xlsx"
]
#
# Call the function
df = process_data(std_file, file_paths)

# Retrieve the list of companies with Debt Issuance in the recent 3 years (2021-08-13 ~ 2024-08-13)
df_debt = pd.read_excel(r".\KICPA\VALUESearch4_DebtIssuance.xls")

df_debt = df_debt['회사명'].drop_duplicates().reset_index(drop=True)
debt_list = df_debt.to_list()

# Create the 'Debt Issued' column based on whether the company is in debt_list
df['채무증권발행(최근 3년)'] = df['종목명'].apply(lambda x: 'O' if x in debt_list else 'X')

# Retrieve the list of companies that belong to KEPCO (한국전력)
df_kepco = pd.read_excel(r".\KICPA\kepco_list.xlsx")
df_kepco_list = df_kepco['상호'].to_list()

df['한전여부'] = df['업체명'].apply(lambda x: 'O' if x in df_kepco_list else 'X')

df['2024총수익'] = df['2024총수익'] * 4

# Melting process

# List of columns to melt
value_vars_audit_firm = ['2020감사법인', '2021감사법인', '2022감사법인', '2023감사법인', '2024감사법인']
value_vars_total_revenue = ['2020총수익', '2021총수익', '2022총수익', '2023총수익', '2024총수익']
value_vars_repair_provision = ['2020하자보수충당부채', '2021하자보수충당부채', '2022하자보수충당부채', '2023하자보수충당부채', '2024하자보수충당부채']
value_vars_warranty_provision = ['2020판매보증충당부채', '2021판매보증충당부채', '2022판매보증충당부채', '2023판매보증충당부채', '2024판매보증충당부채']
value_vars_construction_loss_provision = ['2020공사손실충당부채', '2021공사손실충당부채', '2022공사손실충당부채', '2023공사손실충당부채', '2024공사손실충당부채']
value_vars_guarantee_loss_provision = ['2020보증손실충당부채', '2021보증손실충당부채', '2022보증손실충당부채', '2023보증손실충당부채', '2024보증손실충당부채']
value_vars_total_asset = ['2020자산총계', '2021자산총계', '2022자산총계', '2023자산총계', '2024자산총계']
value_vars_goodwill = ['2020영업권', '2021영업권', '2022영업권', '2023영업권', '2024영업권']
value_vars_goodwill_acc_depreciation = ['2020(영업권상각누계액)', '2021(영업권상각누계액)', '2022(영업권상각누계액)', '2023(영업권상각누계액)', '2024(영업권상각누계액)']
value_vars_goodwill_acc_impairment = ['2020(영업권손상차손누계액)', '2021(영업권손상차손누계액)', '2022(영업권손상차손누계액)', '2023(영업권손상차손누계액)', '2024(영업권손상차손누계액)']
value_vars_goodwill_gov_subsidiary = ['2020(영업권정부보조금)', '2021(영업권정부보조금)', '2022(영업권정부보조금)', '2023(영업권정부보조금)', '2024(영업권정부보조금)']
value_vars_development = ['2020개발비', '2021개발비', '2022개발비', '2023개발비', '2024개발비']
value_vars_development_acc_depreciation = ['2020(개발비상각누계액)', '2021(개발비상각누계액)', '2022(개발비상각누계액)', '2023(개발비상각누계액)', '2024(개발비상각누계액)']
value_vars_development_acc_impairment = ['2020(개발비손상차손누계액)', '2021(개발비손상차손누계액)', '2022(개발비손상차손누계액)', '2023(개발비손상차손누계액)', '2024(개발비손상차손누계액)']
value_vars_development_gov_subsidiary = ['2020(개발비정부보조금)', '2021(개발비정부보조금)', '2022(개발비정부보조금)', '2023(개발비정부보조금)', '2024(개발비정부보조금)']
value_vars_other = ['2020기타무형자산', '2021기타무형자산', '2022기타무형자산', '2023기타무형자산', '2024기타무형자산']
value_vars_other_acc_depreciation = ['2020(기타무형자산상각누계액)', '2021(기타무형자산상각누계액)', '2022(기타무형자산상각누계액)', '2023(기타무형자산상각누계액)', '2024(기타무형자산상각누계액)']
value_vars_other_acc_impairment = ['2020(기타무형자산손상차손누계액)', '2021(기타무형자산손상차손누계액)', '2022(기타무형자산손상차손누계액)', '2023(기타무형자산손상차손누계액)', '2024(기타무형자산손상차손누계액)']
value_vars_other_gov_subsidiary = ['2020(기타무형자산정부보조금)', '2021(기타무형자산정부보조금)', '2022(기타무형자산정부보조금)', '2023(기타무형자산정부보조금)', '2024(기타무형자산정부보조금)']
value_vars_current_asset = ['2020유동자산(계)', '2021유동자산(계)', '2022유동자산(계)', '2023유동자산(계)', '2024유동자산(계)']
value_vars_current_liability = ['2020유동부채(계)', '2021유동부채(계)', '2022유동부채(계)', '2023유동부채(계)', '2024유동부채(계)']

# Melting the data for all columns at once
df_melt = pd.melt(df,
                  id_vars=['업체코드', '종목코드', '종목명', '업체명', '상장법인', '대분류', '중분류', '소분류', '채무증권발행(최근 3년)', '한전여부'],
                  value_vars=value_vars_audit_firm + value_vars_total_revenue + value_vars_repair_provision + value_vars_warranty_provision + value_vars_construction_loss_provision + value_vars_total_asset +
                  value_vars_guarantee_loss_provision + value_vars_goodwill + value_vars_goodwill_acc_depreciation + value_vars_goodwill_acc_impairment + value_vars_goodwill_gov_subsidiary +
                  value_vars_development + value_vars_development_acc_depreciation + value_vars_development_acc_impairment + value_vars_development_gov_subsidiary +
                  value_vars_other + value_vars_other_acc_depreciation + value_vars_other_acc_impairment + value_vars_other_gov_subsidiary +
                  value_vars_current_asset + value_vars_current_liability,
                  var_name='연도_항목',
                  value_name='값')


# Extracting the year and the type of value into each column: 연도 | 항목
df_melt['연도'] = df_melt['연도_항목'].str[:4]
df_melt['항목'] = df_melt['연도_항목'].str[4:]

df_melt.loc[df_melt['항목'] != '감사법인', '값'] = pd.to_numeric(df_melt.loc[df_melt['항목'] != '감사법인', '값'], errors='coerce')

# Use pivot to reshape the DataFrame
df_pivot = df_melt.pivot(index=['업체코드', '종목코드', '종목명', '업체명', '상장법인', '대분류', '중분류', '소분류', '채무증권발행(최근 3년)', '한전여부', '연도'],
                         columns='항목',
                         values='값')

# Reset index to make it a DataFrame again
df_pivot = df_pivot.reset_index()

# Fill missing values if necessary
df_pivot = df_pivot.fillna('')  # or fillna('your_value') depending on your needs

# Export to excel
df_pivot.to_excel("KICPA_중점심사.xlsx", index=False)