# Import libraries
import pandas as pd
import numpy as np
import glob

# Read in Raw Data downloaded from FSS
df_std = pd.read_excel(r".\KICPA\VALUESearch_Standard.xlsx")

# Specify the column that needs renaming
old_column_name = '691005.업체명'
new_column_name = old_column_name.replace('691005.', '')
df_std.rename(columns={old_column_name: new_column_name}, inplace=True)

# Use REGEX to substitute substring
pattern1 = r'/[^/]+?\.'
pattern2 = r'\d+\.\w+-'
pattern3 = r'\(.*?\)'
pattern4 = r'^[A-Za-z]\d{5}\.'
df_std.columns = df_std.columns.str.replace(pattern1, '', regex=True)
df_std.columns = df_std.columns.str.replace(pattern2, '', regex=True)
df_std.columns = df_std.columns.str.replace(pattern3, '', regex=True)
df_std['대분류'] = df_std['대분류'].str.replace(pattern3, '', regex=True)
df_std['대분류'] = df_std['대분류'].str.replace(pattern4, '', regex=True)
df_std['중분류'] = df_std['중분류'].str.replace(pattern4, '', regex=True)
df_std['소분류'] = df_std['소분류'].str.replace(pattern4, '', regex=True)
df_std


# List of file paths and corresponding DataFrames
file_paths = [
    r".\KICPA\VALUESearch1_Liability.xlsx",
    r".\KICPA\VALUESearch2_Goodwill.xlsx",
    r".\KICPA\VALUESearch2_Development.xlsx",
    r".\KICPA\VALUESearch2_Other_Intagibles.xlsx",
    r".\KICPA\VALUESearch4_LiquidityRatio.xlsx"
]

# Read and process each file
dataframes = []
for path in file_paths:
    df = pd.read_excel(path)
    df.columns = df.columns.str.replace(pattern1, '', regex=True)
    dataframes.append(df)

# Merge all DataFrames with the base DataFrame df_std
df = df_std
for df_add in dataframes:
    df = pd.merge(df, df_add, on=['업체코드', '종목코드', '종목명'], how='left')

# Melting process

# List of columns to melt
value_vars_audit_firm = ['2020감사법인', '2021감사법인', '2022감사법인', '2023감사법인', '2024감사법인']
value_vars_total_revenue = ['2020총수익', '2021총수익', '2022총수익', '2023총수익', '2024총수익']
value_vars_repair_provision = ['2020하자보수충당부채', '2021하자보수충당부채', '2022하자보수충당부채', '2023하자보수충당부채', '2024하자보수충당부채']
value_vars_warranty_provision = ['2020판매보증충당부채', '2021판매보증충당부채', '2022판매보증충당부채', '2023판매보증충당부채', '2024판매보증충당부채']
value_vars_construction_loss_provision = ['2020공사손실충당부채', '2021공사손실충당부채', '2022공사손실충당부채', '2023공사손실충당부채', '2024공사손실충당부채']
value_vars_guarantee_loss_provision = ['2020보증손실충당부채', '2021보증손실충당부채', '2022보증손실충당부채', '2023보증손실충당부채', '2024보증손실충당부채']
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
                  id_vars=['업체코드', '종목코드', '종목명', '업체명', '상장법인', '대분류', '중분류', '소분류'],
                  value_vars=value_vars_audit_firm + value_vars_total_revenue + value_vars_repair_provision + value_vars_warranty_provision + value_vars_construction_loss_provision + 
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
df_pivot = df_melt.pivot(index=['업체코드', '종목코드', '종목명', '업체명', '상장법인', '대분류', '중분류', '소분류', '연도'], 
                         columns='항목', 
                         values='값')

# Reset index to make it a DataFrame again
df_pivot = df_pivot.reset_index()

# Fill missing values if necessary
df_pivot = df_pivot.fillna('')  # or fillna('your_value') depending on your needs


# Export to excel
df_pivot.to_excel("KICPA_중점심사.xlsx", index=False)

