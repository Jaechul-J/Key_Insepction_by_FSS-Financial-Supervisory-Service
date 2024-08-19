# Import libraries
import pandas as pd
import numpy as np
import glob
from funct import process_data

# Define the file paths
std_file = r".\FSS\VALUESearch_Standard.xlsx"
file_paths = [
    r".\FSS\VALUESearch1_Revenue_Recognition.xlsx",
    r".\FSS\VALUESearch2_Non-Marketable_Asset.xlsx",
    r".\FSS\VALUESearch3_Related_Party_trans.xlsx"
]

# Call the function
df = process_data(std_file, file_paths)

# Retrieve the list of companies that belong to KEPCO (한국전력)
df_kepco_list = ['N400211', 'N520047', 'N820288']

df['한전여부'] = df['업체코드'].apply(lambda x: 'O' if x in df_kepco_list else 'X')


# Melting process

# List of columns to melt
value_vars_audit_firm = ['2020감사법인', '2021감사법인', '2022감사법인', '2023감사법인', '2024감사법인']
value_vars_total_revenue = ['2020총수익', '2021총수익', '2022총수익', '2023총수익', '2024총수익']
value_vars_total_asset = ['2020자산총계', '2021자산총계', '2022자산총계', '2023자산총계', '2024자산총계']
value_vars_illiquid_asset = ['2020매도가능금융자산', '2021매도가능금융자산', '2022매도가능금융자산', '2023매도가능금융자산', '2024매도가능금융자산']
value_vars_illiquid_asset2 = ['2020매도가능금융자산2', '2021매도가능금융자산2', '2022매도가능금융자산2', '2023매도가능금융자산2', '2024매도가능금융자산2']
value_vars_related_party = ['2020특수관계자-수익합계', '2021특수관계자-수익합계', '2022특수관계자-수익합계', '2023특수관계자-수익합계', '2024특수관계자-수익합계']

# Melting the data for all columns at once
df_melt = pd.melt(df,
                  id_vars=['업체코드', '종목코드', '종목명', '업체명', '대분류', '중분류', '소분류', '한전여부'],
                  value_vars=value_vars_audit_firm + value_vars_total_revenue + value_vars_total_asset + value_vars_illiquid_asset + value_vars_illiquid_asset2 +
                  value_vars_related_party,
                  var_name='연도_항목',
                  value_name='값')


# Extracting the year and the type of value into each column: 연도 | 항목
df_melt['연도'] = df_melt['연도_항목'].str[:4]
df_melt['항목'] = df_melt['연도_항목'].str[4:]

df_melt.loc[df_melt['항목'] != '감사법인', '값'] = pd.to_numeric(df_melt.loc[df_melt['항목'] != '감사법인', '값'], errors='coerce')

# Use pivot to reshape the DataFrame
df_pivot = df_melt.pivot(index=['업체코드', '종목코드', '종목명', '업체명', '대분류', '중분류', '소분류', '한전여부', '연도'],
                         columns='항목',
                         values='값')

# Reset index to make it a DataFrame again
df_pivot = df_pivot.reset_index()

# Fill missing values if necessary
df_pivot = df_pivot.fillna('')  # or fillna('your_value') depending on your needs

# Export to excel
df_pivot.to_excel("FSS_중점심사.xlsx", index=False)