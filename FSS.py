# Import libraries
import pandas as pd
import numpy as np
import glob
from funct import process_data

# Define the file paths
std_path = r".\VALUESearch_FSS_Standard.xlsx"
sheet_names = [
    "금감원1_수익인식",
    "금감원2_비시장성자산",
    "금감원3_특수관계자"
]

# Call the function
df = process_data(std_path, sheet_names)

# # Retrieve the list of companies that belong to KEPCO (한국전력)
# df_kepco_list = ['N400211', 'N520047', 'N820288']
#
# df['한전여부'] = df['업체코드'].apply(lambda x: 'O' if x in df_kepco_list else 'X')


# Melting process

# List of columns to melt
value_vars_total_asset = ['2020자산총계', '2021자산총계', '2022자산총계', '2023자산총계', '2024자산총계']
value_vars_total_revenue = ['2020총수익', '2021총수익', '2022총수익', '2023총수익', '2024총수익']
value_vars_goodwill = ['2020영업권', '2021영업권', '2022영업권', '2023영업권', '2024영업권']
value_vars_goodwill_dp = ['2020(영업권상각누계액)', '2021(영업권상각누계액)', '2022(영업권상각누계액)', '2023(영업권상각누계액)', '2024(영업권상각누계액)']
value_vars_goodwill_dp2 = ['2020(영업권손상차손누계액)', '2021(영업권손상차손누계액)', '2022(영업권손상차손누계액)', '2023(영업권손상차손누계액)', '2024(영업권손상차손누계액)']
value_vars_goodwill_gvmt = ['2020(영업권정부보조금)', '2021(영업권정부보조금)', '2022(영업권정부보조금)', '2023(영업권정부보조금)', '2024(영업권정부보조금)']
value_vars_illiquid_asset = ['2020매도가능금융자산', '2021매도가능금융자산', '2022매도가능금융자산', '2023매도가능금융자산', '2024매도가능금융자산']
value_vars_illiquid_asset_l1 = ['2020(매도가능금융자산평가충당부채)', '2021(매도가능금융자산평가충당부채)', '2022(매도가능금융자산평가충당부채)', '2023(매도가능금융자산평가충당부채)', '2024(매도가능금융자산평가충당부채)']
value_vars_illiquid_asset_l2 = ['2020(매도가능금융자산손상차손누계액)', '2021(매도가능금융자산손상차손누계액)', '2022(매도가능금융자산손상차손누계액)', '2023(매도가능금융자산손상차손누계액)', '2024(매도가능금융자산손상차손누계액)']
value_vars_illiquid_asset_l3 = ['2020매도가능금융자산평가손실(-)', '2021매도가능금융자산평가손실(-)', '2022매도가능금융자산평가손실(-)', '2023매도가능금융자산평가손실(-)', '2024매도가능금융자산평가손실(-)']
value_vars_illiquid_asset_gain = ['2020매도가능금융자산평가이익', '2021매도가능금융자산평가이익', '2022매도가능금융자산평가이익', '2023매도가능금융자산평가이익', '2024매도가능금융자산평가이익']
value_vars_related_party = ['2020특수관계자-수익합계', '2021특수관계자-수익합계', '2022특수관계자-수익합계', '2023특수관계자-수익합계', '2024특수관계자-수익합계']

# Melting the data for all columns at once
df_melt = pd.melt(df,
                  id_vars=['업체코드', '종목코드', '종목명', '업체명', '대분류', '중분류', '소분류', '세분류', '시장구분', '결산월', '소속기업집단', '감사법인'],
                  value_vars= value_vars_total_asset + value_vars_total_revenue + value_vars_goodwill + value_vars_goodwill_dp + value_vars_goodwill_dp2 +
                              value_vars_goodwill_gvmt + value_vars_illiquid_asset + value_vars_illiquid_asset_l1 + value_vars_illiquid_asset_l2 + value_vars_illiquid_asset_l3 + value_vars_illiquid_asset_gain +
                                value_vars_related_party,
                  var_name='연도_항목',
                  value_name='값')


# Extracting the year and the type of value into each column: 연도 | 항목
df_melt['연도'] = df_melt['연도_항목'].str[:4]
df_melt['항목'] = df_melt['연도_항목'].str[4:]

# df_melt.loc[df_melt['항목'] != '감사법인', '값'] = pd.to_numeric(df_melt.loc[df_melt['항목'] != '감사법인', '값'], errors='coerce')

# Use pivot to reshape the DataFrame
df_pivot = df_melt.pivot(index=['업체코드', '종목코드', '종목명', '업체명', '대분류', '중분류', '소분류', '세분류', '시장구분', '결산월', '소속기업집단', '감사법인', '연도'],
                         columns='항목',
                         values='값')

# Reset index to make it a DataFrame again
df_pivot = df_pivot.reset_index()

# Fill missing values if necessary
df_pivot = df_pivot.fillna('')  # or fillna('your_value') depending on your needs

# Export to excel
df_pivot.to_excel("FSS_중점심사.xlsx", index=False)