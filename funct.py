# Import libraries
import pandas as pd
import numpy as np
import glob

# Function Declaration to tidy standard data (Multi use case)
def process_data(std_file, file_paths):
    """
    Process the data up to the merging step.

    Parameters:
    std_file (str): Path to the standard file.
    file_paths (list): List of paths to additional files to be merged.

    Returns:
    pd.DataFrame: Merged DataFrame.
    """
    # Step 1: Read in Raw Data downloaded from FSS
    df_std = pd.read_excel(std_file)

    # Step 2: Specify the column that needs renaming
    old_column_name = '691005.업체명'
    new_column_name = old_column_name.replace('691005.', '')
    df_std.rename(columns={old_column_name: new_column_name}, inplace=True)

    # Step 3: Use REGEX to substitute substrings in column names and data
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

    # Step 4: Read and process each additional file
    dataframes = []
    for path in file_paths:
        df = pd.read_excel(path)
        df.columns = df.columns.str.replace(pattern1, '', regex=True)
        dataframes.append(df)

    # Step 5: Merge all DataFrames with the base DataFrame df_std
    df = df_std
    for df_add in dataframes:
        df = pd.merge(df, df_add, on=['업체코드', '종목코드', '종목명'], how='left')

    return df