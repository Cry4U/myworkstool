# -*- coding: utf-8 -*-
import pandas as pd
import os
import glob
import re
from functools import reduce

def extract_number(filename):
    match = re.search(r'\d+', os.path.basename(filename))
    return int(match.group()) if match else float('inf')

folder_path = "E:\\CODE\\dataAnalysis\\TEST"
filelist = sorted(glob.glob(os.path.join(folder_path, "*.xls*")), key=extract_number)

# 读取所有表格
dfs = [pd.read_excel(file, engine='xlrd') for file in filelist]

# 取“队名”+最后两列（避免重复“队名”）
last_two_cols = []
for df in dfs:
    cols = df.columns
    select_cols = ['队名'] + [col for col in cols[-2:] if col != '队名']
    last_two_cols.append(df.loc[:, select_cols])

# 依次按“队名”合并
merged_df = reduce(lambda left, right: pd.merge(left, right, on='队名', how='outer'), last_two_cols)

# 按第一张表格的“队名”顺序排序
first_team_order = last_two_cols[0]['队名'].tolist()
merged_df['队名'] = pd.Categorical(merged_df['队名'], categories=first_team_order, ordered=True)
merged_df = merged_df.sort_values('队名').reset_index(drop=True)

merged_df.to_excel('merged_output.xlsx', index=False)
