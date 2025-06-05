# -*- coding: utf-8 -*-
import pandas as pd
import glob
import re
from functools import reduce

# 定义一个函数，从文件名中提取数字，用于排序
def extract_number(filename):
    match = re.search(r'\d+', filename)  # 使用正则表达式匹配文件名中的数字
    return int(match.group()) if match else float('inf')  # 如果匹配到数字，返回整数；否则返回正无穷大

# 定义文件夹路径
folder_path = "E:\\CODE\\dataAnalysis\\TEST"

# 获取文件夹中所有以 .xls结尾的文件，并按文件名中的数字排序
filelist = sorted(glob.glob(f"{folder_path}\\*.xls"), key=extract_number)

# 读取所有 Excel 文件，存储为 DataFrame 列表
dfs = [pd.read_excel(file, engine='xlrd') for file in filelist]

# 提取每个 DataFrame 中的“队名”列和最后一列
last_cols = []
for df in dfs:
    cols = df.columns  # 获取列名
    # 选择“队名”列和最后一列（如果最后一列不是“队名”）
    select_cols = ['队名'] + [cols[-1]]
    last_cols.append(df.loc[:, select_cols])  # 提取所需列并添加到列表中

# 按“队名”列进行外连接合并所有 DataFrame
# 如果列名重复，后续表格的列名会添加后缀 '_dup'
merged_df = reduce(lambda left, right: pd.merge(left, right, on='队名', how='outer', suffixes=('', '_dup')), last_cols)

# 获取第一个文件中“队名”的顺序，用于排序
first_team_order = last_cols[0]['队名'].tolist()
# 将“队名”列设置为分类类型，并按照第一个文件的顺序排序
merged_df['队名'] = pd.Categorical(merged_df['队名'], categories=first_team_order, ordered=True)
merged_df = merged_df.sort_values('队名').reset_index(drop=True)  # 按“队名”排序并重置索引

# 将合并后的结果保存为 Excel 文件
merged_df.to_excel('merged_output.xlsx', index=False)