import pandas as pd

def deduplicate_excel(file_path, sheet_name, columns):
    # 读取原始数据（跳过第一行标题）
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=1)
    result = pd.DataFrame(columns=df.columns)
    value_counts = {}  # 值频次计数器: {column_value: occurrence_count}

    # 主处理循环：逐行决策收录逻辑
    for _, row in df.iterrows():
        current_values = row[columns].tolist()  # 当前行的检测列值
        print(current_values)
        # === 阶段1: 初始化首行 ===
        if result.empty:
            result = pd.concat([result, row.to_frame().T], ignore_index=True)
            # 初始化计数器
            for val in current_values:
                value_counts[val] = value_counts.get(val, 0) + 1
            continue  # 跳过后续检测流程

        # === 阶段2: 重复关系检测 ===
        has_two_duplicate = False  # 两列重复标志
        
        # 与已收录行逐行对比
        for _, existing_row in result.iterrows():
            existing_values = existing_row[columns].tolist()
            common_count = len(set(current_values) & set(existing_values))
            
            # 如果发现两列重复，则标记为重复并跳过当前行
            if common_count >= 2:
                has_two_duplicate = True
                break

        # 两列重复直接跳过
        if has_two_duplicate:
            continue

        # === 阶段3: 值频次阀值检测 ===
        # 检查所有检测列的值是否均未达上限（3次）
        if all(value_counts.get(val, 0) < 3 for val in current_values):
            result = pd.concat([result, row.to_frame().T], ignore_index=True)
            # 更新计数器
            for val in current_values:
                value_counts[val] = value_counts.get(val, 0) + 1

    return result

# 执行入口
file_path = "e:/CODE/dataAnalysis/TEST/test.xlsx"
sheet_name = "Sheet2"
columns = ["ID1", "heroID2", "heroID3"]
result_df = deduplicate_excel(file_path, sheet_name, columns)
result_df.to_excel("e:/CODE/dataAnalysis/TEST/deduplicated_result.xlsx", index=False)
