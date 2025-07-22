import pandas as pd

def deduplicate_excel(file_path, sheet_name, columns, max_two_duplicate_rows=1, max_value_frequency=3):
    # 读取原始数据（跳过第一行标题）
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=1)
    result = pd.DataFrame(columns=df.columns)
    value_counts = {}  # 值频次计数器: {column_value: occurrence_count}
    duplicate_counts = {}  # 两列重复计数器: {frozenset_of_common_values: occurrence_count}

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
        duplicate_key = None  # 重复键
        
        # 与已收录行逐行对比
        for _, existing_row in result.iterrows():
            existing_values = existing_row[columns].tolist()
            common_values = set(current_values) & set(existing_values)
            common_count = len(common_values)
            
            # 如果发现两列重复
            if common_count >= 2:
                # 使用交集作为重复键，确保相同的重复关系使用同一个计数器
                duplicate_key = frozenset(common_values)
                has_two_duplicate = True
                break

        # 处理两列重复的情况
        if has_two_duplicate:
            # 初始化计数器
            if duplicate_key not in duplicate_counts:
                duplicate_counts[duplicate_key] = 0
            
            # 检查是否可以保留该重复行
            if duplicate_counts[duplicate_key] < max_two_duplicate_rows:
                duplicate_counts[duplicate_key] += 1
                # 对于两列重复的行，继续执行阶段3检测
                if all(value_counts.get(val, 0) < max_value_frequency for val in current_values):
                    result = pd.concat([result, row.to_frame().T], ignore_index=True)
                    # 更新值频次计数器
                    for val in current_values:
                        value_counts[val] = value_counts.get(val, 0) + 1
            # 如果达到重复行数上限，直接跳过
            continue

        # === 阶段3: 值频次阈值检测 ===
        # 只有非两列重复的行才执行此阶段
        # 检查所有检测列的值是否均未达上限
        if all(value_counts.get(val, 0) < max_value_frequency for val in current_values):
            result = pd.concat([result, row.to_frame().T], ignore_index=True)
            # 更新计数器
            for val in current_values:
                value_counts[val] = value_counts.get(val, 0) + 1

    return result

# 执行入口
file_path = "e:/CODE/dataAnalysis/TEST/test.xlsx"
sheet_name = "Sheet2"
columns = ["ID1", "heroID2", "heroID3"]
max_two_duplicate_rows = 3  # 可设置保留两列重复的最大数量
max_value_frequency = 1    # 可设置值频次的最大阈值
result_df = deduplicate_excel(file_path, sheet_name, columns, max_two_duplicate_rows, max_value_frequency)
result_df.to_excel("e:/CODE/dataAnalysis/TEST/deduplicated_result.xlsx", index=False)