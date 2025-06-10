import pandas as pd

def deduplicate_excel(file_path, sheet_name, columns):
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=1)
    result = pd.DataFrame(columns=df.columns)  # 初始化结果 DataFrame
    value_counts = {}  # 用于记录每个值的出现次数

    for _, row in df.iterrows():
        values = row[columns].tolist()

        # 如果 result 为空，直接添加当前行
        if result.empty:
            result = pd.concat([result, row.to_frame().T], ignore_index=True)
            for val in values:
                value_counts[val] = value_counts.get(val, 0) + 1
            continue

        # 检查重复关系
        duplicate_rows = []
        for _, existing_row in result.iterrows():
            existing_values = existing_row[columns].tolist()
            common_values = set(values) & set(existing_values)
            if len(common_values) == 3:
                duplicate_rows.append(existing_row)
                break  # 如果存在 3 个值一样的行，跳过当前行
            elif len(common_values) == 2:
                duplicate_rows.append(existing_row)

        if len(duplicate_rows) >= 1 and len(set(values) & set(duplicate_rows[0][columns].tolist())) == 3:
            continue  # 如果存在 3 个值一样的行，跳过当前行
        elif len(duplicate_rows) >= 3:
            continue  # 如果存在 2 个值一样的行且已保留 3 行，跳过当前行

        # 检查值出现次数
        if all(value_counts.get(val, 0) < 3 for val in values):
            result = pd.concat([result, row.to_frame().T], ignore_index=True)
            for val in values:
                value_counts[val] = value_counts.get(val, 0) + 1

    return result

file_path = "e:/CODE/dataAnalysis/TEST/test.xlsx"
sheet_name = "Sheet1"
columns = ["ID1", "ID2", "ID3"]

result_df = deduplicate_excel(file_path, sheet_name, columns)
result_df.to_excel("e:/CODE/dataAnalysis/TEST/deduplicated_result.xlsx", index=False)
