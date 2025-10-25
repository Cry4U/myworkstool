import pandas as pd
from datetime import datetime
import time

def deduplicate_excel_optimized(file_path, sheet_name, columns, sum_field='rate', max_two_duplicate_rows=2, max_value_frequency=3):
    """
    Excel数据去重函数：
    1. ID1和(heroID2、heroID3)值完全相同的行不计入(位置可互换)
    2. ID1和(heroID2、heroID3)值有两个相同的组合不再计入(位置可互换)
    3. 单个值在三列中出现超过限制次数后不再计入
    """
    print(f"开始处理数据：{datetime.now().strftime('%H:%M:%S')}")
    start_time = time.time()
    
    # 读取数据
    print("读取Excel文件...")
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=1)
    original_count = len(df)
    
    # 预计算字段累加值
    print(f"计算{sum_field}累加值...")
    rate_sums = {}
    df['ids_set'] = df[columns].apply(lambda x: frozenset(x), axis=1)
    for _, group in df.groupby('ids_set'):
        rate_sums[group.iloc[0]['ids_set']] = group[sum_field].sum()
    
    # 准备数据结构
    result_rows = []
    processed_rows = []
    value_counts = {}  # 记录每个值的出现次数
    
    # 修改两值相同检测部分
    def check_values_match(row1, row2):
        """检查两行的值是否匹配(考虑位置互换)"""
        values1 = set([row1[columns[0]], row1[columns[1]], row1[columns[2]]])
        values2 = set([row2[columns[0]], row2[columns[1]], row2[columns[2]]])
        return values1 == values2

    def count_common_values(row1, row2):
        """计算两行之间共同值的数量(考虑位置互换)"""
        values1 = set([row1[columns[0]], row1[columns[1]], row1[columns[2]]])
        values2 = set([row2[columns[0]], row2[columns[1]], row2[columns[2]]])
        return len(values1 & values2)

    print("开始主要处理流程...")
    for idx, current_row in df.iterrows():
        if idx % 1000 == 0:
            print(f"已处理 {idx}/{len(df)} 行...")
            
        # === 阶段1: 完全相同值检测 ===
        skip_row = False
        for processed_row in processed_rows:
            if check_values_match(current_row, processed_row):
                print(f"跳过完全相同值行: {current_row[columns].tolist()}")
                skip_row = True
                break
        if skip_row:
            continue
            
        # === 阶段2: 两值相同检测 ===
        two_value_matches = 0
        for processed_row in processed_rows:
            if count_common_values(current_row, processed_row) >= 2:
                two_value_matches += 1
                if two_value_matches >= max_two_duplicate_rows:
                    print(f"跳过两值重复超限行: {current_row[columns].tolist()}")
                    skip_row = True
                    break
        if skip_row:
            continue
            
        # === 阶段3: 单个值频次检测 ===
        current_values = current_row[columns].tolist()
        skip_row = False
        for val in current_values:
            if value_counts.get(val, 0) >= max_value_frequency:
                print(f"跳过超频值行: {current_values}, 超频值: {val} (已出现{value_counts[val]}次)")
                skip_row = True
                break
        if skip_row:
            continue
            
        # 更新值计数
        for val in current_values:
            value_counts[val] = value_counts.get(val, 0) + 1
            
        # 通过所有检测，保留该行
        new_row = current_row.copy()
        new_row[sum_field] = rate_sums[current_row['ids_set']]
        result_rows.append(new_row)
        processed_rows.append(current_row)
    
    # 创建结果DataFrame
    result = pd.DataFrame(result_rows)
    
    # 清理临时列
    if 'ids_set' in result.columns:
        result.drop(['ids_set'], axis=1, inplace=True)
    
    # 打印处理结果
    end_time = time.time()
    processing_time = round(end_time - start_time, 2)
    
    print("\n=== 处理完成 ===")
    print(f"处理耗时：{processing_time} 秒")
    print(f"原始数据行数：{original_count}")
    print(f"处理后行数：{len(result)}")
    print(f"去重率：{((original_count - len(result)) / original_count * 100):.2f}%")
    
    return result

# 执行入口
if __name__ == "__main__":
    # === 可配置参数 ===
    file_path = "e:/CODE/dataAnalysis/TEST/test.xlsx"
    sheet_name = "Sheet2"
    columns = ["ID1", "heroID2", "heroID3"]
    sum_field = "场次"  # 需要累加的字段名，可以更换为其他字段
    max_two_duplicate_rows = 3  # 可设置保留两列重复的最大数量
    max_value_frequency = 3     # 可设置单个值出现的最大次数

    try:
        # 执行去重处理
        result_df = deduplicate_excel_optimized(
            file_path=file_path,
            sheet_name=sheet_name,
            columns=columns,
            sum_field=sum_field,  # 添加累加字段参数
            max_two_duplicate_rows=max_two_duplicate_rows,
            max_value_frequency=max_value_frequency
        )

        # 保存结果
        result_df.to_excel("e:/CODE/dataAnalysis/TEST/deduplicated_result.xlsx", index=False)
        print(f"\n结果已保存至：deduplicated_result.xlsx")
        
    except Exception as e:
        print(f"处理过程中出现错误：{str(e)}")
    
    finally:
        input("\n按回车键退出...")