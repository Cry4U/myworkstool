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
    
    # 使用字典优化存储和查找
    complete_matches = {}  # 存储完全匹配的组合
    two_value_matches = {}  # 存储两值匹配的组合
    value_counts = {}  # 存储单值计数
    
    print("开始主要处理流程...")
    for idx, current_row in df.iterrows():
        current_values = tuple(sorted([current_row[col] for col in columns]))
        
        # === 阶段1: 完全相同值检测 ===
        if current_values in complete_matches:
            print(f"跳过完全相同值行: {list(current_values)}")
            continue
        
        # === 阶段2: 两值相同检测 ===
        skip_row = False
        for i in range(len(current_values)):
            for j in range(i + 1, len(current_values)):
                pair = tuple(sorted([current_values[i], current_values[j]]))
                if pair in two_value_matches:
                    two_value_matches[pair] += 1
                    if two_value_matches[pair] > max_two_duplicate_rows:
                        print(f"跳过两值重复超限行: {list(current_values)}")
                        skip_row = True
                        break
                else:
                    two_value_matches[pair] = 1
            if skip_row:
                break
        if skip_row:
            continue
            
        # === 阶段3: 单个值频次检测 ===
        skip_row = False
        for val in current_values:
            if value_counts.get(val, 0) >= max_value_frequency:
                print(f"跳过超频值行: {list(current_values)}, 超频值: {val}")
                skip_row = True
                break
        if skip_row:
            continue
            
        # 更新记录
        complete_matches[current_values] = True
        for val in current_values:
            value_counts[val] = value_counts.get(val, 0) + 1
            
        # 保存结果
        new_row = current_row.copy()
        new_row[sum_field] = rate_sums[current_row['ids_set']]
        result_rows.append(new_row)
    
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