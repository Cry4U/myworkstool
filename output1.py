from collections import Counter

with open("input.txt", "r", encoding="utf-8") as f:
    arr0 = f.read().split("\n")

def try_parse_int(s):
    try:
        return int(s)
    except ValueError:
        return s  # 如果无法解析为整数，就返回原始字符串

arr1 = [[try_parse_int(i) for i in line.split("\t")] for _, line in enumerate(arr0[::-1])]

results = []
for arr in arr1:
    first_value = arr[0] if arr else None
    elements = [e for e in arr[1:] if isinstance(e, int)]
    
    # 确定步长
    div_1000 = [e // 1000 for e in elements]
    if len([i for i in div_1000 if i > 2]) >= 3:
        step = 1000
    elif len([i for i in div_1000 if i >= 2]) >= 3:
        step = 100
    else:
        step = 10

    # 计算区间
    max_value = max(elements)
    intervals = [(i, i + step) for i in range(0, max_value + 1, step)]
    
    # 计算每个区间中的元素个数
    counts = []
    for interval in intervals:
        count = sum(interval[0] <= e <= interval[1] for e in elements)
        counts.append((interval, count))
    
    # 找出元素个数最多的区间，如果个数一样，选择更大的一个区间
    max_count = max(counts, key=lambda x: (x[1], x[0][1]))
    
    # 格式化输出结果
    output = f'"{first_value}"\t【{intervals[0][0]}~{intervals[0][1]}】\t{counts[0][1]}\t【{max_count[0][0]}~{max_count[0][1]}】\t{max_count[1]}'
    
    # 将输出结果添加到结果列表中
    results.append(output)

# 将结果列表中的结果以倒序的方式写入到名为"result.txt"的文件中，每条结果之间用换行符隔开
with open("result.txt", "w", encoding="utf-8") as f:
    for output in reversed(results):
        f.write(output + "\n")