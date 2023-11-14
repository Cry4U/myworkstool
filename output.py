from collections import Counter 

def try_parse_int(s):  
    try:
        return int(s)  
    except ValueError:  
        return s  

def process_file(input_file, output_file):  
    with open(input_file, "r", encoding="utf-8") as f:  
        arr1 = [[try_parse_int(i) for i in line.split("\t")] for line in reversed(f.read().split("\n"))]  

    results = []  
    for arr in arr1:  
        first_value = arr[0] if arr else None  
        elements = [e for e in arr[1:] if isinstance(e, int)]  
        # 确定步长
        div_1000 = [e // 1000 for e in elements]  
        div_100 = [e // 100 for e in elements]  
        step = 1000 if sum(i > 2 for i in div_1000) >= 3 else 100 if sum(i >= 2 for i in div_100) >= 3 else 10  
        
        max_value = max(elements)  
        intervals = [(i, i + step) for i in range(0, max_value + 1, step)] 
        counts = [(interval, sum(interval[0] <= e <= interval[1] for e in elements)) for interval in intervals]  
        max_count = max(counts, key=lambda x: (x[1], x[0][1])) 
        output = f'"{first_value}"\t【{intervals[0][0]}~{intervals[0][1]}】\t{counts[0][1]}\t【{max_count[0][0]}~{max_count[0][1]}】\t{max_count[1]}'     
        results.append(output)  
    with open(output_file, "w", encoding="utf-8") as f:  
        f.write("\n".join(reversed(results))) 
process_file("input.txt", "result.txt") 