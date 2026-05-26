# -*- coding: utf-8 -*-
"""
将商品编码表Excel转换为JSON格式
"""
import openpyxl
import json
import os

# 源文件路径
source_file = r"d:\code\船务部\WA213 HT26030AE01 - 副本\01.下单生产\所有的源文件\2024.12.5 最新出口商品编码及报关成分.xlsx"

# 输出文件路径
output_file = r"d:\code\船务部\shipping_helper\reference\product_codes.json"

def convert():
    """转换Excel到JSON"""
    print(f"正在读取: {source_file}")

    wb = openpyxl.load_workbook(source_file)
    ws = wb.active

    # 获取表头
    headers = [cell.value for cell in ws[1]]

    # 清理表头：去掉None和重复的
    clean_headers = []
    for i, h in enumerate(headers):
        if h is None:
            clean_headers.append(f"col_{i}")
        else:
            clean_headers.append(h)

    print(f"表头: {clean_headers[:15]}")

    # 提取数据
    codes = []
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True)):
        item = {}
        for j, value in enumerate(row):
            if j < len(clean_headers):
                header = clean_headers[j]
                # 处理空值
                if value is None:
                    value = ""
                # 处理数字类型
                elif isinstance(value, (int, float)):
                    # Excel的日期序列号（大于25569且小于50000的可能是日期）
                    if 25569 < value < 50000 and "日期" in header:
                        import datetime
                        value = (datetime.datetime(1899, 12, 30) + datetime.timedelta(days=value)).strftime("%Y-%m-%d")
                    else:
                        value = str(value)
                else:
                    value = str(value)
                item[header] = value
        if item.get('产品内编'):  # 只保存有内部编号的记录
            codes.append(item)

    print(f"共提取 {len(codes)} 条记录")

    # 构建输出数据
    output_data = {
        "version": "2024.12.5",
        "description": "广东宏昊报关成分确认表",
        "codes": codes
    }

    # 确保目录存在
    os.makedirs(os.path.dirname(output_file), exist_ok=True)

    # 写入JSON文件
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(output_data, f, ensure_ascii=False, indent=2)

    print(f"已保存至: {output_file}")

    # 打印示例数据
    if codes:
        print("\n示例数据（第一条）：")
        for k, v in list(codes[0].items())[:10]:
            print(f"  {k}: {v}")

if __name__ == "__main__":
    convert()
