#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
import sys
from pathlib import Path
from datetime import datetime

def generate_beancount_open_records(xlsx_path: str):
    xlsx_path = Path(xlsx_path)
    if not xlsx_path.exists():
        print(f"文件不存在: {xlsx_path}")
        sys.exit(1)

    # 读取 Excel 文件
    xls = pd.ExcelFile(xlsx_path)
    
    # 按一级字段分组生成文件
    records_by_first_level = {}

    for sheet_name in xls.sheet_names:
        if sheet_name == "履历表":
            continue  # 排除履历表

        df = pd.read_excel(xls, sheet_name=sheet_name)

        # 保留必要列
        required_columns = ["开账时间", "货币", "名称", "账户全名", "备注", "一级"]
        for col in required_columns:
            if col not in df.columns:
                df[col] = ""

        for _, row in df.iterrows():
            first_level = row["一级"] if row["一级"] else "Unknown"
            if first_level not in records_by_first_level:
                records_by_first_level[first_level] = []

            # 处理日期
            date_str = row["开账时间"]
            if pd.isna(date_str) or str(date_str).strip() == "":
                date_str = "1970-01-01"
            else:
                try:
                    date_obj = pd.to_datetime(date_str)
                    date_str = date_obj.strftime("%Y-%m-%d")
                except:
                    date_str = "1970-01-01"

            currency = row["货币"] if (pd.notna(row["货币"]) and str(row["货币"]).strip() != "") else "CNY"
            account_full = row["账户全名"] if row["账户全名"] else ""
            note = row["名称"] if row["名称"] else ""

            # 生成 open 记录
            line = f'{date_str} open {account_full} {currency}'
            if note:
                line += f'  ; {note}'

            records_by_first_level[first_level].append(line)

    # 写入对应一级的 .bean 文件
    for first_level, records in records_by_first_level.items():
        bean_file = Path(f"{first_level}.bean")
        with open(bean_file, "w", encoding="utf-8") as f:
            f.write("\n".join(records))
            f.write("\n")
        print(f"生成 {bean_file} 完成.")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python generate_beancount.py <xlsx文件路径>")
        sys.exit(1)
    xlsx_path = sys.argv[1]
    generate_beancount_open_records(xlsx_path)
