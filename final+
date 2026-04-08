# -*- coding: utf-8 -*-
"""
Created on Tue Mar 31 21:06:46 2026

@author: kunlave
"""

"""
PDF/Excel批量转换与合并工具
将文件夹中的PDF表格和Excel文件转换为统一的CSV格式
"""
import csv
import pdfplumber
import pandas as pd
import numpy as np
import os
import re
import glob
from pathlib import Path

# ==================== 配置 ====================
# 自动获取桌面路径
DESKTOP = Path.home() / "Desktop"
INPUT_FOLDER = DESKTOP /  "新增資料夾 (2)" / "20260304" 
OUTPUT_FOLDER = DESKTOP
mapping_csv_path  = DESKTOP / 'contract_spec1_updated.csv'




def load_mapping(csv_file_path):
    """加载CSV文件并构建多个映射字典，将各种输入键映射到(CQ Exchange, CQ Code, Product Name)"""
    # 存储各种映射：key -> (CQ Exchange, CQ Code, Product Name)
    # key的格式为 (exchange, code) 的元组
    exchange_code_to_info = {}

    with open(csv_file_path, mode='r', encoding='utf-8-sig') as file:
        reader = csv.DictReader(file)
        for row in reader:
            cq_ex = row['CQ Exchange'].strip()
            cq_code = row['CQ Code'].strip()
            product_name = row['Product Name'].strip()
            
            if not cq_ex or not cq_code:
                continue

            # 1. CQ Exchange + CQ Code 本身
            key_cq = (row['CQ Exchange'].strip(), row['CQ Code'].strip())
            exchange_code_to_info[key_cq] = (cq_ex, cq_code, product_name)

            # 2. ES Exchange + ES Code
            es_ex = row['ES Exchange'].strip()
            es_code = row['ES Code'].strip()
            if es_ex and es_code:
                exchange_code_to_info[(es_ex, es_code)] = (cq_ex, cq_code, product_name)

            # 3. EX Exchange + EX Code
            ex_ex = row['EX Exchange'].strip()
            ex_code = row['EX Code'].strip()
            if ex_ex and ex_code:
                exchange_code_to_info[(ex_ex, ex_code)] = (cq_ex, cq_code, product_name)

            # 4. SP Code 有四个字段：SP Code, SP Code 1, SP Code 2, SP Code 3
            # 注意：SP Exchange 可能与 SP Code 配对
            sp_ex = row['SP Exchange'].strip()
            sp_codes = [
                row['SP Code'].strip(),
                row['SP Code 1'].strip(),
                row['SP Code 2'].strip(),
                row['SP Code 3'].strip()
            ]
            # 对于每个有效的 SP Code，如果 SP Exchange 存在，则建立映射
            if sp_ex:
                for spc in sp_codes:
                    if spc and spc != '-':
                        exchange_code_to_info[(sp_ex, spc)] = (cq_ex, cq_code, product_name)

    return exchange_code_to_info

# ==================== PDF处理 ====================
def pdf_to_dataframes(pdf_path):
    """将PDF转换为DataFrame列表"""
    all_dfs = []
    
    settings_bordered = {
        "vertical_strategy": "lines", "horizontal_strategy": "lines",
        "snap_tolerance": 3, "intersection_tolerance": 3,
    }
    settings_borderless = {
        "vertical_strategy": "text", "horizontal_strategy": "text",
        "snap_tolerance": 3, "intersection_tolerance": 3, "text_tolerance": 3,
    }
    
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables(settings_bordered) or page.extract_tables(settings_borderless)
            
            for table in tables:
                if table and len(table) > 1:
                    df = pd.DataFrame(table).replace('', None).replace(r'^\s*$', None, regex=True)
                    if len(df) > 0 and df.iloc[0].notna().any():
                        df.columns = df.iloc[0]
                        df = df[1:].dropna(how='all').reset_index(drop=True)
                        if not df.empty:
                            all_dfs.append(df)
    return all_dfs


# ==================== Excel/CSV读取 ====================
def read_excel_flexible(file_path):
    """灵活读取Excel文件"""
    try:
        import xlrd
        try:
            wb = xlrd.open_workbook(file_path, ignore_workbook_corruption=True, formatting_info=False)
        except:
            wb = xlrd.open_workbook(file_path, ignore_workbook_corruption=True)
        
        sheet = wb.sheet_by_index(0)
        data = [[cell.value for cell in sheet.row(r)] for r in range(sheet.nrows)]
        if data and len(data) > 1:
            return pd.DataFrame(data[1:], columns=data[0])
    except:
        pass
    
    for engine in ['xlrd', 'calamine', None]:
        try:
            return pd.read_excel(file_path, engine=engine)
        except:
            continue
    return None


def read_csv_flexible(file_path):
    """灵活读取CSV文件"""
    for encoding in ['utf-8', 'gbk', 'gb2312', 'big5', 'latin1']:
        for sep in [',', '\t']:
            try:
                df = pd.read_csv(file_path, encoding=encoding, sep=sep)
                if df.shape[0] > 0 and df.dropna(how='all').shape[0] > 0:
                    return df
            except:
                continue
    return None


# ==================== 样式解析 ====================
def parse_table(df, filename, style_name, detect_func, extract_func):
    """通用解析模板"""
    try:
        for idx, row in df.iterrows():
            if detect_func(row):
                data_rows = extract_func(df, idx + 1, filename)
                if data_rows:
                    return pd.DataFrame(data_rows)
                break
    except:
        pass
    return pd.DataFrame()


def detect_style1(row):
    return '市场' in str(row.values) and '商品' in str(row.values) and '平仓量' in str(row.values)


def extract_style1(df, start_idx, filename):
    rows = []
    for i in range(start_idx, min(len(df), start_idx + 100)):
        row = df.iloc[i]
        if len(row) >= 7:
            exch = row.iloc[1] if pd.notna(row.iloc[1]) else None
            com = row.iloc[2] if pd.notna(row.iloc[2]) else None
            qty = row.iloc[6]
            if exch and com and qty and str(qty) not in ['nan', '-']:
                try:
                    rows.append({
                        'exch_cd': exch, 'com_cd': com, 'contract_date': row.iloc[3],
                        'closeout_qty': float(qty), 'source_file': filename,
                        'com_type': None, 'client_no': None, 'counter_party': None,
                        'traded_strike_price': None, 'call_put': None, 'month_value_date': None
                    })
                except:
                    pass
    return rows


def detect_style2(row):
    return '交易所' in str(row.values) and '產品' in str(row.values)


def extract_style2(df, start_idx, filename):
    rows = []
    for i in range(start_idx, min(len(df), start_idx + 100)):
        row = df.iloc[i]
        if len(row) >= 7:
            exch = row.iloc[0] if pd.notna(row.iloc[0]) else None
            com = row.iloc[2] if pd.notna(row.iloc[2]) else None
            qty = row.iloc[6]
            if exch and com and qty and str(qty) not in ['nan', '-', '0']:
                try:
                    rows.append({
                        'exch_cd': exch, 'com_cd': com, 'contract_date': row.iloc[3] if row.iloc[3] != '-' else None,
                        'closeout_qty': float(qty), 'source_file': filename,
                        'com_type': 'F', 'client_no': None, 'counter_party': None,
                        'traded_strike_price': None, 'call_put': None, 'month_value_date': None
                    })
                except:
                    pass
    return rows


def parse_style3(df, filename):
    """UOB格式"""
    if 'QTY' in df.columns and 'PRODUCT' in df.columns:
        df_clean = df[df['QTY'].notna() & (df['QTY'] != '')]
        rows = []
        for _, row in df_clean.iterrows():
            try:
                rows.append({
                    'exch_cd': None, 'com_cd': row['PRODUCT'], 'contract_date': None,
                    'closeout_qty': float(row['QTY']), 'source_file': filename,
                    'com_type': None, 'client_no': row.get('ACCOUNT') or row.get('CLIENT'),
                    'counter_party': None, 'traded_strike_price': None, 'call_put': None,
                    'month_value_date': None, 'month': row.get('MONTH')
                })
            except:
                pass
        if rows:
            return pd.DataFrame(rows)
    return pd.DataFrame()


def parse_style4(df, filename):
    """标准Account Number格式"""
    required = ['Market', 'Product Name', 'Closeout Quantity']
    if all(c in df.columns for c in required):
        rows = []
        for _, row in df.iterrows():
            try:
                if pd.notna(row['Market']) and pd.notna(row['Product Name']):
                    rows.append({
                        'exch_cd': row['Market'], 'com_cd': row['Product Name'], 'contract_date': None,
                        'closeout_qty': float(row['Closeout Quantity']), 'source_file': filename,
                        'com_type': None, 'client_no': row.get('Account Number'),
                        'counter_party': None, 'traded_strike_price': None, 'call_put': None,
                        'month_value_date': row.get('Month/Value Date'), 'month': None
                    })
            except:
                pass
        if rows:
            return pd.DataFrame(rows)
    return pd.DataFrame()


def parse_standard(df, filename):
    """标准格式（包含所有必需列）- 兼容多种列名变体"""
    # 定义必需的列名（使用标准名称）
    required = ['client_no', 'com_type', 'exch_cd', 'com_cd', 'contract_date', 
                'counter_party', 'traded_strike_price', 'call_put', 'closeout_qty']
    
    # 列名映射表（处理常见的列名变体）
    column_mapping = {
        'closedout_qty': 'closeout_qty',  # 处理拼写变体
        'closeout_qty': 'closeout_qty',   # 标准名称
        # 可以继续添加其他可能的变体
    }
    
    # 应用列名映射
    for old_name, new_name in column_mapping.items():
        if old_name in df.columns and new_name not in df.columns:
            df = df.rename(columns={old_name: new_name})
    
    # 检查所有必需列是否存在
    if all(c in df.columns for c in required):
        df_out = df[required].copy()
        df_out['source_file'] = filename
        
        # 添加 month_value_date 列（如果不存在）
        if 'month_value_date' not in df_out.columns:
            df_out['month_value_date'] = None
            
        return df_out
    
    return pd.DataFrame()

# ==================== 数据处理函数 ====================
def clean_product_code(code):
    """清理产品代码（去除数字结尾的字符），返回option"""
    if pd.isna(code):
        return code
    
    code = str(code).strip()
    
    # 如果代码为空或长度小于等于2，直接返回
    if not code or len(code) <= 3:
        return code
    
    # 如果代码中包含点号或空格，不做任何修改，直接返回原值
    if '.' in code or ' ' in code:
        return code
    
    # 检查最后一个字符是否为数字
    if code[-1].isdigit():
        # 检查最后两位是否都是数字
        if len(code) >= 3 and code[-2].isdigit():
            # 结尾为2位数字，去除3个字符（最后两位数字）
            return code[:-3]
        else:
            # 结尾为单数字，去除2个字符（最后一位数字）
            return code[:-2]
    
    return code


def parse_contract_date(date_value):
    """将日期转换为标准合约代码（如M2026）"""
    month_map = {1: 'F', 2: 'G', 3: 'H', 4: 'J', 5: 'K', 6: 'M',
                 7: 'N', 8: 'Q', 9: 'U', 10: 'V', 11: 'X', 12: 'Z'}
    month_3letter = {'JAN': 'F', 'FEB': 'G', 'MAR': 'H', 'APR': 'J', 'MAY': 'K', 'JUN': 'M',
                     'JUL': 'N', 'AUG': 'Q', 'SEP': 'U', 'OCT': 'V', 'NOV': 'X', 'DEC': 'Z'}
    
    if pd.isna(date_value):
        return None
    
    date_str = str(date_value).strip()
    
    # 已是标准格式
    if re.match(r'^[A-Z]\d{4}$', date_str):
        return date_str
    
    # Excel数字日期
    if date_str.replace('.', '').isdigit() and len(date_str) >= 5:
        try:
            d = pd.to_datetime('1899-12-30') + pd.Timedelta(days=int(float(date_str)))
            return f"{month_map[d.month]}{d.year}"
        except:
            pass
    
    # 2604, 2606 格式
    if date_str.isdigit() and len(date_str) == 4:
        return f"{month_map.get(int(date_str[2:]), '')}20{date_str[:2]}"
    
    # 202604, 202606 格式
    if date_str.isdigit() and len(date_str) == 6:
        return f"{month_map.get(int(date_str[4:]), '')}{date_str[:4]}"
    
    # 三字母月份
    for m3, mcode in month_3letter.items():
        if m3 in date_str.upper():
            year_match = re.search(r'\d{2,4}', date_str)
            if year_match:
                y = year_match.group()
                y = y if len(y) == 4 else f'20{y}'
                return f"{mcode}{y}"
    
    # 标准日期解析
    try:
        d = pd.to_datetime(date_str)
        return f"{month_map[d.month]}{d.year}"
    except:
        return None



def add_mapping_info_to_df(df, mapping_csv_path):
    """
    为DataFrame添加映射信息
    参数:
        df: 包含 exch_cd 和 com_cd 列的DataFrame
        mapping_csv_path: 映射CSV文件路径
    返回:
        添加了 product_name 和 after_map_com_cd 列的DataFrame
    """
    # 加载映射
    mapping = load_mapping(mapping_csv_path)
    
    
    # 遍历每一行进行匹配
    for idx, row in df.iterrows():
        exch = str(row['exch_cd']).strip().upper() if pd.notna(row['exch_cd']) else ''
        code = str(row['com_cd']).strip().upper() if pd.notna(row['com_cd']) else ''
        
        if exch and code:
            key = (exch, code)
            if key in mapping:
                cq_ex, cq_code, product_name = mapping[key]
                df.at[idx, 'product_name'] = product_name
                df.at[idx, 'after_map_com_cd'] = cq_code
    
    return df

# ==================== 主处理流程 ====================
def process_files():
    """主处理函数"""
    # 创建输出文件夹
    OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)
    
    # 获取所有文件
    pdf_files = set(INPUT_FOLDER.glob('*.pdf')) | set(INPUT_FOLDER.glob('*.PDF'))
    excel_files = set(glob.glob(str(INPUT_FOLDER / '*.xls'))) | \
                  set(glob.glob(str(INPUT_FOLDER / '*.xlsx'))) | \
                  set(glob.glob(str(INPUT_FOLDER / '*.csv')))
    
    all_files = list(pdf_files) + [Path(f) for f in excel_files]
    
    if not all_files:
        print(f"❌ 在 {INPUT_FOLDER} 中没有找到文件")
        return
    
    print(f"📁 找到 {len(all_files)} 个文件\n")
    
    all_data = []
    success_count = 0
    failed_files = []
    
    for file_path in all_files:
        filename = file_path.name
        print(f"处理: {filename}")
        
        # 读取文件
        if file_path.suffix.lower() == '.pdf':
            dfs = pdf_to_dataframes(file_path)
            if not dfs:
                failed_files.append(filename)
                print(f"  ❌ 无法读取PDF")
                continue
        elif file_path.suffix.lower() == '.csv':
            df = read_csv_flexible(file_path)
            dfs = [df] if df is not None else []
        else:
            df = read_excel_flexible(file_path)
            dfs = [df] if df is not None else []
        
        if not dfs:
            failed_files.append(filename)
            print(f"  ❌ 无法读取文件")
            continue
        
        # 解析每个表格
        for df in dfs:
            if df is None or df.empty:
                continue
            
            df.columns = [str(c).strip() for c in df.columns]
            result = None
            
            # 按优先级尝试解析
            parsers = [
                (parse_style3, "样式3 (UOB)"),
                (parse_style4, "样式4"),
                (lambda d, f: parse_table(d, f, "样式1", detect_style1, extract_style1), "样式1 (平仓检核表)"),
                (lambda d, f: parse_table(d, f, "样式2", detect_style2, extract_style2), "样式2"),
                (parse_standard, "标准样式")
            ]
            
            for parser, name in parsers:
                result = parser(df, filename)
                if not result.empty:
                    all_data.append(result)
                    success_count += 1
                    print(f"  ✅ {name} - {len(result)}行")
                    break
            
            if result is None or result.empty:
                print(f"  ❌ 无法识别样式")
        
        if not any(not r.empty for r in all_data[-len(dfs):] if all_data):
            failed_files.append(filename)
    
    # 合并数据
    if not all_data:
        print("❌ 没有找到可合并的数据")
        return
    
    merged = pd.concat(all_data, ignore_index=True, sort=False)
    merged = merged.dropna(subset=['com_cd', 'closeout_qty'], how='all')
    
    # 数据清洗
    merged['closeout_qty'] = pd.to_numeric(merged['closeout_qty'], errors='coerce')
    merged = merged[merged['closeout_qty'].notna()]
    merged['counter_party'] = 'PSC0000'
    
    # 根据 com_cd 列设置 com_type
    # 如果 com_cd 列存在，则根据是否包含点号或空格来设置 com_type
    if 'com_cd' in merged.columns:
        # 检查 com_cd 中是否包含点号或空格，包含则为 'O'，否则为 'F'
        merged['com_type'] = merged['com_cd'].apply(
            lambda x: 'O' if (isinstance(x, str) and ('.' in x )) else 'F'
        )
    else:
        # 如果 com_cd 列不存在，设置默认值或根据需求处理
        merged['com_type'] = 'F'  # 或设置为其他默认值
    
    
    # 合并client_no列
    if 'Client_no' in merged.columns:
        merged['client_no'] = merged.get('client_no', '').fillna('') + merged['Client_no'].fillna('')
        merged = merged.drop('Client_no', axis=1)
    
    # 生成合约代码
    for col in ['contract_date', 'month', 'month_value_date']:
        if col in merged.columns:
            merged['contract_date'] = merged['contract_date'].fillna(merged[col])
    merged['contract_date'] = merged['contract_date'].apply(parse_contract_date)
    
    # 删除不需要的列
    drop_cols = ['month_value_date', 'month', 'traded_strike_price', 'call_put']
    merged = merged.drop([c for c in drop_cols if c in merged.columns], axis=1)
    
    # 先清理代码
    merged['com_cd'] = merged['com_cd'].apply(clean_product_code)
    
    
    # 添加映射信息
    merged = add_mapping_info_to_df(merged, mapping_csv_path)   
    
    
    
    # 调整列顺序
    col_order = ['source_file','product_name',  'after_map_com_cd', 'client_no','com_cd',
                  'exch_cd', 'com_type','contract_date',  'counter_party','closeout_qty']
    final_cols = col_order + [c for c in merged.columns if c not in col_order]
    merged = merged[final_cols]
    
    # 保存结果
    output_file = OUTPUT_FOLDER / 'merged_result.csv'
    merged.to_csv(output_file, index=False, encoding='utf-8-sig')
    
    print(f"\n{'='*60}")
    print(f"✅ 处理完成！")
    print(f"📁 处理文件总数: {len(all_files)}")
    print(f"✅ 成功解析: {success_count} 个文件")
    print(f"❌ 失败文件: {len(failed_files)} 个")
    print(f"📄 输出: {output_file}")
    print(f"📋 总行数: {len(merged)} 行")
    
    if failed_files:
        print(f"\n失败文件列表:")
        for f in failed_files:
            print(f"  - {f}")


# ==================== 程序入口 ====================
if __name__ == "__main__":
    print(f"输入文件夹: {INPUT_FOLDER}")
    print(f"输出文件夹: {OUTPUT_FOLDER}")
    print()
    process_files()
