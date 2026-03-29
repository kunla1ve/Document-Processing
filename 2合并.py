# -*- coding: utf-8 -*-
"""
Created on Sun Mar 29 17:55:07 2026

@author: hongl
"""
import pandas as pd
import numpy as np
import os
import glob
from pathlib import Path

# 设置文件夹路径
folder_path = r'C:\Users\hongl\Desktop\新建文件夹'
output_path = r'C:\Users\hongl\Desktop\sample'

# 创建输出文件夹
os.makedirs(output_path, exist_ok=True)

# 获取所有xls和csv文件
all_files = glob.glob(os.path.join(folder_path, '*.xls')) + \
            glob.glob(os.path.join(folder_path, '*.xlsx')) + \
            glob.glob(os.path.join(folder_path, '*.csv'))

def force_read_excel(file_path):
    """强制读取Excel文件，多种方法尝试"""
    
    # 方法1：使用xlrd 1.2.0版本（支持旧格式）
    try:
        import xlrd
        # 尝试不同版本的处理方式
        try:
            workbook = xlrd.open_workbook(file_path, ignore_workbook_corruption=True, 
                                         formatting_info=False, on_demand=True)
        except:
            workbook = xlrd.open_workbook(file_path, ignore_workbook_corruption=True)
        
        sheet = workbook.sheet_by_index(0)
        data = []
        for row_idx in range(sheet.nrows):
            row = []
            for col_idx in range(sheet.ncols):
                cell = sheet.cell(row_idx, col_idx)
                if cell.ctype == xlrd.XL_CELL_TEXT:
                    row.append(cell.value)
                elif cell.ctype == xlrd.XL_CELL_NUMBER:
                    row.append(cell.value)
                elif cell.ctype == xlrd.XL_CELL_DATE:
                    row.append(cell.value)
                else:
                    row.append(None)
            data.append(row)
        
        if data and len(data) > 1:
            df = pd.DataFrame(data[1:], columns=data[0])
            return df
    except Exception as e:
        pass
    
    # 方法2：尝试用pandas的xlrd引擎
    try:
        return pd.read_excel(file_path, engine='xlrd')
    except:
        pass
    
    # 方法3：尝试用calamine引擎（Rust实现，更强大）
    try:
        return pd.read_excel(file_path, engine='calamine')
    except:
        pass
    
    
    return None

def parse_style1(df, filename):
    """样式1：平仓检核表"""
    try:
        for idx, row in df.iterrows():
            row_str = ' '.join([str(x) for x in row.values if pd.notna(x)])
            if '市场' in row_str and '商品' in row_str and '平仓量' in row_str:
                data_start = idx + 1
                data_rows = []
                for i in range(data_start, min(len(df), data_start + 100)):
                    row_data = df.iloc[i]
                    if len(row_data) >= 7:
                        exch_cd = row_data.iloc[1] if pd.notna(row_data.iloc[1]) else None
                        com_cd = row_data.iloc[2] if pd.notna(row_data.iloc[2]) else None
                        contract_date = row_data.iloc[3] if len(row_data) > 3 and pd.notna(row_data.iloc[3]) else None
                        qty = row_data.iloc[6] if pd.notna(row_data.iloc[6]) else None
                        
                        if exch_cd and com_cd and qty and str(exch_cd) != 'nan' and str(com_cd) != 'nan':
                            try:
                                qty = float(qty)
                                data_rows.append({
                                    'exch_cd': exch_cd,
                                    'com_cd': com_cd,
                                    'contract_date': contract_date,
                                    'closeout_qty': qty,
                                    'source_file': filename,
                                    'com_type': None,
                                    'client_no': None,
                                    'counter_party': None,
                                    'traded_strike_price': None,
                                    'call_put': None,
                                    'month_value_date': None
                                })
                            except:
                                pass
                
                if data_rows:
                    return pd.DataFrame(data_rows)
                break
    except:
        pass
    return pd.DataFrame()

def parse_style2(df, filename):
    """样式2：GOLDEN HEN SECURITIES"""
    try:
        for idx, row in df.iterrows():
            row_str = ' '.join([str(x) for x in row.values if pd.notna(x)])
            if '交易所' in row_str and '產品' in row_str:
                data_start = idx + 1
                data_rows = []
                for i in range(data_start, min(len(df), data_start + 100)):
                    row_data = df.iloc[i]
                    if len(row_data) >= 7:
                        exch_cd = row_data.iloc[0] if pd.notna(row_data.iloc[0]) else None
                        com_cd = row_data.iloc[2] if pd.notna(row_data.iloc[2]) else None
                        contract_date = row_data.iloc[3] if pd.notna(row_data.iloc[3]) else None
                        qty = row_data.iloc[6] if pd.notna(row_data.iloc[6]) else None
                        
                        if exch_cd and com_cd and qty and str(exch_cd) != 'nan' and str(com_cd) != 'nan':
                            if qty != '-' and qty != 0:
                                try:
                                    qty = float(qty)
                                    data_rows.append({
                                        'exch_cd': exch_cd,
                                        'com_cd': com_cd,
                                        'contract_date': contract_date if contract_date != '-' else None,
                                        'closeout_qty': qty,
                                        'source_file': filename,
                                        'com_type': 'F',
                                        'client_no': None,
                                        'counter_party': None,
                                        'traded_strike_price': None,
                                        'call_put': None,
                                        'month_value_date': None
                                    })
                                except:
                                    pass
                
                if data_rows:
                    return pd.DataFrame(data_rows)
                break
    except:
        pass
    return pd.DataFrame()

def parse_style3(df, filename):
    """样式3：UOB KAY HIAN - 保留MONTH列作为独立字段"""
    try:
        if 'QTY' in df.columns and 'PRODUCT' in df.columns:
            df_clean = df[df['QTY'].notna() & (df['QTY'] != '')]
            if not df_clean.empty:
                data_rows = []
                for _, row in df_clean.iterrows():
                    try:
                        qty = float(row['QTY'])
                        product = row['PRODUCT']
                        month = row['MONTH'] if 'MONTH' in df.columns else None
                        
                        # 尝试获取client_no，如果有相关列
                        client_no = None
                        if 'ACCOUNT' in df.columns:
                            client_no = row['ACCOUNT'] if pd.notna(row['ACCOUNT']) else None
                        elif 'CLIENT' in df.columns:
                            client_no = row['CLIENT'] if pd.notna(row['CLIENT']) else None
                        
                        if pd.notna(qty) and pd.notna(product):
                            data_rows.append({
                                'exch_cd': None,
                                'com_cd': product,
                                'contract_date': None,
                                'closeout_qty': qty,
                                'source_file': filename,
                                'com_type': None,
                                'client_no': client_no,
                                'counter_party': None,
                                'traded_strike_price': None,
                                'call_put': None,
                                'month_value_date': None,  # 这个字段留空
                                'month': month  # 新增独立的MONTH字段
                            })
                    except:
                        pass
                
                if data_rows:
                    return pd.DataFrame(data_rows)
    except:
        pass
    return pd.DataFrame()

def parse_style4(df, filename):
    """样式4：Account Number标准格式 - 保留client_no和month_value_date"""
    try:
        if 'Market' in df.columns and 'Product Name' in df.columns and 'Closeout Quantity' in df.columns:
            data_rows = []
            for _, row in df.iterrows():
                try:
                    market = row['Market'] if pd.notna(row['Market']) else None
                    product = row['Product Name'] if pd.notna(row['Product Name']) else None
                    qty = float(row['Closeout Quantity']) if pd.notna(row['Closeout Quantity']) else None
                    month_date = row['Month/Value Date'] if 'Month/Value Date' in df.columns and pd.notna(row['Month/Value Date']) else None
                    account = row['Account Number'] if 'Account Number' in df.columns and pd.notna(row['Account Number']) else None
                    
                    if market and product and qty:
                        data_rows.append({
                            'exch_cd': market,
                            'com_cd': product,
                            'contract_date': None,
                            'closeout_qty': qty,
                            'source_file': filename,
                            'com_type': None,
                            'client_no': account,
                            'counter_party': None,
                            'traded_strike_price': None,
                            'call_put': None,
                            'month_value_date': month_date,  # 保留这个字段
                            'month': None  # 这个字段留空
                        })
                except:
                    pass
            
            if data_rows:
                return pd.DataFrame(data_rows)
    except:
        pass
    return pd.DataFrame()

def parse_standard(df, filename):
    """标准样式 - 保留所有可能的列"""
    # 标准所需的列
    required_cols = ['client_no', 'com_type', 'exch_cd', 'com_cd', 'contract_date', 
                     'counter_party', 'traded_strike_price', 'call_put', 'closeout_qty']
    
    # 检查是否包含所有必需的列
    if all(col in df.columns for col in required_cols):
        # 检查是否有额外的列需要保留
        additional_cols = ['month_value_date', 'MONTH', 'Account Number', 'PRODUCT', 'QTY']
        cols_to_keep = required_cols.copy()
        
        # 添加额外存在的列
        for col in additional_cols:
            if col in df.columns and col not in cols_to_keep:
                cols_to_keep.append(col)
        
        df_out = df[cols_to_keep].copy()
        df_out['source_file'] = filename
        
        # 确保month_value_date存在
        if 'month_value_date' not in df_out.columns:
            df_out['month_value_date'] = None
        
        return df_out
    
    return pd.DataFrame()

def parse_csv_flexible(file_path):
    """灵活解析CSV文件"""
    for encoding in ['utf-8', 'gbk', 'gb2312', 'big5', 'latin1', 'cp1252']:
        try:
            df = pd.read_csv(file_path, encoding=encoding)
            if df.shape[0] > 0 and len(df.columns) > 1:
                if df.dropna(how='all').shape[0] > 0:
                    return df
        except:
            continue
    
    for encoding in ['utf-8', 'gbk', 'gb2312', 'big5', 'latin1', 'cp1252']:
        try:
            df = pd.read_csv(file_path, encoding=encoding, sep='\t')
            if df.shape[0] > 0 and len(df.columns) > 1:
                if df.dropna(how='all').shape[0] > 0:
                    return df
        except:
            continue
    
    return None

# 存储所有数据
all_data = []
success_count = 0
failed_files = []

# 用于记录所有出现过的列名
all_columns = set()

for file in all_files:
    filename = os.path.basename(file)
    print(f"处理: {filename}")
    
    # 读取文件
    df = None
    if file.endswith('.csv'):
        df = parse_csv_flexible(file)
        if df is None:
            failed_files.append(filename)
            print(f"  ❌ 无法读取CSV")
            continue
    else:
        df = force_read_excel(file)
        if df is None or df.empty:
            failed_files.append(filename)
            print(f"  ❌ 无法读取Excel")
            continue
    
    # 清理列名
    df.columns = [str(col).strip() for col in df.columns]
    
    # 尝试各种样式解析
    parsed = False
    result_df = None
    
    # 按优先级解析
    result_df = parse_style3(df, filename)
    if not result_df.empty:
        all_data.append(result_df)
        success_count += 1
        all_columns.update(result_df.columns)
        print(f"  ✅ 样式3 (UOB) - {len(result_df)}行")
        continue
    
    result_df = parse_style4(df, filename)
    if not result_df.empty:
        all_data.append(result_df)
        success_count += 1
        all_columns.update(result_df.columns)
        print(f"  ✅ 样式4 - {len(result_df)}行")
        continue
    
    result_df = parse_style1(df, filename)
    if not result_df.empty:
        all_data.append(result_df)
        success_count += 1
        all_columns.update(result_df.columns)
        print(f"  ✅ 样式1 (平仓检核表) - {len(result_df)}行")
        continue
    
    result_df = parse_style2(df, filename)
    if not result_df.empty:
        all_data.append(result_df)
        success_count += 1
        all_columns.update(result_df.columns)
        print(f"  ✅ 样式2 - {len(result_df)}行")
        continue
    
    result_df = parse_standard(df, filename)
    if not result_df.empty:
        all_data.append(result_df)
        success_count += 1
        all_columns.update(result_df.columns)
        print(f"  ✅ 标准样式 - {len(result_df)}行")
        continue
    
    failed_files.append(filename)
    print(f"  ❌ 无法识别样式")

# 合并所有数据
if all_data:
    # 使用concat时设置sort=False，并处理列的对齐
    merged_df = pd.concat(all_data, ignore_index=True, sort=False)
    
    # 清理数据：删除com_cd和closeout_qty都为空的行
    merged_df = merged_df.dropna(subset=['com_cd', 'closeout_qty'], how='all')
    
    # 确保所有需要的列都存在
    required_columns = ['exch_cd', 'com_cd', 'contract_date', 'closeout_qty', 
                       'source_file', 'com_type', 'client_no', 'counter_party',
                       'traded_strike_price', 'call_put', 'month_value_date']
    
    for col in required_columns:
        if col not in merged_df.columns:
            merged_df[col] = None
    
    # 重新排列列的顺序，将重要列放在前面
    column_order = ['source_file', 'exch_cd', 'com_cd', 'closeout_qty', 
                   'client_no', 'month_value_date', 'contract_date', 
                   'com_type', 'counter_party', 'traded_strike_price', 'call_put']
    
    # 添加其他可能存在的列
    other_columns = [col for col in merged_df.columns if col not in column_order]
    final_columns = column_order + other_columns
    
    merged_df = merged_df[final_columns]
    
    # 重置索引
    merged_df = merged_df.reset_index(drop=True)
    
    # 保存结果为CSV文件（使用UTF-8 with BOM编码，兼容Excel打开）
    output_file = os.path.join(output_path, 'merged_result.csv')
    merged_df.to_csv(output_file, index=False, encoding='utf-8-sig')
    
    print(f"\n{'='*60}")
    print(f"✅ 合并完成！")
    print(f"📁 处理文件总数: {len(all_files)}")
    print(f"✅ 成功解析: {success_count} 个文件")
    print(f"❌ 失败文件: {len(failed_files)} 个")
    print(f"📄 输出CSV文件: {output_file}")
    print(f"📋 总列数: {len(merged_df.columns)}")
    print(f"📋 列名: {', '.join(merged_df.columns)}")
    
    if failed_files:
        print(f"\n失败文件列表:")
        for f in failed_files:
            print(f"  - {f}")
    
else:
    print("❌ 没有找到可合并的数据")