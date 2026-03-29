import pandas as pd
import numpy as np
import re

# 读取CSV文件
df = pd.read_csv(r'C:\Users\hongl\Desktop\sample\merged_result.csv', encoding='utf-8')

# 1. 去掉closeout_qty不为数字的所有行
df['closeout_qty'] = pd.to_numeric(df['closeout_qty'], errors='coerce')
df = df[df['closeout_qty'].notna()].copy()

# 2. 把counter_party列全填PSC0000
df['counter_party'] = 'PSC0000'

# 3. 把com_type列全填com_type
df['com_type'] = 'F'

# 4. 把Client_no列合并到client_no，然后删去Client_no
if 'Client_no' in df.columns:
    if 'client_no' in df.columns:
        df['client_no'] = df['client_no'].fillna('') + df['Client_no'].fillna('')
    else:
        df['client_no'] = df['Client_no']
    df = df.drop('Client_no', axis=1)

# 5. 根据contract_date和MONTH,Month/Value Date生成标准合约代码
def parse_to_contract_code(date_value):
    """将各种日期格式转换为标准合约代码，如J2026、M2026等"""
    
    month_map_3letter = {
        'JAN': 'F', 'FEB': 'G', 'MAR': 'H', 'APR': 'J', 'MAY': 'K', 'JUN': 'M',
        'JUL': 'N', 'AUG': 'Q', 'SEP': 'U', 'OCT': 'V', 'NOV': 'X', 'DEC': 'Z'
    }
    
    month_map_num = {
        1: 'F', 2: 'G', 3: 'H', 4: 'J', 5: 'K', 6: 'M',
        7: 'N', 8: 'Q', 9: 'U', 10: 'V', 11: 'X', 12: 'Z'
    }
    
    if pd.isna(date_value):
        return None
    
    date_str = str(date_value).strip()
    
    # 检查是否已经是标准合约代码格式
    if re.match(r'^[A-Z]\d{4}$', date_str):
        return date_str
    
    # 处理Excel数字日期格式
    if date_str.replace('.', '').isdigit() and len(date_str) >= 5:
        try:
            excel_date = int(float(date_str))
            date_obj = pd.to_datetime('1899-12-30') + pd.Timedelta(days=excel_date)
            return f"{month_map_num.get(date_obj.month, '')}{date_obj.year}"
        except:
            pass
    
    # 处理如 2604, 2606 格式
    if date_str.isdigit() and len(date_str) == 4:
        year = '20' + date_str[:2]
        month_num = int(date_str[2:])
        if 1 <= month_num <= 12:
            return f"{month_map_num.get(month_num, '')}{year}"
    
    # 处理如 202604, 202606 格式
    if date_str.isdigit() and len(date_str) == 6:
        year = date_str[:4]
        month_num = int(date_str[4:])
        if 1 <= month_num <= 12:
            return f"{month_map_num.get(month_num, '')}{year}"
    
    # 处理三字母月份格式
    for month_name_3letter, month_code in month_map_3letter.items():
        if month_name_3letter in date_str.upper():
            year_match = re.search(r'\d{2,4}', date_str)
            if year_match:
                year_num = year_match.group()
                year = '20' + year_num if len(year_num) == 2 else year_num
                return f"{month_code}{year}"
    
    # 尝试直接解析日期格式
    try:
        date_obj = pd.to_datetime(date_str)
        return f"{month_map_num.get(date_obj.month, '')}{date_obj.year}"
    except:
        return None

def get_contract_code(row):
    if 'contract_date' in row.index and pd.notna(row['contract_date']):
        code = parse_to_contract_code(row['contract_date'])
        if code:
            return code
    
    if 'MONTH' in row.index and pd.notna(row['MONTH']):
        code = parse_to_contract_code(row['MONTH'])
        if code:
            return code
    
    if 'Month/Value Date' in row.index and pd.notna(row['Month/Value Date']):
        code = parse_to_contract_code(row['Month/Value Date'])
        if code:
            return code
    
    return None

# 生成合约代码并更新contract_date列
df['contract_date'] = df.apply(get_contract_code, axis=1)

# 删除不需要的列
if 'Month/Value Date' in df.columns:
    df = df.drop('Month/Value Date', axis=1)
if 'MONTH' in df.columns:
    df = df.drop('MONTH', axis=1)
if 'traded_strike_price' in df.columns:
    df = df.drop('traded_strike_price', axis=1)
if 'call_put' in df.columns:
    df = df.drop('call_put', axis=1)

# 保存结果
output_path = r'C:\Users\hongl\Desktop\sample\merged_result1.csv'
df.to_csv(output_path, index=False, encoding='utf-8-sig')

print("处理完成！")