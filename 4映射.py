# -*- coding: utf-8 -*-
"""
Created on Sun Mar 29 16:19:42 2026

@author: hongl
"""

import pandas as pd

# 读取CSV文件
df = pd.read_csv(r'C:\Users\hongl\Desktop\sample\merged_result1.csv')

# 映射字典 (Globex代码 -> ClearPort旧码)
mapping = {
    'SI': 'SV', 'MGC': 'MGC', '6B': 'BP', 'QG': 'QG', 'CL': 'CL',
    'CN': 'CN', 'NQ': 'NQ', 'E-MINI S&P': 'E-MINI S&P', 'YM': 'YM',
    'MYM': 'MYM', 'ZB': 'US', '6A': 'AD', 'ES': 'E-MINI S&P',
    'MES': 'MES', 'MNQ': 'MNQ', 'GC': 'GD', 'QI': 'QI', 'PL': 'PL',
    'MET': 'MET', 'MBT': 'MBT', 'TWN': 'TW', 'TOPIXM': 'TPXTOPIXM',
    'BTC': 'BTC', 'SIL': 'SIL', 'UC': 'UC', 'CT': 'CT', 'NK': 'NK',
    'HO': 'HO', 'MEU': 'MEU', 'ZN': 'TY', 'JY': 'JY', 'BO': 'BO',
    'W': 'W', 'GD': 'GC', 'SV': 'SI', 'MAL': 'ML', 'MCU': 'CA'
}

# 产品名称字典
product_names = {
    'SI': '白銀', 'SV': '白銀', 'MGC': '微型黃金', '6B': '英鎊', 'BP': '英鎊',
    'QG': '迷你天然氣', 'CL': '輕原油', 'CN': '肉牛', 'NQ': '迷你納指',
    'E-MINI S&P': '迷你標普', 'YM': '小道瓊', 'MYM': '微型道瓊', 'ZB': '30年美債',
    'US': '30年美債', '6A': '澳元', 'AD': '澳元', 'ES': '迷你標普',
    'MES': '微型標普', 'MNQ': '微型納指', 'GC': '黃金', 'GD': '黃金',
    'QI': '迷你白銀', 'PL': '鉑金', 'MET': '金屬指數', 'MBT': '迷你美債',
    'TWN': '台股指數', 'TW': '台股指數', 'TOPIXM': '東證指數', 'TPXTOPIXM': '東證指數',
    'BTC': '比特幣', 'SIL': '微型白銀', 'UC': '美元指數', 'CT': '棉花',
    'NK': '日經指數', 'HO': '熱燃油', 'MEU': '歐盟指數', 'ZN': '10年美債',
    'TY': '10年美債', 'JY': '日圓', 'BO': '黃豆油', 'W': '小麥',
    'MAL': '馬來西亞棕櫚油', 'ML': '馬來西亞棕櫚油', 'MCU': '高級銅', 'CA': '高級銅'
}

def transform_com_cd(code):
    if pd.isna(code):
        return code
    code = str(code).strip()
    # 如果以数字结尾，去掉最后两个字符
    if code and code[-1].isdigit():
        code = code[:-2]
    # 映射转换
    return mapping.get(code, code)

def get_product_name(code):
    if pd.isna(code):
        return ''
    code = str(code).strip()
    # 如果以数字结尾，去掉最后两个字符
    if code and code[-1].isdigit():
        code = code[:-2]
    # 先映射再查找产品名称
    mapped_code = mapping.get(code, code)
    return product_names.get(mapped_code, product_names.get(code, ''))

# 应用转换
df['com_cd'] = df['com_cd'].apply(transform_com_cd)
df['product_name'] = df['com_cd'].apply(get_product_name)

# 调整列顺序，将product_name放在com_cd后面
cols = df.columns.tolist()
com_cd_index = cols.index('com_cd')
cols.insert(com_cd_index + 1, cols.pop(cols.index('product_name')))
df = df[cols]

# 保存结果
output_path = r'C:\Users\hongl\Desktop\sample\merged_result2.csv'
df.to_csv(output_path, index=False, encoding='utf-8-sig')
print(f"处理完成！已保存到：{output_path}")
