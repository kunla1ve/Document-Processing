# -*- coding: utf-8 -*-
"""
Created on Tue Mar 31 20:31:55 2026

@author: kunlave
"""



import pandas as pd
import os
import glob
import time

# 设置路径
input_folder = r'C:\Users\kunlave\Desktop\新增資料夾'
output_file = r'C:\Users\kunlave\Desktop\sample.csv'

# 定义列名映射（处理大小写和不一致的情况）
column_mapping = {
    'client_no': 'Client_no',
    'Client_no': 'Client_no',
    'com_type': 'com_type',
    'exch_cd': 'exch_cd',
    'com_cd': 'com_cd',
    'contract_date': 'contract_date',
    'counter_party': 'counter_party',
    'traded_strike_price': 'traded_strike_price',
    'call_put': 'call_put',
    'closeout_qty': 'closeout_qty',
    'closedout_qty': 'closeout_qty',  # 处理拼写不一致
    'closeout_qty': 'closeout_qty'
}

# 最终需要的列（按顺序）
final_columns = ['Client_no', 'com_type', 'exch_cd', 'com_cd', 'contract_date', 
                 'counter_party', 'traded_strike_price', 'call_put', 'closeout_qty', '文件来源']

# 获取所有CSV文件
csv_files = glob.glob(os.path.join(input_folder, '*.csv'))

if not csv_files:
    print(f"在 {input_folder} 中没有找到CSV文件")
else:
    # 创建列表存储所有数据
    all_data_list = []
    file_stats = {}
    
    # 逐个读取并合并CSV文件
    for file in csv_files:
        try:
            # 获取文件名（不含路径）
            file_name = os.path.basename(file)
            
            # 读取CSV文件
            df = pd.read_csv(file)
            
            # 重命名列（统一列名）
            df.rename(columns=column_mapping, inplace=True)
            
            # 检查是否包含所需的列（不包括文件来源）
            required_cols = final_columns[:-1]
            existing_cols = [col for col in required_cols if col in df.columns]
            
            if len(existing_cols) == len(required_cols):
                # 只保留需要的列
                df_filtered = df[required_cols].copy()
                
                # 添加文件来源列
                df_filtered['文件来源'] = file_name
                
                # 添加到列表
                all_data_list.append(df_filtered)
                file_stats[file_name] = len(df_filtered)
                print(f"✓ 已读取: {file_name} - {len(df_filtered)} 行")
            else:
                missing_cols = set(required_cols) - set(existing_cols)
                print(f"✗ 跳过 {file_name}: 缺少列 {missing_cols}")
                print(f"  文件中的列: {list(df.columns)}")
                
        except Exception as e:
            print(f"✗ 读取 {os.path.basename(file)} 时出错: {e}")
    
    # 合并所有数据
    if all_data_list:
        all_data = pd.concat(all_data_list, ignore_index=True)
        
        # 按Client_no排序
        all_data = all_data.sort_values('Client_no').reset_index(drop=True)
        
        # 确保输出目录存在
        output_dir = os.path.dirname(output_file)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        # 如果文件已存在，尝试关闭可能打开的文件
        if os.path.exists(output_file):
            try:
                os.remove(output_file)
                time.sleep(0.5)  # 等待文件系统释放
            except PermissionError:
                print(f"\n警告: 文件 {output_file} 正在被使用中")
                # 尝试使用不同的文件名
                output_file = output_file.replace('.csv', '_new.csv')
                print(f"将使用新文件名: {output_file}")
        
        # 保存合并后的数据
        all_data.to_csv(output_file, index=False, encoding='utf-8-sig')
        
        print(f"\n{'='*50}")
        print(f"合并完成！")
        print(f"总共处理了 {len(csv_files)} 个文件")
        print(f"成功读取 {len(file_stats)} 个文件")
        print(f"合并后共 {len(all_data)} 行数据")
        print(f"输出文件: {output_file}")
        


    else:
        print("\n没有成功读取任何文件，请检查文件格式")




