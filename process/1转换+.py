# -*- coding: utf-8 -*-
"""
Created on Sun Mar 29 11:45:50 2026

@author: hongl
"""

# -*- coding: utf-8 -*-
"""
PDF表格转换工具 - 专为标准表格格式优化
"""
import pdfplumber
import pandas as pd
from pathlib import Path

def pdf_to_csv(pdf_path, csv_path):
    """使用pdfplumber将PDF表格转换为CSV（针对标准表格优化）"""
    try:
        all_dfs = []
        
        # 标准表格提取设置 - 优先使用边框检测
        # 先尝试有边框表格策略
        bordered_settings = {
            "vertical_strategy": "lines",      # 使用线条检测垂直线
            "horizontal_strategy": "lines",    # 使用线条检测水平线
            "snap_tolerance": 3,
            "intersection_tolerance": 3,
        }
        
        # 无边框表格备选策略（用于传真格式）
        borderless_settings = {
            "vertical_strategy": "text",       # 使用文本对齐检测垂直线
            "horizontal_strategy": "text",     # 使用文本对齐检测水平线
            "snap_tolerance": 3,
            "intersection_tolerance": 3,
            "text_tolerance": 3,
        }
        
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages, 1):
                # 首先尝试有边框表格提取
                tables = page.extract_tables(bordered_settings)
                
                # 如果没有找到表格或表格质量差，尝试无边框策略（传真格式）
                if not tables or len(tables) == 0:
                    tables = page.extract_tables(borderless_settings)
                
                for table in tables:
                    if table and len(table) > 1:  # 确保表格至少有两行
                        df = pd.DataFrame(table)
                        
                        # 清洗：去除空值
                        df = df.replace('', None)
                        df = df.replace(r'^\s*$', None, regex=True)
                        
                        # 检查是否为有效表格（至少有一列有数据）
                        if len(df) > 0 and df.iloc[0].notna().any():
                            # 使用第一行作为列名
                            df.columns = df.iloc[0]
                            df = df[1:].reset_index(drop=True)
                            
                            # 再次清洗：移除全为空的行
                            df = df.dropna(how='all')
                            
                            if not df.empty:
                                all_dfs.append(df)
        
        if not all_dfs:
            print(f"  ⚠️ 未找到表格: {pdf_path.name}")
            return False
        
        # 保存结果
        if len(all_dfs) > 1:
            for i, df in enumerate(all_dfs):
                output_file = csv_path.parent / f"{csv_path.stem}_table{i+1}.csv"
                df.to_csv(output_file, index=False, encoding='utf-8-sig')
                print(f"  ✅ 表格{i+1} → {output_file.name}")
        else:
            all_dfs[0].to_csv(csv_path, index=False, encoding='utf-8-sig')
            print(f"  ✅ {csv_path.name}")
        
        return True
        
    except Exception as e:
        print(f"  ❌ 出错: {e}")
        return False


def batch_convert(input_folder, output_folder):
    """批量转换文件夹中的所有PDF文件"""
    input_path = Path(input_folder)
    output_path = Path(output_folder)
    output_path.mkdir(parents=True, exist_ok=True)
    
    # 查找所有PDF文件 - 使用 case_sensitive=False 避免重复
    pdf_files = []
    for ext in ['*.pdf', '*.PDF']:
        pdf_files.extend(input_path.glob(ext))
    
    # 去重（根据文件名小写去重）
    unique_files = {}
    for f in pdf_files:
        key = f.name.lower()
        if key not in unique_files:
            unique_files[key] = f
    
    pdf_files = list(unique_files.values())
    
    # 按文件名排序
    pdf_files.sort(key=lambda x: x.name)
    
    if not pdf_files:
        print(f"❌ 在 {input_folder} 中没有找到PDF文件")
        return
    
    print(f"📁 找到 {len(pdf_files)} 个PDF文件\n")
    
    success = 0
    for i, pdf_file in enumerate(pdf_files, 1):
        print(f"[{i}/{len(pdf_files)}] {pdf_file.name}")
        csv_path = output_path / f"{pdf_file.stem}.csv"
        
        if pdf_to_csv(pdf_file, csv_path):
            success += 1
    
    print(f"\n{'='*40}")
    print(f"✅ 完成: {success}/{len(pdf_files)} 个文件")
    print(f"📂 输出: {output_path}")
    
    # 显示输出文件夹中的CSV文件
    csv_files = list(output_path.glob("*.csv"))
    if csv_files:
        print(f"\n📄 生成的文件:")
        for csv_file in csv_files:
            print(f"   - {csv_file.name}")


# 使用示例
if __name__ == "__main__":
    # 输入文件夹（你的PDF文件所在位置）
    input_folder = r"C:\Users\hongl\Desktop\yuan"
    
    # 输出文件夹
    output_folder = r"C:\Users\hongl\Desktop\新建文件夹"
    
    # 开始转换
    batch_convert(input_folder, output_folder)
