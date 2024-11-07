import pandas as pd
import numpy as np
import os
from datetime import datetime
from werkzeug.utils import secure_filename
import chardet
import csv
from io import StringIO
from flask import current_app

class ImportProcessor:
    def __init__(self, upload_folder=None, download_folder=None):
        self.upload_folder = upload_folder or current_app.config['UPLOAD_FOLDER']
        self.download_folder = download_folder or current_app.config['DOWNLOAD_FOLDER']
        os.makedirs(self.upload_folder, exist_ok=True)
        os.makedirs(self.download_folder, exist_ok=True)
    
    def detect_encoding(self, file_path):
        """检测文件编码"""
        try:
            # 首先使用 chardet 检测
            with open(file_path, 'rb') as f:
                raw_data = f.read()
                result = chardet.detect(raw_data)
                detected_enc = result['encoding']
                confidence = result['confidence']
                
                print(f"chardet 检测结果: 编码={detected_enc}, 置信度={confidence}")
                
                # 如果置信度较高，直接返回
                if confidence > 0.8:
                    return detected_enc
            
            # 如果 chardet 检测结果不够理想，尝试常用编码
            encodings_to_try = ['utf-8', 'gbk', 'gb2312', 'gb18030', 'big5']
            for enc in encodings_to_try:
                try:
                    with open(file_path, 'r', encoding=enc) as f:
                        f.read()
                        print(f"成功使用编码: {enc}")
                        return enc
                except UnicodeDecodeError:
                    continue
            
            # 如果都失败了，返回 chardet 的结果
            return detected_enc or 'utf-8'
            
        except Exception as e:
            print(f"编码检测出错: {str(e)}")
            return 'utf-8'  # 默认返回 UTF-8
    
    def read_csv_like_file(self, file_path, encoding=None, delimiter=None):
        """读取类CSV格式的文件，正确处理带引号的字段"""
        try:
            # 尝试不同的编码
            encodings_to_try = ['utf-8', 'gbk', 'gb2312', 'gb18030', 'big5']
            if encoding:
                encodings_to_try.insert(0, encoding)  # 如果指定了编码，优先使用
            
            last_error = None
            for enc in encodings_to_try:
                try:
                    with open(file_path, 'r', encoding=enc) as f:
                        print(f"尝试使用编码: {enc}")
                        # 使用 csv.reader 的高级配置
                        csv.register_dialect('custom',
                            delimiter=delimiter or ',',    # 默认使用逗号分隔符
                            quotechar='"',                # 使用双引号作为引用字符
                            doublequote=True,             # 处理字段中的双引号
                            skipinitialspace=True,        # 跳过分隔符后的空格
                            quoting=csv.QUOTE_MINIMAL     # 仅在必要时使用引号
                        )
                        
                        reader = csv.reader(f, dialect='custom')
                        data = []
                        
                        for row in reader:
                            # 处理每个字段
                            processed_row = []
                            for value in row:
                                # 去除首尾空白，但保留内部空格
                                value = value.strip()
                                
                                # 如果字段包含引号和逗号，保留原始格式
                                if ',' in value and value.startswith('"') and value.endswith('"'):
                                    value = value[1:-1]  # 移除外部引号但保留内容
                                elif value.startswith('"') and value.endswith('"'):
                                    value = value[1:-1]  # 移除不必要的引号
                                    
                                processed_row.append(value)
                            data.append(processed_row)
                        
                        print(f"成功使用编码 {enc} 读取文件")
                        return data, True  # True 表示有表头
                        
                except UnicodeDecodeError as e:
                    print(f"编码 {enc} 失败: {str(e)}")
                    last_error = e
                    continue
                
            # 如果所有编码都失败了
            if last_error:
                raise Exception(f"无法读取文件，所有编码尝试都失败: {str(last_error)}")
                
        except Exception as e:
            print(f"读取文件出错: {str(e)}")
            raise
    
    def preview_data(self, file, source_type, delimiter=None, 
                    start_row=1, header_row=1, auto_split=True):
        """预览数据"""
        try:
            print(f"开始处理文件: {file.filename}")
            
            # 保存上传的文件
            filename = secure_filename(file.filename)
            file_path = os.path.join(self.upload_folder, filename)
            file.save(file_path)
            
            try:
                # 检测编码
                encoding = self.detect_encoding(file_path)
                print(f"检测到的编码: {encoding}")
                
                # 读取文件内容
                data, has_header = self.read_csv_like_file(file_path, encoding, delimiter)
                print(f"读取到 {len(data)} 行数据")
                
                if not data:
                    raise ValueError("文件为空")
                
                # 确保数据非空且有足够的行数
                if len(data) < header_row:
                    raise ValueError(f"数据行数不足: 需要至少 {header_row} 行，实际只有 {len(data)} 行")
                
                # 1. 获取表头（从第 header_row 行获取）
                headers = data[header_row - 1]
                print(f"表头行: {headers}")
                
                # 2. 获取数据行（从表头行之后开始）
                rows = data[header_row:]
                
                # 3. 跳过指定行数
                if start_row > 1:
                    rows = rows[start_row-1:]
                
                # 4. 确保所有行的列数一致
                max_cols = len(headers)
                normalized_rows = []
                for row in rows:
                    # 如果行的列数不足，补充空值
                    if len(row) < max_cols:
                        row = row + [''] * (max_cols - len(row))
                    # 如果行的列数过多，截断
                    elif len(row) > max_cols:
                        row = row[:max_cols]
                    normalized_rows.append(row)
                
                preview_data = {
                    'columns': headers,  # 确保返回表头
                    'rows': normalized_rows[:5]  # 只返回前5行用于预览
                }
                
                print(f"预览数据: {len(preview_data['columns'])} 列, {len(preview_data['rows'])} 行")
                print(f"表头: {preview_data['columns']}")
                print(f"第一行数据: {preview_data['rows'][0] if preview_data['rows'] else []}")
                
                return preview_data
                
            finally:
                # 清理临时文件
                if os.path.exists(file_path):
                    os.remove(file_path)
                    print(f"临时文件已清理: {file_path}")
                    
        except Exception as e:
            print(f"预览错误: {str(e)}")
            raise Exception(f"预览数据失败: {str(e)}")
    
    def import_data(self, files, source_type, column_settings=None, encoding=None, 
                   delimiter=None, sheet=None, start_row=1, header_row=1, auto_split=True):
        """导入数据"""
        try:
            all_data = []
            
            for file in files:
                filename = secure_filename(file.filename)
                file_path = os.path.join(self.upload_folder, filename)
                file.save(file_path)
                
                try:
                    # 检测编码
                    if not encoding:
                        encoding = self.detect_encoding(file_path)
                    
                    # 读取文件内容
                    data, has_header = self.read_csv_like_file(file_path, encoding, delimiter)
                    
                    # 创建DataFrame
                    df = pd.DataFrame(data)
                    
                    # 处理表头
                    if header_row > 0 and len(df) >= header_row:
                        df.columns = df.iloc[header_row-1]
                        df = df.iloc[header_row:]
                    
                    # 跳过指定行数
                    if start_row > 1:
                        df = df.iloc[start_row-1:]
                    
                    # 重置索引
                    df = df.reset_index(drop=True)
                    
                    # 添加到总数据中
                    all_data.append(df)
                
                finally:
                    # 清理临时文件
                    if os.path.exists(file_path):
                        os.remove(file_path)
            
            # 合并所有数据
            final_df = pd.concat(all_data, ignore_index=True)
            
            # 保存为Excel文件
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_filename = f"imported_{timestamp}.xlsx"
            output_path = os.path.join(self.download_folder, output_filename)
            
            # 确保下载目录存在
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False)
                
                # 调整列宽
                worksheet = writer.sheets['Sheet1']
                for idx, col in enumerate(final_df.columns):
                    max_length = max(
                        final_df[col].astype(str).apply(len).max(),
                        len(str(col))
                    ) + 2
                    worksheet.column_dimensions[chr(65 + idx)].width = min(max_length, 50)
            
            return {
                'success': True,
                'filename': output_filename,
                'downloadUrl': f'/download/{output_filename}'
            }
            
        except Exception as e:
            print(f"导入错误: {str(e)}")
            return {
                'success': False,
                'error': str(e)
            }