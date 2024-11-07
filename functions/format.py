import pandas as pd
import os
from datetime import datetime
import zipfile
import io
from werkzeug.utils import secure_filename
import shutil
import gc
import time

class ExcelProcessor:
    def __init__(self):
        """
        初始化 ExcelProcessor
        不再需要传递参数，而是从 config 中获取
        """
        from config import UPLOAD_FOLDER, RESULT_FOLDER
        
        self.upload_folder = UPLOAD_FOLDER
        self.result_folder = RESULT_FOLDER
        
        # 确保目录存在
        os.makedirs(self.upload_folder, exist_ok=True)
        os.makedirs(self.result_folder, exist_ok=True)
        
    def split_excel(self, file, split_all=True, selected_sheets=None):
        """
        拆分Excel文件
        :param file: 上传的文件对象
        :param split_all: 是否拆分所有工作表
        :param selected_sheets: 选择要拆分的工作表列表
        :return: 拆分后的文件信息列表
        """
        try:
            # 保存上传的文件并获取原始文件名
            original_filename = os.path.splitext(file.filename)[0]  # 获取不带扩展名的原文件名
            
            # 只对临时保存的文件使用secure_filename
            temp_filename = secure_filename(file.filename)
            file_path = os.path.join(self.upload_folder, temp_filename)
            file.save(file_path)
            
            result_files = []
            used_filenames = set()  # 用于跟踪已使用的文件名
            
            try:
                with pd.ExcelFile(file_path) as excel_file:
                    sheets_to_split = excel_file.sheet_names if split_all else selected_sheets
                    
                    for sheet_name in sheets_to_split:
                        try:
                            # 读取工作表
                            df = pd.read_excel(file_path, sheet_name=sheet_name)
                            
                            # 生成安全的文件名（用于存储）
                            safe_base_name = secure_filename(original_filename)
                            safe_sheet_name = secure_filename(sheet_name)
                            safe_output_filename = f"{safe_base_name}_{safe_sheet_name}.xlsx"
                            
                            # 处理文件名冲突
                            counter = 1
                            while safe_output_filename in used_filenames:
                                safe_output_filename = f"{safe_base_name}_{safe_sheet_name}_{counter}.xlsx"
                                counter += 1
                            
                            used_filenames.add(safe_output_filename)
                            output_path = os.path.join(self.result_folder, safe_output_filename)
                            
                            # 使用 ExcelWriter 保存并调整列宽
                            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                                df.to_excel(writer, sheet_name=sheet_name, index=False)
                                
                                # 调整列宽
                                worksheet = writer.sheets[sheet_name]
                                for idx, col in enumerate(df.columns):
                                    max_length = max(
                                        df[col].astype(str).apply(len).max(),
                                        len(str(col))
                                    ) + 2
                                    worksheet.column_dimensions[chr(65 + idx)].width = min(max_length, 50)
                            
                            result_files.append({
                                'filename': safe_output_filename,
                                'originalName': f"{original_filename}_{sheet_name}.xlsx",
                                'downloadUrl': f'/download/{safe_output_filename}?original_name={original_filename}_{sheet_name}.xlsx'
                            })
                            
                        except Exception as e:
                            print(f"处理工作表 {sheet_name} 时出错: {str(e)}")
                            continue
                            
            finally:
                # 清理临时文件
                try:
                    if os.path.exists(file_path):
                        os.remove(file_path)
                except Exception as e:
                    print(f"清理临时文件失败: {str(e)}")
            
            if not result_files:
                raise Exception("没有成功拆分任何工作表")
            
            return result_files
            
        except Exception as e:
            raise Exception(f"拆分Excel文件失败: {str(e)}")
        
    def merge_excel(self, files, merge_all=True, add_source=False, output_filename='merged_file', merge_mode='sheets'):
        """合并Excel文件"""
        temp_files = []
        
        try:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_filename = f"{secure_filename(output_filename)}_{timestamp}.xlsx"
            output_path = os.path.join(self.result_folder, output_filename)
            
            print(f"合并模式: {merge_mode}")
            print(f"添加来源列: {add_source}")
            
            add_source = True
            
            if merge_mode == 'by-name':
                # 按工作表名合并
                sheet_data = {}  # 用于存储各个工作表的数据
                
                for file in files:
                    if file.filename == '':
                        continue
                    
                    temp_filename = secure_filename(file.filename)
                    temp_path = os.path.join(self.upload_folder, temp_filename)
                    temp_files.append(temp_path)
                    file.save(temp_path)
                    
                    try:
                        with pd.ExcelFile(temp_path) as excel_file:
                            sheets_to_merge = excel_file.sheet_names if merge_all else [excel_file.sheet_names[0]]
                            
                            for sheet_name in sheets_to_merge:
                                try:
                                    # 读取工作表
                                    df = pd.read_excel(temp_path, sheet_name=sheet_name)
                                    
                                    # 总是添加来源信息列
                                    df['来源文件'] = file.filename
                                    df['来源工作表'] = sheet_name
                                    
                                    # 将所有列名转换为字符串类型
                                    df.columns = df.columns.astype(str)
                                    
                                    if sheet_name not in sheet_data:
                                        # 第一次遇到这个工作表名
                                        sheet_data[sheet_name] = {
                                            'columns': list(df.columns),
                                            'data': [df]
                                        }
                                    else:
                                        # 检查列数是否相同（不包括来源列）
                                        base_columns = [col for col in sheet_data[sheet_name]['columns'] 
                                                      if col not in ['来源文件', '来源工作表']]
                                        current_columns = [col for col in df.columns 
                                                      if col not in ['来源文件', '来源工作表']]
                                        
                                        if len(base_columns) == len(current_columns):
                                            # 使用第一个文件的列名顺序
                                            df = df[sheet_data[sheet_name]['columns']]
                                            sheet_data[sheet_name]['data'].append(df)
                                        else:
                                            print(f"工作表 {sheet_name} 在文件 {file.filename} 中的列数不匹配")
                                            print(f"预期列: {base_columns}")
                                            print(f"实际列: {current_columns}")
                                except Exception as e:
                                    print(f"处理工作表 {sheet_name} 时出错: {str(e)}")
                                    continue
                    except Exception as e:
                        print(f"处理文件 {file.filename} 时出错: {str(e)}")
                        continue
                
                # 合并并保存数据
                with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                    for sheet_name, sheet_info in sheet_data.items():
                        if sheet_info['data']:  # 确保有数据要合并
                            try:
                                merged_df = pd.concat(sheet_info['data'], ignore_index=True)
                                
                                # 如果不需要来源信息，则删除这些列
                                if not add_source:
                                    merged_df = merged_df.drop(['来源文件', '来源工作表'], axis=1)
                                
                                merged_df.to_excel(writer, sheet_name=sheet_name, index=False)
                                
                                # 调整列宽
                                worksheet = writer.sheets[sheet_name]
                                for idx, col in enumerate(merged_df.columns):
                                    max_length = max(
                                        merged_df[col].astype(str).apply(len).max(),
                                        len(str(col))
                                    ) + 2
                                    worksheet.column_dimensions[chr(65 + idx)].width = min(max_length, 50)
                            except Exception as e:
                                print(f"合并工作表 {sheet_name} 时出错: {str(e)}")
                                continue
                            finally:
                                if 'merged_df' in locals():
                                    del merged_df
                                    gc.collect()
                
                return {
                    'filename': output_filename,
                    'downloadUrl': f'/download/{output_filename}'
                }
            
            elif merge_mode == 'single':
                # 合并到单个工作表
                all_data = []
                first_df = None  # 用于存储第一个数据框的结构
                
                for file in files:
                    if file.filename == '':
                        continue
                    
                    # 保存上传的文件
                    temp_filename = secure_filename(file.filename)
                    temp_path = os.path.join(self.upload_folder, temp_filename)
                    temp_files.append(temp_path)
                    file.save(temp_path)
                    
                    try:
                        # 使用with语句确保Excel文件正确关闭
                        with pd.ExcelFile(temp_path) as excel_file:
                            sheets_to_merge = excel_file.sheet_names if merge_all else [excel_file.sheet_names[0]]
                            
                            for sheet_name in sheets_to_merge:
                                try:
                                    # 读取工作表
                                    df = pd.read_excel(temp_path, sheet_name=sheet_name)
                                    
                                    # 添加来源信息列（无论是否选择add_source，都添加这些信息）
                                    df['来源文件'] = file.filename
                                    df['来源工作表'] = sheet_name
                                    
                                    # 如果是第一个数据框，保存其结构
                                    if first_df is None:
                                        first_df = df
                                        all_data.append(df)
                                    else:
                                        # 检查列结构是否匹配（不包括来源列）
                                        if len(df.columns) == len(first_df.columns):  # 只检查列数是否相同
                                            # 将所有列名转换为字符串
                                            df.columns = df.columns.astype(str)
                                            # 使用第一个数据框的列名
                                            df.columns = first_df.columns
                                            all_data.append(df)
                                        else:
                                            print(f"工作表 {sheet_name} 在文件 {file.filename} 中的列数不匹配")
                                            print(f"预期列数: {len(first_df.columns)}")
                                            print(f"实际列数: {len(df.columns)}")
                                except Exception as e:
                                    print(f"处理工作表 {sheet_name} 时出错: {str(e)}")
                                    continue
                            
                        # 确保文件句柄已关闭
                        gc.collect()
                    
                    except Exception as e:
                        print(f"处理文件 {file.filename} 时出错: {str(e)}")
                        continue
                
                # 合并所有数据
                if not all_data:
                    raise Exception("没有可合并的数据，可能是由于列结构不匹配")
                
                merged_df = pd.concat(all_data, ignore_index=True)
                
                # 如果不需要来源信息，则删除这些列
                if not add_source:
                    merged_df = merged_df.drop(['来源文件', '来源工作表'], axis=1)
                
                # 保存合并后的文件
                with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                    merged_df.to_excel(writer, sheet_name='合并数据', index=False)
                    
                    # 调整列宽
                    worksheet = writer.sheets['合并数据']
                    for idx, col in enumerate(merged_df.columns):
                        max_length = max(
                            merged_df[col].astype(str).apply(len).max(),
                            len(str(col))
                        ) + 2
                        worksheet.column_dimensions[chr(65 + idx)].width = min(max_length, 50)
                
                # 清理内存
                del merged_df
                gc.collect()
                
                # 确保文件已经保存
                time.sleep(0.5)
                
                if not os.path.exists(output_path):
                    raise Exception("文件保存失败")
                
                return {
                    'filename': output_filename,
                    'downloadUrl': f'/download/{output_filename}'
                }
            
            else:
                # 保持在不同工作表
                with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                    for file in files:
                        if file.filename == '':
                            continue
                        
                        temp_filename = secure_filename(file.filename)
                        temp_path = os.path.join(self.upload_folder, temp_filename)
                        temp_files.append(temp_path)
                        file.save(temp_path)
                        
                        try:
                            with pd.ExcelFile(temp_path) as excel_file:
                                sheets_to_merge = excel_file.sheet_names if merge_all else [excel_file.sheet_names[0]]
                                
                                for sheet_name in sheets_to_merge:
                                    df = pd.read_excel(temp_path, sheet_name=sheet_name)
                                    
                                    # 添加来源信息（如果需要）
                                    if add_source:
                                        df['来源文件'] = file.filename
                                        df['来源工作表'] = sheet_name
                                    
                                    new_sheet_name = f"{os.path.splitext(file.filename)[0]}_{sheet_name}"
                                    if len(new_sheet_name) > 31:
                                        new_sheet_name = new_sheet_name[:31]
                                    
                                    df.to_excel(writer, sheet_name=new_sheet_name, index=False)
                                    
                                    # 调整列宽
                                    worksheet = writer.sheets[new_sheet_name]
                                    for idx, col in enumerate(df.columns):
                                        max_length = max(
                                            df[col].astype(str).apply(len).max(),
                                            len(str(col))
                                        ) + 2
                                        worksheet.column_dimensions[chr(65 + idx)].width = min(max_length, 50)
                                    
                                    # 清理内存
                                    del df
                                    gc.collect()
                            
                        except Exception as e:
                            print(f"处理文件 {file.filename} 时出错: {str(e)}")
                            continue
                
                # 确保所有文件句柄都已关闭
                gc.collect()
                
                # 等待一小段时间确保文件已完全关闭
                time.sleep(0.5)
                
                # 清理临时文件
                for temp_path in temp_files:
                    try:
                        if os.path.exists(temp_path):
                            os.remove(temp_path)
                    except Exception as e:
                        print(f"清理临时文件失败: {str(e)}")
                        # 如果删除失败，将在下一次请求时重试
                
                return {
                    'filename': output_filename,
                    'downloadUrl': f'/download/{output_filename}'
                }
            
        except Exception as e:
            if 'output_path' in locals() and os.path.exists(output_path):
                try:
                    os.remove(output_path)
                except Exception as cleanup_error:
                    print(f"清理输出文件失败: {str(cleanup_error)}")
            raise Exception(f"合并Excel文件失败: {str(e)}")
            
        finally:
            # 清理所有临时文件
            for temp_path in temp_files:
                try:
                    if os.path.exists(temp_path):
                        os.remove(temp_path)
                except Exception as e:
                    print(f"清理临时文件失败: {str(e)}")

    def get_sheet_names(self, file):
        """获取Excel文件的工作表名列表"""
        file_path = None
        try:
            # 保存上传的文件
            filename = secure_filename(file.filename)
            file_path = os.path.join(self.upload_folder, filename)
            file.save(file_path)
            
            # 使用 with 语句读取工作表名
            with pd.ExcelFile(file_path) as excel_file:
                return excel_file.sheet_names
            
        except Exception as e:
            raise Exception(f"获取工作表名失败: {str(e)}")
        
        finally:
            # 清理临时文件
            if file_path and os.path.exists(file_path):
                try:
                    os.remove(file_path)
                except Exception as e:
                    print(f"清理临时文件失败: {str(e)}")