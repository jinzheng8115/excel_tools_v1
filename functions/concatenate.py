import pandas as pd
import os
from typing import Tuple, List

class ConcatenateProcessor:
    def __init__(self, file_path: str, result_path: str):
        self.file_path = file_path
        self.result_path = result_path

    @staticmethod
    def get_sheets(file_path: str) -> List[str]:
        """获取Excel文件中的所有工作表名称"""
        try:
            with pd.ExcelFile(file_path) as xl:
                return xl.sheet_names
        except Exception as e:
            raise Exception(f"读取工作表失败: {str(e)}")

    @staticmethod
    def get_sheet_info(file_path: str, sheet_name: str) -> dict:
        """获取工作表的列信息"""
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
            # 获取第一行作为参考
            headers = [str(x).strip() if pd.notna(x) else '' for x in df.iloc[0].tolist()]
            columns = [chr(65 + i) for i in range(len(headers))]
            
            del df
            return {
                'headers': headers,
                'columns': columns
            }
        except Exception as e:
            raise Exception(f"读取列信息失败: {str(e)}")

    def process(self, sheet: str, columns_data: List[dict]) -> Tuple[bool, str]:
        """
        处理文本合并操作
        参数:
            sheet: 工作表名
            columns_data: 列和分隔符数据列表，格式为[{'column': 'A', 'separator': ' ', 'separator_before': True}, ...]
        """
        df = None
        try:
            # 将列号转换为索引
            def get_column_index(col_letter: str) -> int:
                return ord(col_letter.upper()) - 65

            # 获取所有列索引和分隔符
            columns = [item['column'] for item in columns_data]
            # 如果分隔符为空字符串，则不使用分隔符（不是空格）
            separators = [item.get('separator', None) or '' for item in columns_data]
            separator_before = [item.get('separator_before', False) for item in columns_data]
            col_indices = [get_column_index(col) for col in columns]

            # 读取Excel文件
            df = pd.read_excel(self.file_path, sheet_name=sheet, header=None)

            # 验证列索引是否有效
            if any(idx >= len(df.columns) for idx in col_indices):
                return False, "列号超出范围"

            # 创建用于合并的数据副本
            df_merge = df.copy()

            # 处理每一列的数据
            for idx in col_indices:
                try:
                    numeric_data = pd.to_numeric(df_merge[idx], errors='coerce')
                    if numeric_data.notna().all():
                        df_merge[idx] = numeric_data.astype(int).astype(str)
                    else:
                        df_merge[idx] = (df_merge[idx]
                                       .fillna('')
                                       .astype(str)
                                       .str.strip()
                                       .replace('nan', ''))
                except:
                    df_merge[idx] = (df_merge[idx]
                                   .fillna('')
                                   .astype(str)
                                   .str.strip()
                                   .replace('nan', ''))

            # 合并选定的列
            def merge_columns(row):
                result = []
                
                # 处理所有列和分隔符
                for i, idx in enumerate(col_indices):
                    val = str(row[idx]).strip()
                    if val and val != 'nan':
                        # 如果设置了在值之前添加分隔符且不是第一个值
                        if separator_before[i] and separators[i]:
                            result.append(separators[i])
                        
                        result.append(val)
                        
                        # 如果没有设置在值之前添加分隔符，则在值之后添加
                        if not separator_before[i] and separators[i]:
                            result.append(separators[i])
                
                return ''.join(result)

            # 创建结果DataFrame
            result = df.copy()
            result[len(result.columns)] = df_merge.apply(merge_columns, axis=1)

            # 保存结果
            result.to_excel(self.result_path, index=False, header=False)

            return True, ""
            
        except Exception as e:
            if df is not None:
                del df
            return False, str(e)

    def cleanup(self):
        """清理临时文件"""
        try:
            if os.path.exists(self.file_path):
                os.close(os.open(self.file_path, os.O_RDONLY))
                os.remove(self.file_path)
        except Exception as e:
            print(f"清理文件时出错: {e}")