import pandas as pd
import os
from typing import Tuple, List, Dict, Union
import time
import json

class VlookupProcessor:
    def __init__(self, main_file_path: str, lookup_file_path: str, result_path: str):
        self.main_file_path = main_file_path
        self.lookup_file_path = lookup_file_path
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
    def get_sheet_info(file_path: str, sheet_name: str) -> Dict[str, Union[List[str], List[int]]]:
        """获取工作表的列信息"""
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            headers = df.columns.tolist()  # 获取实际的列名
            columns = [chr(65 + i) for i in range(len(headers))]  # A, B, C...
            
            return {
                'headers': headers,
                'columns': columns
            }
        except Exception as e:
            raise Exception(f"读取列信息失败: {str(e)}")

    def process(self, main_sheet: str, main_match_type: str, main_columns: List[str],
                lookup_sheet: str, lookup_match_type: str, lookup_match_columns: List[str],
                return_type: str, return_columns: List[str]) -> Tuple[bool, str]:
        """处理VLOOKUP操作"""
        try:
            print(f"Debug - 处理参数:")
            print(f"主表工作表: {main_sheet}")
            print(f"主表匹配类型: {main_match_type}")
            print(f"主表匹配列: {main_columns}")
            print(f"查找表工作表: {lookup_sheet}")
            print(f"查找表匹配类型: {lookup_match_type}")
            print(f"查找表匹配列: {lookup_match_columns}")
            print(f"返回类型: {return_type}")
            print(f"返回列: {return_columns}")
            
            # 读取Excel文件，不使用第一行作为列名
            main_df = pd.read_excel(self.main_file_path, sheet_name=main_sheet, header=None)
            lookup_df = pd.read_excel(self.lookup_file_path, sheet_name=lookup_sheet, header=None)
            
            print(f"Debug - 表格列数:")
            print(f"主表列数: {len(main_df.columns)}")
            print(f"查找表列数: {len(lookup_df.columns)}")
            
            # 将字母列标识转换为列索引
            def get_column_index(df, col_letter: str, table_name: str) -> int:
                col_idx = ord(col_letter.upper()) - ord('A')
                if col_idx < 0 or col_idx >= len(df.columns):
                    raise ValueError(f"{table_name}中的列标识 {col_letter} 超出范围（表格只有 {len(df.columns)} 列，从A到{chr(ord('A') + len(df.columns) - 1)}）")
                return col_idx
            
            # 转换列标识为索引
            try:
                main_col_indices = [get_column_index(main_df, col, "主表") for col in main_columns]
                lookup_col_indices = [get_column_index(lookup_df, col, "查找表") for col in lookup_match_columns]
                return_col_indices = [get_column_index(lookup_df, col, "查找表") for col in return_columns]
                
                print(f"Debug - 列索引:")
                print(f"主表列索引: {main_col_indices}")
                print(f"查找表匹配列索引: {lookup_col_indices}")
                print(f"返回列索引: {return_col_indices}")
                
            except ValueError as e:
                raise ValueError(str(e))
            
            def standardize_value(value) -> str:
                """标准化数据值"""
                if pd.isna(value):
                    return ''
                
                # 转换为字符串
                value = str(value).strip()
                
                try:
                    # 尝试转换为数字并格式化
                    num = float(value)
                    if num.is_integer():
                        return str(int(num))  # 整数去掉小数点
                    return str(num)  # 保留小数
                except ValueError:
                    # 不是数字，进行文本处理
                    value = value.strip()  # 去除首尾空格
                    value = ' '.join(value.split())  # 合并多个空格
                    return value
            
            def standardize_series(series: pd.Series) -> pd.Series:
                """标准化数据列"""
                return series.apply(standardize_value)
            
            # 创建匹配键
            if main_match_type == 'single':
                main_key = standardize_series(main_df[main_col_indices[0]])
            else:
                # 对多列分别标准化后再合并
                standardized_cols = [standardize_series(main_df[idx]) for idx in main_col_indices]
                main_key = pd.concat(standardized_cols, axis=1).apply(
                    lambda x: '|'.join(x), axis=1
                )
            
            if lookup_match_type == 'single':
                lookup_key = standardize_series(lookup_df[lookup_col_indices[0]])
            else:
                standardized_cols = [standardize_series(lookup_df[idx]) for idx in lookup_col_indices]
                lookup_key = pd.concat(standardized_cols, axis=1).apply(
                    lambda x: '|'.join(x), axis=1
                )
            
            # 创建查找字典
            lookup_dict = {}
            for idx, key in enumerate(lookup_key):
                if return_type == 'single':
                    value = lookup_df.iloc[idx][return_col_indices[0]]
                    if pd.notna(value):
                        lookup_dict[key] = value
                else:
                    values = lookup_df.iloc[idx][return_col_indices].to_dict()
                    if all(pd.notna(v) for v in values.values()):
                        lookup_dict[key] = values
            
            # 执行查找并检查匹配结果
            matched_count = 0
            total_count = len(main_key)
            
            # 记录未匹配的数据
            unmatched_values = set()
            
            if return_type == 'single':
                result_series = main_key.map(lookup_dict)
                matched_mask = result_series.notna()
                matched_count = matched_mask.sum()
                
                # 收集未匹配的值
                if matched_count < total_count:
                    unmatched_values = set(main_key[~matched_mask].unique())
                
                if matched_count == 0:
                    return False, "未找到任何匹配的数据，请检查匹配条件是否正确"
                
                main_df[len(main_df.columns)] = result_series.fillna('')
            else:
                first_result = main_key.map({k: v[return_col_indices[0]] for k, v in lookup_dict.items()})
                matched_mask = first_result.notna()
                matched_count = matched_mask.sum()
                
                # 收集未匹配的值
                if matched_count < total_count:
                    unmatched_values = set(main_key[~matched_mask].unique())
                
                if matched_count == 0:
                    return False, "未找到任何匹配的数据，请检查匹配条件是否正确"
                
                for idx, col_idx in enumerate(return_col_indices):
                    result_series = main_key.map({k: v[col_idx] for k, v in lookup_dict.items()})
                    main_df[len(main_df.columns) + idx] = result_series.fillna('')
            
            # 保存结果
            main_df.to_excel(self.result_path, index=False, header=False)
            
            # 返回匹配统计信息和未匹配的值
            match_rate = (matched_count / total_count) * 100
            result_message = f"匹配完成：共 {total_count} 条数据，成功匹配 {matched_count} 条（{match_rate:.1f}%）"
            
            if unmatched_values:
                # 最多显示5个未匹配的值作为示例
                example_values = list(unmatched_values)[:5]
                example_str = '、'.join(str(v) for v in example_values)
                if len(unmatched_values) > 5:
                    example_str += ' 等'
                result_message += f"\n未匹配的值示例：{example_str}"
            
            return True, result_message
            
        except Exception as e:
            print(f"Error: {str(e)}")
            return False, str(e)

    def cleanup(self):
        """清理临时文件"""
        max_attempts = 3
        delay = 1

        for _ in range(max_attempts):
            try:
                if os.path.exists(self.main_file_path):
                    os.close(os.open(self.main_file_path, os.O_RDONLY))
                    os.remove(self.main_file_path)

                if os.path.exists(self.lookup_file_path):
                    os.close(os.open(self.lookup_file_path, os.O_RDONLY))
                    os.remove(self.lookup_file_path)

                break
            except Exception as e:
                if _ < max_attempts - 1:
                    time.sleep(delay)
                else:
                    print(f"清理文件时出错: {e}")