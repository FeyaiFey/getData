import yaml
import pandas as pd
import numpy as np
from typing import Dict, List, Any, Tuple, Optional
import os
from datetime import datetime
import re
from .log_handler import LogHandler
import json
import shutil

class ExcelProcessor:
    def __init__(self, rules_file: str = 'excel_rules.yaml', log_file: str = 'process_dates.json'):
        """初始化Excel处理器"""
        try:
            # 尝试多个可能的路径
            possible_paths = [
                rules_file,  # 当前目录
                os.path.join('config', rules_file),  # config目录
                os.path.join(os.path.dirname(__file__), '..', rules_file),  # 项目根目录
                os.path.join(os.path.dirname(__file__), '..', 'config', rules_file)  # config目录（相对于utils）
            ]
            
            rules_path = None
            for path in possible_paths:
                if os.path.exists(path):
                    rules_path = path
                    break
                    
            if not rules_path:
                raise FileNotFoundError(f"找不到规则文件，尝试过的路径: {possible_paths}")
                
            self.rules = self._load_rules(rules_path)
            self.log_handler = LogHandler()
            self.process_dates_file = log_file
            print(f"成功加载规则文件: {rules_path}")
            
        except Exception as e:
            print(f"初始化Excel处理器失败: {str(e)}")
            raise

    def _load_rules(self, rules_file: str) -> Dict[str, Any]:
        """加载Excel处理规则"""
        with open(rules_file, 'r', encoding='utf-8') as f:
            config = yaml.safe_load(f)
        return config['rules']

    def _extract_date_from_cell(self, df: pd.DataFrame, rule: Dict[str, Any]) -> Optional[datetime]:
        """从指定单元格提取日期"""
        date_cell = rule.get('date_cell', {})
        if not date_cell:
            print("未配置日期单元格")
            return None

        try:
            row = date_cell.get('row', 0)
            col = ord(date_cell.get('column', 'A').upper()) - ord('A')
            
            # 检查索引是否有效
            if col >= len(df.columns) or row >= len(df):
                print(f"日期单元格位置无效: 行={row}, 列={col}, DataFrame大小: {len(df)}行 x {len(df.columns)}列")
                return None
            
            date_str = str(df.iloc[row, col]).strip()
            print(f"提取到日期字符串: {date_str}")
            
            # 如果配置了正则表达式模式，使用它来提取日期
            pattern = date_cell.get('pattern')
            if pattern:
                print(f"使用正则表达式模式: {pattern}")
                match = re.search(pattern, date_str)
                if match:
                    date_str = match.group(1)
                    print(f"使用正则表达式提取到日期: {date_str}")
                    try:
                        # 处理不同的日期格式
                        if '/' in date_str:
                            date_parts = date_str.split('/')
                            if len(date_parts) == 3:
                                year = int(date_parts[0])
                                if year < 100:
                                    year += 2000
                                return datetime(year, int(date_parts[1]), int(date_parts[2]))
                        elif '-' in date_str:
                            return datetime.strptime(date_str, '%Y-%m-%d')
                    except ValueError as e:
                        print(f"日期格式无效: {date_str}, 错误: {str(e)}")
                        return None
                else:
                    print(f"正则表达式未匹配: {date_str}")
            
            # 如果没有配置正则表达式或匹配失败，尝试提取数字
            date_nums = re.findall(r'\d+', date_str)
            if len(date_nums) >= 3:
                print(f"从数字中提取日期: {date_nums}")
                # 确保年份是4位数
                year = int(date_nums[0])
                if year < 100:
                    year += 2000
                return datetime(year, int(date_nums[1]), int(date_nums[2]))
            return None
            
        except Exception as e:
            print(f"提取日期时出错: {str(e)}")
            return None

    def _extract_date_from_sheet(self, sheet_name: str, rule: Dict[str, Any]) -> Optional[datetime]:
        """从sheet名称提取日期"""
        sheet_format = rule.get('sheet_format', 'date')
        print(f"处理sheet: {sheet_name}, 格式: {sheet_format}")
        
        if sheet_format == 'date':
            try:
                # 移除可能的后缀（如 -2, -3）
                base_name = re.sub(r'-\d+$', '', sheet_name)
                print(f"处理sheet名称: {base_name}")
                # 尝试解析日期
                return datetime.strptime(base_name, '%Y-%m-%d')
            except ValueError as e:
                print(f"解析sheet名称中的日期失败: {str(e)}")
                return None
        elif sheet_format == 'number':
            print(f"数字格式sheet名称: {sheet_name}")
            # 对于类似 "1-2", "1-2-2" 这样的格式，返回None，使用单元格中的日期
            return None
        else:
            return None

    def _should_process_sheet(self, sheet_name: str, rule: Dict[str, Any], rule_name: str) -> bool:
        """检查是否应该处理该sheet"""
        sheet_format = rule.get('sheet_format', 'date')
        
        if sheet_format == 'date':
            # 对于日期格式的sheet名称，检查日期
            sheet_date = self._extract_date_from_sheet(sheet_name, rule)
            if not sheet_date:
                return False
            return not rule.get('check_date', False) or self.log_handler.should_process_date(rule_name, sheet_date)
        elif sheet_format == 'number':
            # 对于数字格式的sheet名称，检查格式是否匹配
            pattern = rule.get('sheet_pattern', r'^\d+-\d+(-\d+)?$')
            if not re.match(pattern, sheet_name):
                return False
                
            # 如果启用了日期检查，检查sheet编号是否在上次处理日期之后
            if rule.get('check_date', False):
                last_date = self.log_handler.get_last_process_date(rule_name)
                if last_date:
                    # 从sheet名称中提取日期编号（例如从"1-9"中提取9）
                    sheet_day = int(sheet_name.split('-')[1].split('-')[0])
                    last_day = last_date.day
                    
                    print(f"检查sheet日期: sheet_day={sheet_day}, last_day={last_day}")
                    # 只处理上次处理日期之后的sheet
                    return sheet_day > last_day
                    
            return True
        elif sheet_format == 'fixed':
            # 对于固定名称的sheet，检查名称是否匹配
            return sheet_name == rule.get('sheet_name', '')
        else:
            return True

    def _find_data_range(self, df: pd.DataFrame, rule: Dict[str, Any]) -> Tuple[int, int]:
        """查找数据的起始和结束行"""
        header_marker = rule.get('header_marker', {})
        footer_marker = rule.get('footer_marker', {})
        
        if not header_marker:
            return 0, len(df)

        # 查找表头行
        header_col = header_marker.get('column', 'A')
        header_value = header_marker.get('value', '')
        header_col_idx = ord(header_col.upper()) - ord('A')
        
        # 获取表尾标记
        footer_col = footer_marker.get('column', 'A')
        footer_value = footer_marker.get('value', '')
        footer_col_idx = ord(footer_col.upper()) - ord('A')
        
        start_row = None
        end_row = None
        
        # 遍历查找表头和表尾
        for idx, row in df.iterrows():
            # 检查表头
            if start_row is None:
                header_cell = str(row[header_col_idx]).strip()
                if header_value in header_cell:
                    start_row = idx + 1
                    continue
            
            # 检查表尾
            if start_row is not None:
                footer_cell = str(row[footer_col_idx]).strip()
                # 如果找到表尾标记或空行
                if (footer_value and footer_value in footer_cell) or \
                   (not footer_value and (pd.isna(footer_cell) or footer_cell == '')):
                    end_row = idx
                    break
        
        # 如果没找到明确的结束行，就用最后一行
        if end_row is None and start_row is not None:
            end_row = len(df)
        
        # 如果没找到表头，使用默认值
        if start_row is None:
            start_row = rule.get('default_start_row', 0)
            end_row = len(df)

        return start_row, end_row

    def _extract_row_data(self, row: pd.Series, fields: List[Dict[str, Any]]) -> Dict[str, Any]:
        """从一行数据中提取字段"""
        item = {}
        for field in fields:
            field_name = field['name']
            
            # 如果是固定值字段
            if 'value' in field:
                item[field_name] = field['value']
                continue
                
            # 如果是空列（如送货日期），跳过
            if not field.get('column'):
                item[field_name] = ''
                continue
                
            # 从Excel列提取数据
            col_index = ord(field['column'].upper()) - ord('A')
            value = row[col_index]
            
            # 处理空值和特殊值
            if pd.isna(value):
                item[field_name] = ''
            else:
                item[field_name] = str(value).strip()
        return item

    def process_excel(self, file_path: str, rule_name: str) -> List[Dict[str, Any]]:
        """处理Excel文件，提取指定字段"""
        if not os.path.exists(file_path):
            print(f"文件或目录已被处理，跳过: {file_path}")
            return []
            
        if rule_name not in self.rules:
            raise ValueError(f"未找到规则: {rule_name}")

        rule = self.rules[rule_name]
        all_results = []
        excel_files_info = []  # 记录Excel文件信息
        
        try:
            # 获取目录下所有的Excel文件
            excel_dir = os.path.dirname(file_path)
            if not os.path.exists(excel_dir):
                print(f"目录已被处理，跳过: {excel_dir}")
                return []
                
            excel_files = [f for f in os.listdir(excel_dir) 
                         if f.lower().endswith(('.xlsx', '.xls'))]
            print(f"找到 {len(excel_files)} 个Excel文件")

            # 如果没有找到Excel文件，直接返回
            if not excel_files:
                print("没有找到Excel文件，跳过处理")
                return []

            # 处理每个Excel文件
            for excel_file in excel_files:
                excel_path = os.path.join(excel_dir, excel_file)
                if not os.path.exists(excel_path):
                    print(f"文件已被处理，跳过: {excel_path}")
                    continue
                    
                print(f"\n处理文件: {excel_file}")
                excel_files_info.append({'path': excel_path, 'xl': None})
                
                try:
                    # 获取所有sheet
                    xl = pd.ExcelFile(excel_path)
                    excel_files_info[-1]['xl'] = xl  # 保存ExcelFile对象以便后续关闭
                    sheet_names = xl.sheet_names
                    print(f"找到的sheet: {sheet_names}")

                    # 处理每个sheet
                    for sheet_name in sheet_names:
                        print(f"\n处理sheet: {sheet_name}")
                        # 检查是否应该处理该sheet
                        if not self._should_process_sheet(sheet_name, rule, rule_name):
                            print(f"跳过sheet: {sheet_name}")
                            continue

                        # 读取sheet
                        df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
                        
                        # 提取日期
                        date = self._extract_date_from_cell(df, rule) or self._extract_date_from_sheet(sheet_name, rule)
                        if date:
                            print(f"提取到日期: {date.strftime('%Y-%m-%d')}")
                        else:
                            print("未找到日期")
                        
                        # 查找数据范围
                        start_row, end_row = self._find_data_range(df, rule)
                        print(f"数据范围: {start_row} - {end_row}")
                        
                        # 提取数据
                        sheet_results = []
                        for idx in range(start_row, end_row):
                            row = df.iloc[idx]
                            # 检查是否是有效的数据行
                            if rule.get('skip_empty', True):
                                # 找到第一个有效的列名（非空且不是固定值的字段）
                                valid_field = None
                                for field in rule['fields']:
                                    if field.get('column') and not field.get('value'):
                                        valid_field = field
                                        break
                                
                                if valid_field:
                                    first_col = ord(valid_field['column'].upper()) - ord('A')
                                    if pd.isna(row[first_col]) or str(row[first_col]).strip() == '':
                                        continue
                            
                            item = self._extract_row_data(row, rule['fields'])
                            if date:
                                item['送货日期'] = date.strftime('%Y-%m-%d')
                            sheet_results.append(item)

                        if sheet_results:
                            print(f"提取到 {len(sheet_results)} 条记录")
                            all_results.extend(sheet_results)
                            # 更新处理日期
                            if date and rule.get('check_date', False):
                                self.log_handler.update_process_date(rule_name, date)
                        else:
                            print("未提取到数据")
                except Exception as e:
                    print(f"处理文件 {excel_file} 时出错: {str(e)}")

            # 如果有数据，按送货日期分组并保存到不同的JSON文件
            if all_results:
                print(f"总共提取到 {len(all_results)} 条记录")
                
                # 获取供应商名称（从规则名称中提取）
                supplier_name = rule.get('supplier_name', rule_name.split('_')[0])
                
                # 按送货日期分组
                date_groups = {}
                for item in all_results:
                    delivery_date = item.get('送货日期')
                    if delivery_date:
                        # 转换日期格式从 YYYY-MM-DD 到 YYYYMMDD
                        date_key = delivery_date.replace('-', '')
                        if date_key not in date_groups:
                            date_groups[date_key] = []
                        date_groups[date_key].append(item)
                
                # 确保目录存在
                json_dir = os.path.join("downloads", "shipping", rule_name)
                os.makedirs(json_dir, exist_ok=True)
                
                # 确保Summary目录存在
                summary_dir = os.path.join("downloads", "shipping", "Summary")
                os.makedirs(summary_dir, exist_ok=True)
                
                # 为每个日期组生成单独的JSON文件
                generated_json_files = []  # 记录生成的JSON文件
                for date_key, group_data in date_groups.items():
                    # 构建JSON文件名
                    json_filename = f"{supplier_name}_{date_key}.json"
                    json_path = os.path.join(json_dir, json_filename)
                    
                    # 保存数据
                    with open(json_path, 'w', encoding='utf-8') as f:
                        json.dump(group_data, f, ensure_ascii=False, indent=2)
                    print(f"成功生成JSON文件: {json_path}，包含 {len(group_data)} 条记录")
                    generated_json_files.append(json_path)
                
                # 移动JSON文件到Summary目录
                for json_file in generated_json_files:
                    target_path = os.path.join(summary_dir, os.path.basename(json_file))
                    try:
                        # 如果目标文件已存在，先删除
                        if os.path.exists(target_path):
                            os.remove(target_path)
                        # 移动文件
                        shutil.move(json_file, target_path)
                        print(f"已移动JSON文件到: {target_path}")
                    except Exception as e:
                        print(f"移动JSON文件时出错: {str(e)}")

            # 关闭所有Excel文件连接
            for file_info in excel_files_info:
                if file_info['xl'] is not None:
                    try:
                        file_info['xl'].close()
                    except:
                        pass

            # 删除整个送货单文件夹
            try:
                if os.path.exists(excel_dir):
                    shutil.rmtree(excel_dir)
                    print(f"已删除文件夹: {excel_dir}")
                    # 直接返回结果，不再继续处理
                    return all_results
            except Exception as e:
                print(f"删除文件夹时出错: {str(e)}")
                # 如果删除文件夹失败，尝试删除单个文件
                for file_info in excel_files_info:
                    try:
                        if os.path.exists(file_info['path']):
                            os.remove(file_info['path'])
                            print(f"已删除文件: {file_info['path']}")
                    except Exception as e:
                        print(f"删除文件时出错: {str(e)}")

            return all_results
            
        except Exception as e:
            print(f"处理Excel文件时出错: {str(e)}")
            import traceback
            print(traceback.format_exc())
            raise

    def get_rule_names(self) -> List[str]:
        """获取所有规则名称"""
        return list(self.rules.keys()) 