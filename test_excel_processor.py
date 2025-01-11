from utils.excel_processor import ExcelProcessor
import json
from datetime import datetime
import os
import yaml

def print_results(results, vendor_name):
    """打印提取结果"""
    print(f"\n{'-'*20} {vendor_name} {'-'*20}")
    for idx, item in enumerate(results, 1):
        print(f"\n记录 {idx}:")
        for key, value in item.items():
            print(f"{key}: {value}")
        print("-" * 50)

def load_email_rules():
    """加载邮件规则配置"""
    with open('email_rules.yaml', 'r', encoding='utf-8') as f:
        return yaml.safe_load(f)['rules']

def get_excel_files(directory):
    """获取目录下的所有Excel文件（排除临时文件）"""
    return [f for f in os.listdir(directory) 
            if (f.endswith('.xlsx') or f.endswith('.xls')) 
            and not f.startswith('~$')]

def main():
    # 初始化处理器
    processor = ExcelProcessor()
    
    try:
        # 加载邮件规则配置
        email_rules = load_email_rules()
        
        # 处理华宇送货单
        for rule in email_rules:
            if rule['name'] == "池州华宇送货单":
                download_path = rule['download_path']
                print(f"\n处理华宇送货单，路径: {download_path}")
                
                # 获取目录下的所有Excel文件
                excel_files = get_excel_files(download_path)
                
                all_results = []  # 存储所有文件的处理结果
                for file in excel_files:
                    file_path = os.path.join(download_path, file)
                    print(f"\n处理文件: {file}")
                    results = processor.process_excel(file_path, '华宇送货单')
                    print_results(results, "华宇送货单")
                    all_results.extend(results)
                    
                # 保存所有结果到一个JSON文件
                if all_results:
                    json_path = os.path.join(download_path, '处理结果.json')
                    with open(json_path, 'w', encoding='utf-8') as f:
                        json.dump(all_results, f, ensure_ascii=False, indent=2)
                    print(f"所有结果已保存到: {json_path}")
                break
                
        # 处理汉旗送货单
        for rule in email_rules:
            if rule['name'] == "汉旗送货单":
                download_path = rule['download_path']
                print(f"\n处理汉旗送货单，路径: {download_path}")
                
                # 获取目录下的所有Excel文件
                excel_files = get_excel_files(download_path)
                
                all_results = []  # 存储所有文件的处理结果
                for file in excel_files:
                    file_path = os.path.join(download_path, file)
                    print(f"\n处理文件: {file}")
                    results = processor.process_excel(file_path, '汉旗送货单')
                    print_results(results, "汉旗送货单")
                    all_results.extend(results)
                    
                # 保存所有结果到一个JSON文件
                if all_results:
                    json_path = os.path.join(download_path, '处理结果.json')
                    with open(json_path, 'w', encoding='utf-8') as f:
                        json.dump(all_results, f, ensure_ascii=False, indent=2)
                    print(f"所有结果已保存到: {json_path}")
                break

        # 显示处理日期记录
        if os.path.exists('process_dates.json'):
            print("\n处理日期记录:")
            with open('process_dates.json', 'r', encoding='utf-8') as f:
                dates = json.load(f)
                for vendor, date in dates.items():
                    print(f"{vendor}: {date}")

    except FileNotFoundError as e:
        print(f"文件不存在: {e}")
    except Exception as e:
        print(f"处理出错: {e}")
        import traceback
        print(traceback.format_exc())

if __name__ == "__main__":
    main()