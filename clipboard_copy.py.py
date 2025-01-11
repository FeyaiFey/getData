import json
import pyperclip
import os
from typing import List, Dict, Any

def load_json_file(file_path: str) -> List[Dict[str, Any]]:
    """加载JSON文件
    
    Args:
        file_path: JSON文件路径
        
    Returns:
        List[Dict[str, Any]]: JSON数据
        
    Raises:
        FileNotFoundError: 文件不存在
        json.JSONDecodeError: JSON格式错误
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            return json.load(file)
    except FileNotFoundError:
        print(f"错误：找不到文件 '{file_path}'")
        raise
    except json.JSONDecodeError as e:
        print(f"错误：JSON格式错误 - {str(e)}")
        raise
    except Exception as e:
        print(f"错误：读取文件时出错 - {str(e)}")
        raise

def extract_fields(data: List[Dict[str, Any]], fields: List[str]) -> List[str]:
    """从数据中提取多个指定字段
    
    Args:
        data: JSON数据
        fields: 要提取的字段名列表
        
    Returns:
        List[str]: 提取的字段值列表
    """
    result = []
    for entry in data:
        values = [str(entry.get(field, "")) for field in fields]
        result.append(" ".join(values))
    return result

def main():
    """主函数"""
    try:
        # 配置
        json_path = r'D:\PythonProject\getData\downloads\shipping\华宇送货单\处理结果.json'
        fields = ["订单号", "数量"]
        
        # 检查文件是否存在
        if not os.path.exists(json_path):
            print(f"错误：文件 '{json_path}' 不存在")
            return
        
        # 加载数据
        data = load_json_file(json_path)
        if not data:
            print("警告：JSON文件为空")
            return
            
        # 提取字段
        values = extract_fields(data, fields)
        if not values:
            print(f"警告：未找到任何数据")
            return
            
        # 复制到剪贴板
        text = "\n".join(values)
        pyperclip.copy(text)
        print(f"已成功复制 {len(values)} 条记录到剪贴板！")
        print("\n预览前5条记录:")
        for value in values[:5]:
            print(value)
        
    except Exception as e:
        print(f"程序执行出错: {str(e)}")

if __name__ == "__main__":
    main()
