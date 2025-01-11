import re
import yaml
from utils.log_handler import LogHandler

def load_rules(rules_file='email_rules.yaml'):
    """加载规则配置文件"""
    with open(rules_file, 'r', encoding='utf-8') as f:
        config = yaml.safe_load(f)
        print(f"加载到 {len(config['rules'])} 条规则")
        for rule in config['rules']:
            print(f"规则名称: {rule['name']}")
            if 'attachment_name_pattern' in rule:
                print(f"附件匹配模式: {rule['attachment_name_pattern']}")
    return config['rules']

def test_attachment_pattern(filename, rule):
    """测试附件名匹配"""
    print(f"\n测试规则: {rule['name']}")
    print(f"测试文件名: {filename}")
    
    for pattern in rule['attachment_name_pattern']:
        print(f"匹配模式: {pattern}")
        try:
            # 测试完整的字符串匹配
            match_result = re.match(pattern, filename)
            print(f"完整匹配结果: {match_result is not None}")
            if match_result:
                print(f"匹配到的内容: {match_result.group()}")
                print(f"匹配的组: {match_result.groups()}")
            else:
                # 如果匹配失败，打印更多调试信息
                print("匹配失败的可能原因:")
                # 检查文件名长度
                print(f"文件名长度: {len(filename)}")
                # 检查数字部分
                numbers = re.findall(r'\d+', filename)
                print(f"文件名中的数字: {numbers}")
                # 检查文件扩展名
                print(f"文件扩展名: {filename.split('.')[-1]}")
            
        except re.error as e:
            print(f"正则表达式错误: {str(e)}")

def main():
    # 加载规则
    rules = load_rules()
    print("\n已加载规则配置")
    
    # 测试芯丰送货单
    xinf_files = [
        "销售出库单-XSCKD2024120404-314.xlsx"
    ]
    
    # 测试华宇送货单
    huayu_files = [
        "12-06 008  04.xlsx",
        "12-06  008 01.xlsx",
        "12-06  008  03.xlsx"
    ]
    
    # 找到对应的规则
    xinf_rule = next((rule for rule in rules if rule['name'] == "芯丰送货单"), None)
    huayu_rule = next((rule for rule in rules if rule['name'] == "池州华宇送货单"), None)
    
    if xinf_rule:
        print("\n=== 测试芯丰送货单规则 ===")
        for file in xinf_files:
            test_attachment_pattern(file, xinf_rule)
    else:
        print("未找到芯丰送货单规则")
    
    if huayu_rule:
        print("\n=== 测试华宇送货单规则 ===")
        for file in huayu_files:
            test_attachment_pattern(file, huayu_rule)
    else:
        print("未找到华宇送货单规则")

if __name__ == "__main__":
    main() 