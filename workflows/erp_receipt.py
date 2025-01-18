import sys
import os

# 将项目根目录添加到Python路径
root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(root_dir)

from services.autogui_processor import AutoGuiProcessor, check_emergency_stop
from utils.log_handler import LogHandler
import time
import pyautogui
import pyperclip
from typing import List, Dict, Any

def shutdown_system(processor: AutoGuiProcessor, logger: LogHandler) -> bool:
    """关闭E10"""
    if not processor.setup_window(processor.ERP_WINDOW):
        logger.error("ERP窗口未打开")
        return False
    pyautogui.hotkey('alt', 'space')
    pyautogui.press('c')
    time.sleep(1)
    pyautogui.press('y')
    if processor.locate_and_click("NO",click=False):
        if processor.locate_and_click("NO",click=True):
            logger.info("已关闭E10")
            return True
        else:
            logger.error("关闭E10失败")
            return False
    else:
        logger.info("已关闭E10")
        return True

def execute_step(processor: AutoGuiProcessor, logger: LogHandler, 
                step_name: str, template: str, window_title: str, 
                click: bool = True) -> bool:
    """执行单个自动化步骤
    
    根据提供的参数执行一个自动化操作步骤,包括定位和点击模板图像。
    
    Args:
        processor: 自动化处理器实例,用于执行具体的自动化操作
        logger: 日志处理器实例,用于记录操作日志
        step_name: 步骤名称,用于日志记录
        template: 模板键名或模板文件路径
        window_title: 需要激活的窗口标题
        click: 是否需要点击模板位置,默认为True
        
    Returns:
        bool: 步骤执行是否成功
              - True: 成功执行步骤
              - False: 执行失败或被用户中断
    """
    # 检查紧急停止
    if check_emergency_stop():
        logger.warning("检测到ESC按键，停止操作")
        return False
        
    # 执行步骤
    action = "点击" if click else "定位"
    if processor.locate_and_click(template, window_title=window_title, click=click):
        logger.info("已成功%s%s", action, step_name)
        return True
    else:
        logger.error("%s%s失败", action, step_name)
        return False

def erp_to_new_receipt(processor: AutoGuiProcessor, logger: LogHandler):
    """执行打开ERP到新建到货单流程"""
    try:
        # 打开ERP
        if not processor.open_erp():
            logger.error("打开ERP失败")
            return False
            
        # 等待ERP加载
        time.sleep(2)
        
        # 执行自动化步骤
        steps = [
            ("IC设计管理系统", "IC_DESIGN_SYSTEM", processor.ERP_WINDOW, True),
            ("维护到货单按钮", "RECEIPT_BUTTON", processor.ERP_WINDOW,True),
            ("新增按钮", "RECEIPT_NEW", processor.RECEIPT_WINDOW, True),
            ("新增到货单界面", "RECEIPT_NEW_MAIN", processor.NEW_RECEIPT_WINDOW, False)
        ]
        
        for step_name, template, window_title, click in steps:
            if not execute_step(processor, logger, step_name, template, window_title, click):
                return False
            time.sleep(3)  # 步骤间隔
        logger.info("到货单处理流程执行完成")
        return True
        
    except Exception as e:
        logger.error("执行到货单处理流程时出错: %s", str(e))
        return False
    
def process_delivery_data(processor: AutoGuiProcessor, logger: LogHandler, 
                         date: str, supplier: str, data: List[Dict[str, Any]]) -> bool:
    """处理送货单数据
    
    Args:
        processor: 自动化处理器
        logger: 日志处理器
        date: 送货日期
        supplier: 供应商名称
        data: 送货单数据列表
        
    Returns:
        bool: 是否成功处理数据
    """
    try:
        # 打开ERP到维护到货单界面
        if not erp_to_new_receipt(processor, logger):
            logger.error("打开ERP失败")
            return False
        
        # 单别
        if not processor.locate_and_click('DOCUMENT_TYPE', click=True):
            logger.error("单别定位失败")
            return False
        
        # 输入单别
        pyperclip.copy("3601")
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(5)
        pyautogui.press('enter')
        time.sleep(2)
        
        if not processor.locate_and_click('RECEIPT_SUPPLY',click=True):
            logger.error("供应商定位失败")
            return False
        # 输入供应商
        pyperclip.copy(supplier)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(5)
        pyautogui.press('enter')
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(1)
        pyautogui.press('enter')

        # 粘贴日期
        pyperclip.copy(date)
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')

        # 输入备注
        if not processor.locate_and_click('RECEIPT_REMARK', click=True):
            logger.error("备注定位失败")
            return False
        time.sleep(1)
        pyperclip.copy(f"{supplier} {date}")
        pyautogui.hotkey('ctrl', 'v')

        # 粘贴订单号
        if not processor.locate_and_click('RECEIPT_RESOURCE_ID', click=False):
            logger.error("订单号定位失败")
            return False
        time.sleep(1)
        pyautogui.click(175, 479)
        time.sleep(1)
        # 复制并粘贴订单号
        order_numbers = "\r\n".join([item["订单号"] for item in data])
        pyperclip.copy(order_numbers)
        # 右键点击
        pyautogui.click(175, 479, button='right')
        time.sleep(1)
        if not processor.locate_and_click('RECEIPT_REGION_PASTE', click=True):
            logger.error("定位粘贴区域失败")
            return False
        time.sleep(30)

        # 粘贴数量
        if not processor.locate_and_click('RECEIPT_BUSINESS_QTY', click=False):
            logger.error("数量定位失败")
            return False
        time.sleep(1)
        pyautogui.click(961, 479)
        time.sleep(1)
        quantities = "\r\n".join([str(item["数量"]) for item in data])
        pyperclip.copy(quantities)
        pyautogui.click(961, 479, button='right')
        # 点击粘贴区域
        if not processor.locate_and_click('RECEIPT_REGION_PASTE', click=True):
            logger.error("定位粘贴区域失败")
            return False
        time.sleep(5)

        while True:
            error_pos = processor.locate_template(processor.get_template_path('RECEIPT_ERROR_1'))
            error2_pos = processor.locate_template(processor.get_template_path('RECEIPT_ERROR_2'))
            if error_pos:
                logger.info("检测到错误,执行删除操作")
                pyautogui.click(24, error_pos[1])
                processor.take_screenshot("locate_receipt_error_1")
                time.sleep(1)
                pyautogui.hotkey('ctrl', 'd')
                time.sleep(1)
                pyautogui.press('y')
                time.sleep(1)
                logger.info(f"{supplier} 行{error_pos[1]}删除完成！")
            elif error2_pos:
                logger.info("检测到错误,执行删除操作")
                pyautogui.click(24, error2_pos[1])
                processor.take_screenshot("locate_receipt_error_2")
                time.sleep(1)
                pyautogui.hotkey('ctrl', 'd')
                time.sleep(1)
                pyautogui.press('y')
                time.sleep(1)
                logger.info(f"{supplier} 行{error2_pos[1]}删除完成！")
            else:
                break

        # 保存
        if not processor.locate_and_click('SAVE', window_title=processor.NEW_RECEIPT_WINDOW):
            logger.error("保存按钮定位失败")
            return False
        time.sleep(2)

        # 若出现警告
        if processor.locate_and_click('WARNING', click=False):
            logger.info("检测到警告窗口")
            # 点击否按钮
            if not processor.locate_and_click('NO'):
                logger.error("否按钮定位失败") 
                return False
            time.sleep(1)

            # 循环处理直到没有警告和错误
            while True:
                # 检查页面是否有警告或者错误
                warning_pos = processor.locate_template(processor.get_template_path('RECEIPT_WARNING_1'))
                warning2_pos = processor.locate_template(processor.get_template_path('RECEIPT_WARNING_2'))
                error_pos = processor.locate_template(processor.get_template_path('RECEIPT_ERROR_1'))
                error2_pos = processor.locate_template(processor.get_template_path('RECEIPT_ERROR_2'))
                
                if warning_pos:
                    logger.info("检测到警告,执行删除操作")
                    pyautogui.click(24, warning_pos[1])
                    processor.take_screenshot("locate_receipt_warning_1")
                    time.sleep(1)
                    pyautogui.hotkey('ctrl', 'd')
                    time.sleep(1)
                    pyautogui.press('y')
                    time.sleep(1)
                    logger.info(f"{supplier} 行{warning_pos[1]}删除完成！")
                elif warning2_pos:
                    logger.info("检测到警告,执行删除操作")
                    pyautogui.click(24, warning2_pos[1])
                    processor.take_screenshot("locate_receipt_warning_2")
                    time.sleep(1)
                    pyautogui.hotkey('ctrl', 'd')
                    time.sleep(1)
                    pyautogui.press('y')
                    time.sleep(1)
                    logger.info(f"{supplier} 行{warning2_pos[1]}删除完成！")
                elif error_pos:
                    logger.info("检测到错误,执行删除操作")
                    pyautogui.click(24, error_pos[1])
                    processor.take_screenshot("locate_receipt_error_1")
                    time.sleep(1)
                    pyautogui.hotkey('ctrl', 'd')
                    time.sleep(1)
                    pyautogui.press('y')
                    time.sleep(1)
                    logger.info(f"{supplier} 行{error_pos[1]}删除完成！")
                elif error2_pos:
                    logger.info("检测到错误,执行删除操作")
                    pyautogui.click(24, error2_pos[1])
                    processor.take_screenshot("locate_receipt_error_2")
                    time.sleep(1)
                    pyautogui.hotkey('ctrl', 'd')
                    time.sleep(1)
                    pyautogui.press('y')
                    time.sleep(1)
                    logger.info(f"{supplier} 行{error2_pos[1]}删除完成！")
                else:
                    break   

        # 点击保存按钮
        if not processor.locate_and_click('SAVE'):
            logger.error("保存按钮定位失败")
            return False
                
        # 点击审核按钮
        if not processor.locate_and_click('AUDIT'):
            logger.error("审核按钮定位失败")
            return False
        
        # 点击是按钮确认审核
        if not processor.locate_and_click('CONFIRM'):
            logger.error("确认按钮定位失败")
            return False
        time.sleep(2)
        
        logger.info(f"{supplier}-{date}送货单处理完成")
        shutdown_system(processor, logger)
        return True
        
    except Exception as e:
        logger.error("处理送货单数据时出错: %s", str(e))
        return False
    
def process_delivery_orders(data_dict: Dict[str, List[Dict[str, Any]]]) -> bool:
    """处理所有送货单数据
    
    Args:
        data_dict: 送货单数据字典，格式为 {日期: [{订单数据}, ...]}
        
    Returns:
        bool: 是否成功处理所有数据
    """
    logger = LogHandler().get_logger('ErpReceipt')
    try:
        logger.info("开始执行自动化录入到货单操作...")
        
        # 初始化自动化处理器
        processor = AutoGuiProcessor(data_dict)
        
        # 处理每个日期的数据
        for date, delivery_data in data_dict.items():
            if not delivery_data:
                continue
                
            supplier = delivery_data[0]["供应商"]  # 假设同一天的数据都来自同一个供应商
            logger.info("开始处理 %s %s 的送货单数据", date, supplier)
            
            if not process_delivery_data(processor, logger, date, supplier, delivery_data):
                return False
        
        logger.info("自动化操作完成")
        return True

    except Exception as e:
        logger.error("程序执行失败: %s", LogHandler.format_error(e))
        shutdown_system(processor, logger)
        return False

if __name__ == "__main__":
    data_dict = {'2025-01-16': 
                 [{
                    '送货日期': '2025-01-16', 
                    '订单号': 'HX-20250101023', 
                   '品名': 'HS9000P-P16', 
                   '晶圆名称': 'HS5122', 
                   '晶圆批号': 'RFEAR9000', 
                   '封装形式': 'SOP16(9.9X3.9X1.4 e=1.27)', 
                   '数量': 50000, 
                   '打印批号': 'C1QG0', 
                   '供应商': '江苏芯丰'}, 
                   {'送货日期': '2025-01-16', 
                    '订单号': 'HX-20241219025', 
                    '品名': 'HS16F3211W', 
                    '晶圆名称': 'HS5130', 
                    '晶圆批号': 'RDKPR4000', 
                    '封装形式': 'SOP16(9.9X3.9X1.4 e=1.27)', 
                    '数量': 72217, 
                    '打印批号': 'F2F5NP-2403', 
                    '供应商': '江苏芯丰'}]}
    process_delivery_orders(data_dict)