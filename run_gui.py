from services.autogui import AutoGuiProcessor, check_emergency_stop
from utils.log_handler import LogHandler
from typing import List, Tuple, Dict, Any
import pyautogui
import time
import pyperclip
from datetime import datetime
from pathlib import Path

def execute_step(processor: AutoGuiProcessor, logger: LogHandler, 
                step_name: str, template_name: str, window_title: str, 
                click: bool = True) -> bool:
    """执行单个自动化步骤
    
    Args:
        processor: 自动化处理器
        logger: 日志处理器
        step_name: 步骤名称
        template_name: 模板名称
        window_title: 窗口标题
        click: 是否需要点击
        
    Returns:
        bool: 步骤是否执行成功
    """
    # 检查紧急停止
    if check_emergency_stop():
        logger.warning("检测到ESC按键，停止操作")
        return False
        
    # 执行步骤
    action = "点击" if click else "定位"
    if processor.locate_and_click_template(template_name, window_title=window_title, click=click):
        logger.info("已成功%s%s", action, step_name)
        return True
    else:
        logger.error("%s%s失败", action, step_name)
        return False

def get_automation_steps() -> List[Tuple[str, str, str, bool]]:
    """获取自动化步骤列表
    
    Returns:
        List[Tuple[str, str, str, bool]]: 步骤列表，每个步骤包含(步骤名称, 模板名称, 窗口标题, 是否点击)
    """
    return [
        ("IC设计管理系统", "index", AutoGuiProcessor.ERP_WINDOW, True),
        ("IC设计管理系统", "ic_design_system", AutoGuiProcessor.ERP_WINDOW, False),
        ("维护到货单按钮", "receipt_button", AutoGuiProcessor.ERP_WINDOW, True),
        ("维护到货单界面", "receipt_main", AutoGuiProcessor.RECEIPT_WINDOW, False),
        ("新增按钮", "receipt_new", AutoGuiProcessor.RECEIPT_WINDOW, True),
        ("新增到货单界面", "receipt_new_main", AutoGuiProcessor.NEW_RECEIPT_WINDOW, False)
    ]

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
        pyautogui.press('enter')
        pyperclip.copy("3601")
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(1)
        pyautogui.press('enter')
        if not processor.locate_and_click_template("receipt_supply"):
            logger.error("定位供应商输入框失败")
            return False
        pyperclip.copy(supplier)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(1)
        pyperclip.copy(date)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(1)
        pyautogui.press('enter')

        
        # 定位并输入备注
        if not processor.locate_and_click_template("receipt_remark"):
            logger.error("定位备注输入框失败")
            return False
        time.sleep(1)
        pyperclip.copy(f"{supplier} {date}")
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(1)
        
        # 处理订单号
        if processor.locate_template("receipt_resource_id"):
            # 点击订单号输入区域
            pyautogui.click(175, 479)
            time.sleep(1)
            # 复制并粘贴订单号
            order_numbers = "\r\n".join([item["订单号"] for item in data])
            pyperclip.copy(order_numbers)
            # 右键点击
            print(f"订单号: {order_numbers}")
            pyautogui.click(175, 479, button='right')
            # 点击粘贴区域
            if not processor.locate_and_click_template("receipt_region_paste",wait_time=5):
                logger.error("定位粘贴区域失败")
                return False
        else:
            logger.error("定位订单号模板失败")
            return False
  
        # 处理数量
        if processor.locate_template("receipt_businessQty"):
            # 点击数量输入区域
            pyautogui.click(961, 479)
            # 复制并粘贴数量
            quantities = "\r\n".join([str(item["数量"]) for item in data])
            pyperclip.copy(quantities)
            # 右键点击
            pyautogui.click(961, 479, button='right')
            # 点击粘贴区域
            if not processor.locate_and_click_template("receipt_region_paste"):
                logger.error("定位粘贴区域失败")
                return False
            time.sleep(5)
        else:
            logger.error("定位数量模板失败")
            return False
        
        # 确认是否报错
        while processor.locate_and_click_template("receipt_error",click=False):
            # 截图并保存
            try:
                # 获取当前时间作为文件名
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                # 创建screenshots目录(如果不存在)
                screenshot_dir = Path("screenshots")
                screenshot_dir.mkdir(exist_ok=True)
                # 保存截图
                screenshot_path = screenshot_dir / f"error_{timestamp}.png"
                screenshot = pyautogui.screenshot()
                screenshot.save(screenshot_path)
                logger.warning("检测到报错窗口,已保存截图: %s", screenshot_path)
                center_x,center_y = processor.locate_template("receipt_error")
                pyautogui.click(23,center_y)
                time.sleep(1)
                pyautogui.hotkey('ctrl', 'd')
                time.sleep(1)
                if processor.locate_and_click_template("yes"):
                    logger.info("已删除此行")
                else:
                    logger.error("删除失败")
            except Exception as e:
                logger.error("保存报错截图失败: %s", LogHandler.format_error(e))
        
        # 点击保存
        if processor.locate_and_click_template("save"):
            logger.info("已保存")
        else:
            logger.error("保存失败")

        # 点击审核
        if processor.locate_and_click_template("audit"):
            logger.info("已审核")
        else:
            logger.error("审核失败")
        time.sleep(1)
        pyautogui.hotkey('alt', 'f4')

        # 关闭维护到货单窗口
        if processor._setup_window(processor.RECEIPT_WINDOW):
            pyautogui.hotkey('alt', 'f4')
            logger.info("已关闭维护到货单窗口")
            time.sleep(1)
        else:
            logger.warning("未找到维护到货单窗口")

        # 关闭E10
        if processor._setup_window(processor.ERP_WINDOW):
            pyautogui.hotkey('alt', 'f4')
            logger.info("已关闭E10")
            time.sleep(1)
        else:
            logger.warning("未找到E10")
        pyautogui.press('enter')

        return True

    except Exception as e:
        logger.error("处理送货单数据失败: %s", LogHandler.format_error(e))
        return False

def process_delivery_orders(data_dict: Dict[str, List[Dict[str, Any]]]) -> bool:
    """处理所有送货单数据
    
    Args:
        data_dict: 送货单数据字典，格式为 {日期: [{订单数据}, ...]}
        
    Returns:
        bool: 是否成功处理所有数据
    """
    logger = LogHandler().get_logger('RunGui')
    try:
        logger.info("开始执行自动化操作...")
        
        # 初始化自动化处理器
        processor = AutoGuiProcessor(data_dict)
        
        # 检查并设置digiwin窗口
        if not processor.check_and_setup_digiwin():
            logger.error("鼎捷ERP设置失败")
            return False
            
        # 执行初始化步骤
        steps = get_automation_steps()
        for step_name, template_name, window_title, click in steps:
            if not execute_step(processor, logger, step_name, template_name, window_title, click):
                return False
        
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
        return False

if __name__ == "__main__":
    # 测试数据
    test_data = {
        "2025-01-15": [
            {"订单号": "HX-20250107012", "数量": 100, "供应商": "池州华宇"},
            {"订单号": "HX-20250107013", "数量": 200, "供应商": "池州华宇"}
        ]
    }
    process_delivery_orders(test_data) 