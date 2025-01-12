import cv2
import numpy as np
import pyperclip
import json
import time
import pyautogui
import keyboard
import logging
from pathlib import Path
from datetime import datetime
import win32gui
import win32com.client
import os

class AutoGuiProcessor:
    def __init__(self, json_file):
        self.json_file = json_file
        self.logger = self._setup_logger()
        self.templates = self._load_templates()
        self.orders = self._load_orders()
        self.excel = None
        
    def _setup_logger(self):
        logger = logging.getLogger('AutoGui')
        logger.setLevel(logging.INFO)
        
        # 创建logs目录
        log_dir = Path('logs')
        log_dir.mkdir(exist_ok=True)
        
        # 文件处理器
        log_file = log_dir / f'autogui_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log'
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setLevel(logging.INFO)
        
        # 控制台处理器
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        
        # 设置格式
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        file_handler.setFormatter(formatter)
        console_handler.setFormatter(formatter)
        
        logger.addHandler(file_handler)
        logger.addHandler(console_handler)
        
        return logger

    def _load_templates(self):
        templates = {}
        templates_dir = Path('templates')
        
        if not templates_dir.exists():
            self.logger.error(f"模板目录不存在: {templates_dir}")
            raise FileNotFoundError(f"模板目录不存在: {templates_dir}")
        
        # 加载所需的模板
        template_files = {
            '序号': '序号.png',
            '订单号': '订单号.png',
            '数量': '数量.png'
        }
        
        for name, filename in template_files.items():
            template_path = templates_dir / filename
            if not template_path.exists():
                self.logger.error(f"模板文件不存在: {template_path}")
                raise FileNotFoundError(f"模板文件不存在: {template_path}")
            
            try:
                abs_path = str(template_path.absolute())
                template = cv2.imdecode(np.fromfile(abs_path, dtype=np.uint8), cv2.IMREAD_COLOR)
                
                if template is None:
                    self.logger.error(f"无法读取模板文件内容: {abs_path}")
                    raise FileNotFoundError(f"无法读取模板文件内容: {abs_path}")
                
                templates[name] = template
                self.logger.info(f"成功加载模板: {name}")
                
            except Exception as e:
                self.logger.error(f"加载模板失败: {str(e)}")
                raise
        
        return templates

    def _load_orders(self):
        try:
            with open(self.json_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            self.logger.error(f"加载订单数据失败: {str(e)}")
            raise

    def locate_template(self, template_name):
        try:
            template = self.templates[template_name]
            screenshot = pyautogui.screenshot()
            screenshot_cv = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
            
            # 显示模板图片
            cv2.imshow('Template', template)
            cv2.imshow('Screenshot', screenshot_cv)
            
            result = cv2.matchTemplate(screenshot_cv, template, cv2.TM_CCOEFF_NORMED)
            min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
            
            if max_val < 0.8:  # 匹配度阈值
                self.logger.warning(f"模板 {template_name} 匹配度过低: {max_val}")
                cv2.waitKey(5000)  # 显示5秒
                cv2.destroyAllWindows()
                return None
                
            template_h, template_w = template.shape[:2]
            center_x = max_loc[0] + template_w // 2
            center_y = max_loc[1] + template_h // 2 + 70  # 增加70像素的y轴偏移
            
            # 在截图上绘制匹配位置和点击位置
            cv2.rectangle(screenshot_cv, 
                         max_loc, 
                         (max_loc[0] + template_w, max_loc[1] + template_h), 
                         (0, 255, 0), 
                         2)
            # 在点击位置画一个红色圆点
            cv2.circle(screenshot_cv, 
                      (center_x, center_y), 
                      5, 
                      (0, 0, 255), 
                      -1)
            cv2.imshow('Matched Result', screenshot_cv)
            cv2.waitKey(5000)  # 显示5秒
            cv2.destroyAllWindows()
            
            self.logger.info(f"找到模板 {template_name} 位置: ({center_x}, {center_y}), 匹配度: {max_val:.2f}")
            return (center_x, center_y)
            
        except Exception as e:
            self.logger.error(f"定位模板失败: {str(e)}")
            cv2.destroyAllWindows()
            return None

    def find_last_row(self):
        try:
            # 定位序号列
            serial_loc = self.locate_template('序号')
            if not serial_loc:
                self.logger.error("无法找到序号列位置")
                return None
            
            # 点击序号列
            pyautogui.click(serial_loc[0], serial_loc[1])
            time.sleep(0.5)
            
            # 按Ctrl+End跳转到最后一行
            pyautogui.hotkey('ctrl', 'end')
            time.sleep(0.5)
            
            # 按Ctrl+Up回到最后一个非空行
            pyautogui.hotkey('ctrl', 'up')
            time.sleep(0.5)
            
            # 获取当前位置
            last_row_pos = pyautogui.position()
            
            # 移动到下一行
            pyautogui.moveRel(0, 30)
            next_row_pos = pyautogui.position()
            
            self.logger.info(f"找到最后一行位置: {next_row_pos}")
            return next_row_pos
            
        except Exception as e:
            self.logger.error(f"查找最后一行失败: {str(e)}")
            return None

    def batch_copy_data(self):
        try:
            # 找到最后一行的位置
            last_row_pos = self.find_last_row()
            if not last_row_pos:
                return False
            
            # 定位订单号列和数量列
            order_loc = self.locate_template('订单号')
            quantity_loc = self.locate_template('数量')
            
            if not order_loc or not quantity_loc:
                self.logger.error("无法找到订单号列或数量列位置")
                return False
            
            # 提取订单号和数量
            order_numbers = [order.get("订单号", "") for order in self.orders]
            quantities = [str(order.get("数量", "")) for order in self.orders]
            
            if not order_numbers or not quantities:
                self.logger.warning("未找到订单号或数量数据")
                return False
            
            # 粘贴订单号
            order_text = "\n".join(order_numbers)
            pyperclip.copy(order_text)
            pyautogui.click(order_loc[0], last_row_pos.y)
            time.sleep(0.5)
            pyautogui.hotkey('ctrl', 'v')
            self.logger.info(f"已粘贴 {len(order_numbers)} 个订单号")
            
            # 粘贴数量
            quantity_text = "\n".join(quantities)
            pyperclip.copy(quantity_text)
            pyautogui.click(quantity_loc[0], last_row_pos.y)
            time.sleep(0.5)
            pyautogui.hotkey('ctrl', 'v')
            self.logger.info(f"已粘贴 {len(quantities)} 个数量")
            
            return True
            
        except Exception as e:
            self.logger.error(f"批量复制数据失败: {str(e)}")
            return False

    def check_and_setup_excel(self):
        try:
            # 尝试获取Excel应用程序实例
            self.excel = win32com.client.GetActiveObject('Excel.Application')
            self.logger.info("已找到打开的Excel应用程序")
        except:
            try:
                # 如果Excel未打开，创建新实例
                self.excel = win32com.client.Dispatch('Excel.Application')
                self.logger.info("已创建新的Excel应用程序实例")
            except Exception as e:
                self.logger.error(f"无法启动Excel: {str(e)}")
                return False
        
        try:
            # 使Excel窗口可见
            self.excel.Visible = True
            # 最大化Excel窗口
            self.excel.WindowState = -4137  # xlMaximized
            
            # 将Excel窗口置于最前面
            excel_hwnd = win32gui.FindWindow('XLMAIN', None)
            if excel_hwnd:
                win32gui.SetForegroundWindow(excel_hwnd)
                self.logger.info("Excel窗口已最大化并置于最前面")
            return True
            
        except Exception as e:
            self.logger.error(f"设置Excel窗口失败: {str(e)}")
            return False

def check_emergency_stop():
    return keyboard.is_pressed('esc')

def main():
    try:
        json_file = "downloads/shipping/池州华宇送货单/处理结果.json"
        processor = AutoGuiProcessor(json_file)
        
        # 检查并设置Excel窗口
        if not processor.check_and_setup_excel():
            logging.error("Excel设置失败")
            return
            
        if check_emergency_stop():
            logging.info("检测到紧急停止信号")
            return
            
        success = processor.batch_copy_data()
        if success:
            logging.info("数据批量复制完成")
        else:
            logging.error("数据批量复制失败")
            
    except Exception as e:
        logging.error(f"程序执行失败: {str(e)}")
        raise

if __name__ == "__main__":
    main()