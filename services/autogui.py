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
import win32con
import sys

# 添加项目根目录到Python路径
current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(current_dir)
if project_root not in sys.path:
    sys.path.insert(0, project_root)

from utils.window_finder import find_window_by_title, start_and_find_window

class AutoGuiProcessor:
    def __init__(self, json_file):
        self.json_file = json_file
        self.logger = self._setup_logger()
        self.templates = self._load_templates()
        
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
            'index': 'index.png',
            'ic_design_system': 'ic_design_system.png',
            'receipt_button': 'receipt_button.png',
            'receipt_main': 'receipt_main.png'
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

    def locate_template(self, template_name):
        try:
            template = self.templates[template_name]
            screenshot = pyautogui.screenshot()
            screenshot_cv = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
            
            result = cv2.matchTemplate(screenshot_cv, template, cv2.TM_CCOEFF_NORMED)
            min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
            
            if max_val < 0.8:  # 匹配度阈值
                self.logger.warning(f"模板 {template_name} 匹配度过低: {max_val}")
                cv2.waitKey(2000)  # 显示2秒
                cv2.destroyAllWindows()
                return None
                
            template_h, template_w = template.shape[:2]
            center_x = max_loc[0] + template_w // 2
            center_y = max_loc[1] + template_h // 2
            
            # 获取原始截图尺寸
            height, width = screenshot_cv.shape[:2]
            # 计算缩放比例
            scale = 1440 / width
            new_width = 1440
            new_height = int(height * scale)
            # 缩放截图
            display_img = cv2.resize(screenshot_cv, (new_width, new_height))
            
            # 在缩放后的截图上绘制匹配位置和点击位置
            scaled_max_loc = (int(max_loc[0] * scale), int(max_loc[1] * scale))
            scaled_template_w = int(template_w * scale)
            scaled_template_h = int(template_h * scale)
            # 计算点击位置(匹配位置的中心点)
            scaled_center_x = scaled_max_loc[0] + scaled_template_w // 2
            scaled_center_y = scaled_max_loc[1] + scaled_template_h // 2
            
            # cv2.rectangle(display_img,
            #              scaled_max_loc,
            #              (scaled_max_loc[0] + scaled_template_w, scaled_max_loc[1] + scaled_template_h),
            #              (0, 255, 0),
            #              2)
            # # 在点击位置画一个红色圆点
            # cv2.circle(display_img,
            #           (int(scaled_center_x), int(scaled_center_y)),
            #           5,
            #           (0, 0, 255),
            #           -1)
            # cv2.imshow('Matched Result', display_img)
            # cv2.waitKey(2000)  # 等待按键
            # cv2.destroyAllWindows()
            
            self.logger.info(f"找到模板 {template_name} 位置: ({center_x}, {center_y}), 匹配度: {max_val:.2f}")
            return (center_x, center_y)
            
        except Exception as e:
            self.logger.error(f"定位模板失败: {str(e)}")
            cv2.destroyAllWindows()
            return None

    def check_and_setup_digiwin(self):
        """检查并设置鼎捷ERP窗口"""
        digiwin_title = '鼎捷ERP E10 [华芯微正式|xinxf|苏州华芯微电子股份有限公司|华芯微工厂|华芯微销售域|华芯微公司采购域]'
        digiwin_path = r'D:\Programs\Digiwin\E10\Client\Digiwin.Mars.Deployment.Client.exe'
        
        try:
            # 首先尝试查找已存在的窗口
            hwnd, class_name = find_window_by_title(digiwin_title, timeout=2)
            
            if not hwnd:
                self.logger.info("未找到鼎捷ERP窗口，正在启动程序...")
                # 启动程序并等待窗口出现
                hwnd, class_name = start_and_find_window(digiwin_path, digiwin_title, timeout=30)
                
                if not hwnd:
                    self.logger.error("无法启动鼎捷ERP或找到窗口")
                    return False
                    
                self.logger.info("已成功启动鼎捷ERP")
            else:
                self.logger.info("已找到现有的鼎捷ERP窗口")
            
            # 激活并最大化窗口
            try:
                # 使用win32gui激活窗口
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)  # 先还原窗口
                win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)  # 再最大化
                win32gui.SetForegroundWindow(hwnd)  # 设置为前台窗口
                
                self.logger.info("鼎捷ERP窗口已最大化并置于最前面")
                time.sleep(2)  # 等待窗口动画完成
                return True
                
            except Exception as e:
                self.logger.error(f"设置窗口状态失败: {str(e)}")
                return False
                
        except Exception as e:
            self.logger.error(f"处理鼎捷ERP窗口时出错: {str(e)}")
            return False

def check_emergency_stop():
    return keyboard.is_pressed('esc')

def main():
    try:
        json_file = "downloads/shipping/池州华宇送货单/处理结果.json"
        processor = AutoGuiProcessor(json_file)
        # 检查并设置digiwin窗口
        if not processor.check_and_setup_digiwin():
            logging.error("鼎捷ERP设置失败")
            return
        
        # 定位模板
        center_x, center_y = processor.locate_template('index')
            
        if check_emergency_stop():
            logging.info("检测到紧急停止信号")
            return
            
    except Exception as e:
        logging.error(f"程序执行失败: {str(e)}")
        raise

if __name__ == "__main__":
    main()