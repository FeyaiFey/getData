import cv2
import numpy as np
import time
import pyautogui
import keyboard
from pathlib import Path
from datetime import datetime
import win32gui
import win32con
from utils.window_finder import find_window_by_title, start_and_find_window
from utils.log_handler import LogHandler
from typing import Dict, Tuple, Optional, List, Any

class AutoGuiProcessor:
    # 窗口标题配置
    ERP_WINDOW = '鼎捷ERP E10 [华芯微正式|xinxf|苏州华芯微电子股份有限公司|华芯微工厂|华芯微销售域|华芯微公司采购域]'
    RECEIPT_WINDOW = '浏览 - 维护到货单'
    NEW_RECEIPT_WINDOW = '维护到货单'
    
    # 模板配置
    TEMPLATES = {
        'index': 'index.png',
        'ic_design_system': 'ic_design_system.png',
        'receipt_button': 'receipt_button.png',
        'receipt_main': 'receipt_main.png',
        'receipt_new': 'receipt_new.png',
        'receipt_new_main': 'receipt_new_main.png',
        'receipt_remark': 'receipt_remark.png',
        'receipt_resource_id': 'receipt_resource_id.png',
        'receipt_businessQty': 'receipt_businessQty.png',
        'receipt_region_paste': 'receipt_region_paste.png',
        'receipt_supply': 'receipt_supply.png',
    }
    
    def __init__(self, data_dict: Dict[str, List[Dict[str, Any]]] = None):
        """初始化自动化处理器
        
        Args:
            data_dict: 送货单数据字典，格式为 {日期: [{订单数据}, ...]}
        """
        self.data_dict = data_dict
        self.logger = LogHandler().get_logger('AutoGui')
        self.templates = self._load_templates()
        
    def _load_templates(self) -> Dict[str, np.ndarray]:
        """加载所有模板图片
        
        Returns:
            Dict[str, np.ndarray]: 模板名称和图片数据的字典
        """
        templates = {}
        templates_dir = Path('templates')
        
        if not templates_dir.exists():
            self.logger.error("模板目录不存在: %s", templates_dir)
            raise FileNotFoundError(f"模板目录不存在: {templates_dir}")
        
        for name, filename in self.TEMPLATES.items():
            template_path = templates_dir / filename
            if not template_path.exists():
                self.logger.error("模板文件不存在: %s", template_path)
                raise FileNotFoundError(f"模板文件不存在: {template_path}")
            
            try:
                abs_path = str(template_path.absolute())
                template = cv2.imdecode(np.fromfile(abs_path, dtype=np.uint8), cv2.IMREAD_COLOR)
                
                if template is None:
                    self.logger.error("无法读取模板文件内容: %s", abs_path)
                    raise FileNotFoundError(f"无法读取模板文件内容: {abs_path}")
                
                templates[name] = template
                self.logger.info("成功加载模板: %s", name)
                
            except Exception as e:
                self.logger.error("加载模板失败: %s", LogHandler.format_error(e))
                raise
        
        return templates

    def locate_template(self, template_name: str) -> Optional[Tuple[int, int]]:
        """定位模板在屏幕上的位置
        
        Args:
            template_name: 模板名称
            
        Returns:
            Optional[Tuple[int, int]]: 模板中心点坐标，如果未找到则返回None
        """
        try:
            template = self.templates[template_name]
            screenshot = pyautogui.screenshot()
            screenshot_cv = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
            
            result = cv2.matchTemplate(screenshot_cv, template, cv2.TM_CCOEFF_NORMED)
            min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
            
            if max_val < 0.8:  # 匹配度阈值
                self.logger.warning("模板 %s 匹配度过低: %.2f", template_name, max_val)
                return None
                
            template_h, template_w = template.shape[:2]
            center_x = max_loc[0] + template_w // 2
            center_y = max_loc[1] + template_h // 2
            
            self.logger.info("找到模板 %s 位置: (%d, %d), 匹配度: %.2f", 
                           template_name, center_x, center_y, max_val)
            return (center_x, center_y)
            
        except Exception as e:
            self.logger.error("定位模板失败: %s", LogHandler.format_error(e))
            return None
        
    def locate_and_click_template(self, template_name: str, window_title: Optional[str] = None, 
                                duration: float = 0.5, wait_time: float = 1, 
                                click: bool = True) -> bool:
        """定位并点击模板
        
        Args:
            template_name: 模板名称
            window_title: 需要最大化的窗口标题
            duration: 鼠标移动持续时间
            wait_time: 点击后等待时间
            click: 是否点击模板
            
        Returns:
            bool: 是否成功定位或点击模板
        """
        try:
            # 如果提供了窗口标题,先最大化窗口
            if window_title and not self._setup_window(window_title):
                return False

            # 重试5次定位模板
            for i in range(5):
                # 检查紧急停止
                if check_emergency_stop():
                    self.logger.warning("检测到ESC按键，停止操作")
                    return False
                    
                # 定位模板
                center_pos = self.locate_template(template_name)
                if center_pos:
                    center_x, center_y = center_pos
                    
                    # 根据click参数决定是否点击
                    if click:
                        pyautogui.moveTo(center_x, center_y, duration=duration)
                        pyautogui.click(center_x, center_y)
                        time.sleep(wait_time)
                    
                    return True
                
                self.logger.warning("第%d次定位模板 %s 失败，等待1秒后重试", i+1, template_name)
                time.sleep(1)
            
            self.logger.error("定位模板 %s 失败，已重试5次", template_name)
            return False
            
        except Exception as e:
            self.logger.error("%s模板 %s 失败: %s", 
                            '点击' if click else '定位', 
                            template_name, 
                            LogHandler.format_error(e))
            return False

    def _setup_window(self, window_title: str) -> bool:
        """设置窗口状态（最大化并置于前台）
        
        Args:
            window_title: 窗口标题
            
        Returns:
            bool: 是否成功设置窗口
        """
        try:
            hwnd = win32gui.FindWindow(None, window_title)
            if hwnd:
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
                win32gui.SetForegroundWindow(hwnd)
                time.sleep(1)
                return True
            else:
                self.logger.warning("未找到窗口: %s", window_title)
                return False
        except Exception as e:
            self.logger.error("设置窗口状态失败: %s", LogHandler.format_error(e))
            return False

    def check_and_setup_digiwin(self) -> bool:
        """检查并设置鼎捷ERP窗口
        
        Returns:
            bool: 是否成功设置ERP窗口
        """
        digiwin_path = r'D:\Programs\Digiwin\E10\Client\Digiwin.Mars.Deployment.Client.exe'
        
        try:
            # 首先尝试查找已存在的窗口
            hwnd, class_name = find_window_by_title(self.ERP_WINDOW, timeout=2)
            
            if not hwnd:
                self.logger.info("未找到鼎捷ERP窗口，正在启动程序...")
                # 启动程序并等待窗口出现
                hwnd, class_name = start_and_find_window(digiwin_path, self.ERP_WINDOW, timeout=30)
                
                if not hwnd:
                    self.logger.error("无法启动鼎捷ERP或找到窗口")
                    return False
                    
                self.logger.info("已成功启动鼎捷ERP")
            else:
                self.logger.info("已找到现有的鼎捷ERP窗口")
            
            # 激活并最大化窗口
            return self._setup_window(self.ERP_WINDOW)
                
        except Exception as e:
            self.logger.error("处理鼎捷ERP窗口时出错: %s", LogHandler.format_error(e))
            return False

def check_emergency_stop() -> bool:
    """检查是否按下ESC键以紧急停止
    
    Returns:
        bool: 是否需要停止操作
    """
    return keyboard.is_pressed('esc')