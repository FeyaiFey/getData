import cv2
import numpy as np
import time
import pyautogui
import keyboard
import win32gui
import win32con
import os
from datetime import datetime
from pathlib import Path
from typing import Optional, Tuple, Dict, List, Any
from utils.log_handler import LogHandler


class AutoGuiProcessor:
    """自动化处理器，负责执行自动化操作
    
    主要功能：
    1. ERP窗口管理（打开、最大化、置前）
    2. 模板匹配和定位
    3. 错误处理和日志记录
    4. 截图保存
    5. 紧急停止
    """
    
    # 配置常量
    MAX_RETRIES = 5          # 最大重试次数
    RETRY_INTERVAL = 1.0     # 重试间隔（秒）
    MATCH_THRESHOLD = 0.8    # 模板匹配阈值
    SCREENSHOT_DIR = "screenshots"  # 截图保存目录
    TEMPLATE_DIR = "templates"      # 模板图片目录
    
    # ERP相关常量
    ERP_PATH = r'D:\Programs\Digiwin\E10\Client\Digiwin.Mars.Deployment.Client.exe'
    ERP_WINDOW = "鼎捷ERP E10 [华芯微正式|xinxf|苏州华芯微电子股份有限公司|华芯微工厂|华芯微销售域|华芯微公司采购域]"
    RECEIPT_WINDOW = '浏览 - 维护到货单'
    NEW_RECEIPT_WINDOW = '维护到货单'
    AUDIT_WINDOW = '审核'
    
    # 模板文件名常量
    TEMPLATES = {
        # 系统操作
        'IC_DESIGN_SYSTEM': 'ic_design_system_button.png',   # IC设计管理系统
        'RECEIPT_BUTTON': 'receipt_button.png',               # 维护到货单按钮
        'RECEIPT_NEW': 'receipt_new.png',                   # 新增按钮
        'RECEIPT_MAIN': 'receipt_main.png',                   # 新增到货单界面
        'RECEIPT_NEW_MAIN': 'receipt_new_main.png',         # 新增到货单界面
        'SAVE': 'save.png',                                 # 保存按钮
        'AUDIT': 'audit.png',                               # 审核按钮
        'CONFIRM': 'confirm.png',                         # 确认按钮
        'YES': 'yes.png',                                   # 是按钮
        'NO': 'no.png',                                     # 否按钮
        'WARNING': 'warning.png',
        
        # 到货单字段
        'RECEIPT_SUPPLY': 'receipt_supply.png',             # 供应商
        'RECEIPT_RESOURCE_ID': 'receipt_resource_id.png',   # 来源单号
        'RECEIPT_BUSINESS_QTY': 'receipt_businessQty.png',   # 数量
        'RECEIPT_REGION_PASTE': 'receipt_region_paste.png', # 区域粘贴
        'RECEIPT_REMARK': 'receipt_remark.png',             # 备注
        'DOCUMENT_TYPE': 'document_type.png',               # 单别
        
        # 错误和警告
        'RECEIPT_ERROR_1': 'receipt_error_1.png',               # 错误
        'RECEIPT_ERROR_2': 'receipt_error_2.png',               # 错误
        'RECEIPT_WARNING_1': 'receipt_warning_1.png',       # 警告
        'RECEIPT_WARNING_2': 'receipt_warning_2.png',       # 单据警告1
    }

    def __init__(self, data_dict: Dict[str, List[Dict[str, Any]]] = None):
        """初始化自动化处理器"""
        self.logger = LogHandler().get_logger('AutoGui', file_level='DEBUG', console_level='INFO')
        self._init_dirs()
        self.templates = {}  # 模板缓存
        self.data_dict = data_dict

        
    def _init_dirs(self):
        """初始化必要的目录"""
        try:
            # 创建截图目录
            os.makedirs(self.SCREENSHOT_DIR, exist_ok=True)
            self.logger.debug("截图目录已就绪: %s", self.SCREENSHOT_DIR)
            
            # 检查模板目录
            if not os.path.exists(self.TEMPLATE_DIR):
                self.logger.error("模板目录不存在: %s", self.TEMPLATE_DIR)
                raise FileNotFoundError(f"模板目录不存在: {self.TEMPLATE_DIR}")
                
            # 验证模板文件
            missing_templates = []
            for template_name in self.TEMPLATES.values():
                template_path = os.path.join(self.TEMPLATE_DIR, template_name)
                if not os.path.exists(template_path):
                    missing_templates.append(template_name)
                    
            if missing_templates:
                self.logger.error("以下模板文件缺失: %s", ", ".join(missing_templates))
                raise FileNotFoundError(f"缺失模板文件: {', '.join(missing_templates)}")
                
            self.logger.debug("模板目录验证完成")
            
        except Exception as e:
            self.logger.error("初始化目录失败: %s", str(e))
            raise
            
    def get_template_path(self, template_key: str) -> str:
        """获取模板文件的完整路径
        
        Args:
            template_key: 模板键名，使用TEMPLATES中定义的常量
            
        Returns:
            str: 模板文件的完整路径
            
        Raises:
            KeyError: 如果模板键名不存在
        """
        if template_key not in self.TEMPLATES:
            self.logger.error("未知的模板键名: %s", template_key)
            raise KeyError(f"未知的模板键名: {template_key}")
            
        return os.path.join(self.TEMPLATE_DIR, self.TEMPLATES[template_key])
            
    def take_screenshot(self, name: str) -> Optional[str]:
        """保存当前屏幕截图
        
        Args:
            name: 截图名称（不含扩展名）
            
        Returns:
            Optional[str]: 截图文件路径，失败返回None
        """
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"{name}_{timestamp}.png"
            filepath = os.path.join(self.SCREENSHOT_DIR, filename)
            
            screenshot = pyautogui.screenshot()
            screenshot.save(filepath)
            
            self.logger.debug("已保存截图: %s", filepath)
            return filepath
        except Exception as e:
            self.logger.error("保存截图失败: %s", str(e))
            return None
            
    def setup_window(self, window_title: str, timeout: int = 1) -> bool:
        """设置窗口状态（最大化并置于前台）
        
        Args:
            window_title: 窗口标题
            timeout: 等待超时时间（秒）
            
        Returns:
            bool: 是否成功设置窗口
        """
        try:
            start_time = time.time()
            while time.time() - start_time < timeout:
                hwnd = win32gui.FindWindow(None, window_title)
                if hwnd:
                    # 恢复窗口（如果最小化）
                    if win32gui.IsIconic(hwnd):
                        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                        time.sleep(0.5)
                    
                    # 最大化窗口
                    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
                    time.sleep(0.5)
                    
                    # 置于前台
                    win32gui.SetForegroundWindow(hwnd)
                    time.sleep(0.5)
                    
                    self.logger.debug("窗口已设置: %s", window_title)
                    return True
                    
                time.sleep(0.5)
                
            self.logger.error("未找到窗口: %s", window_title)
            return False
            
        except Exception as e:
            self.logger.error("设置窗口失败 [%s]: %s", window_title, str(e))
            return False
            
    def open_erp(self) -> bool:
        """打开或激活ERP窗口
        
        Args:
            erp_path: ERP程序路径
            window_title: ERP窗口标题
            
        Returns:
            bool: 是否成功打开或激活ERP
        """
        try:
            # 先尝试查找现有窗口
            if self.setup_window(self.ERP_WINDOW, timeout=1):
                self.logger.info("已激活现有ERP窗口")
                return True
                
            # 如果没有找到窗口，启动程序
            self.logger.info("正在启动ERP程序...")
            os.startfile(self.ERP_PATH)
            
            # 等待窗口出现并设置
            if self.setup_window(self.ERP_WINDOW, timeout=30):
                self.logger.info("ERP程序已启动")
                return True
                
            self.logger.error("启动ERP程序失败")
            self.take_screenshot("erp_start_failed")
            return False
            
        except Exception as e:
            self.logger.error("打开ERP失败: %s", str(e))
            self.take_screenshot("erp_open_error")
            return False
            
    def load_template(self, template_path: str) -> Optional[np.ndarray]:
        """加载模板图片
        
        Args:
            template_path: 模板图片路径
            
        Returns:
            Optional[np.ndarray]: 模板图像数据，失败返回None
        """
        try:
            if template_path in self.templates:
                return self.templates[template_path]
                
            if not os.path.exists(template_path):
                self.logger.error("模板文件不存在: %s", template_path)
                return None
                
            # 使用imdecode读取图片，支持中文路径
            template = cv2.imdecode(
                np.fromfile(template_path, dtype=np.uint8), 
                cv2.IMREAD_COLOR
            )
            
            if template is None:
                self.logger.error("无法读取模板: %s", template_path)
                return None
                
            self.templates[template_path] = template
            self.logger.debug("已加载模板: %s", template_path)
            return template
            
        except Exception as e:
            self.logger.error("加载模板失败 [%s]: %s", template_path, str(e))
            return None
            
    def locate_template(self, template_path: str, 
                       confidence: float = None) -> Optional[Tuple[int, int]]:
        """定位模板在屏幕上的位置
        
        Args:
            template_path: 模板图片路径
            confidence: 匹配阈值，None则使用默认值
            
        Returns:
            Optional[Tuple[int, int]]: 模板中心点坐标，失败返回None
        """
        try:
            # 加载模板
            template = self.load_template(template_path)
            if template is None:
                return None
                
            # 获取屏幕截图
            screenshot = pyautogui.screenshot()
            screenshot_cv = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
            
            # 模板匹配
            result = cv2.matchTemplate(screenshot_cv, template, cv2.TM_CCOEFF_NORMED)
            min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
            
            # 检查匹配度
            threshold = confidence if confidence is not None else self.MATCH_THRESHOLD
            self.logger.info("模板 [%s] 匹配度: %.4f (阈值: %.4f)", 
                           os.path.basename(template_path), max_val, threshold)
            
            if max_val < threshold:
                self.logger.debug("模板匹配度过低: %.4f < %.4f", max_val, threshold)
                return None
                
            # 计算中心点坐标
            h, w = template.shape[:2]
            center_x = max_loc[0] + w // 2
            center_y = max_loc[1] + h // 2
            
            self.logger.debug("找到模板 [%s] - 位置: (%d, %d), 匹配度: %.4f", 
                            os.path.basename(template_path), center_x, center_y, max_val)
            return (center_x, center_y)
            
        except Exception as e:
            self.logger.error("定位模板失败 [%s]: %s", template_path, str(e))
            return None
            
    def locate_and_click(self, template: str, 
                        window_title: Optional[str] = None,
                        confidence: float = None,
                        click: bool = True) -> bool:
        """定位并点击模板
        
        Args:
            template: 模板键名或模板文件路径
            window_title: 需要激活的窗口标题
            confidence: 匹配阈值
            click: 是否执行点击
            
        Returns:
            bool: 是否成功定位或点击
        """
        try:
            # 检查紧急停止
            if keyboard.is_pressed('esc'):
                self.logger.warning("检测到ESC键，停止操作")
                return False
                
            # 如果指定了窗口，先激活窗口
            if window_title and not self.setup_window(window_title):
                return False
                
            # 获取模板路径
            template_path = (
                self.get_template_path(template) 
                if template in self.TEMPLATES 
                else template
            )
                
            # 重试定位模板
            for i in range(self.MAX_RETRIES):
                center_pos = self.locate_template(template_path, confidence)
                if center_pos:
                    # 成功前先截图
                    action = "click" if click else "locate"
                    self.take_screenshot(f"{action}_{os.path.basename(template_path)}")
                    
                    if click:
                        center_x, center_y = center_pos
                        pyautogui.moveTo(center_x, center_y, duration=0.5)
                        pyautogui.click()
                        # 点击后再截一次图
                        time.sleep(0.5)  # 等待界面响应
                        self.take_screenshot(f"after_click_{os.path.basename(template_path)}")
                        self.logger.info("已点击模板: %s", template)
                    return True
                    
                if i < self.MAX_RETRIES - 1:
                    self.logger.debug("重试定位模板 [%s] - 第%d次", template, i + 1)
                    time.sleep(self.RETRY_INTERVAL)
                    
            # 所有重试都失败了
            self.logger.error("定位模板失败 [%s] - 已重试%d次", template, self.MAX_RETRIES)
            self.take_screenshot(f"locate_failed_{os.path.basename(template_path)}")
            return False
            
        except Exception as e:
            self.logger.error("%s模板失败 [%s]: %s", 
                            "点击" if click else "定位",
                            template, str(e))
            self.take_screenshot(f"locate_error_{os.path.basename(template_path)}")
            return False

def check_emergency_stop() -> bool:
    """检查是否触发紧急停止（按下ESC键）
    
    Returns:
        bool: 是否需要停止
    """
    return keyboard.is_pressed('esc')
