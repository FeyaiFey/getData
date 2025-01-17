import win32gui
import win32con
import time
import os
from utils.log_handler import LogHandler

logger = LogHandler().get_logger('WindowFinder', file_level='DEBUG', console_level='INFO')

def enum_windows_callback(hwnd, window_info):
    """窗口枚举回调函数"""
    if win32gui.IsWindowVisible(hwnd):
        window_text = win32gui.GetWindowText(hwnd)
        class_name = win32gui.GetClassName(hwnd)
        if window_text:  # 只记录有标题的窗口
            window_info.append({
                'hwnd': hwnd,
                'title': window_text,
                'class': class_name
            })

def get_all_windows():
    """获取所有可见窗口的信息"""
    window_info = []
    win32gui.EnumWindows(enum_windows_callback, window_info)
    return window_info

def print_window_info():
    """打印所有窗口信息"""
    windows = get_all_windows()
    logger.info("\n所有可见窗口:")
    logger.info("-" * 80)
    logger.info("%-40s %-30s %s", "窗口标题", "类名", "句柄")
    logger.info("-" * 80)
    for window in windows:
        logger.info("%-40s %-30s %d", 
                   window['title'][:38], 
                   window['class'][:28], 
                   window['hwnd'])

def find_window_by_title(title_part, timeout=10):
    """查找包含指定标题部分的窗口句柄"""
    start_time = time.time()
    while time.time() - start_time < timeout:
        windows = get_all_windows()
        for window in windows:
            if title_part.lower() in window['title'].lower():
                logger.debug("找到窗口 [%s] - 句柄: %d", window['title'], window['hwnd'])
                return window['hwnd'], window['class']
        logger.debug("未找到窗口 [%s]，继续搜索...", title_part)
        time.sleep(0.5)
    logger.warning("查找窗口超时 [%s]", title_part)
    return None, None

def start_and_find_window(file_path, title_part=None, timeout=10):
    """启动应用程序并获取其窗口句柄"""
    # 如果没有指定标题部分，使用文件名
    if title_part is None:
        title_part = os.path.basename(file_path)
    
    try:
        # 启动应用程序
        logger.info("启动程序 [%s]", file_path)
        os.startfile(file_path)
        
        # 等待并查找窗口
        hwnd, class_name = find_window_by_title(title_part, timeout)
        if hwnd:
            logger.info("成功启动并找到窗口 [%s]", title_part)
        else:
            logger.error("启动程序后未找到窗口 [%s]", title_part)
        return hwnd, class_name
    except Exception as e:
        logger.error("启动程序失败 [%s]: %s", file_path, LogHandler.format_error(e))
        return None, None

if __name__ == "__main__":
    # 示例：打印所有窗口信息
    print_window_info()
    
    # 示例：启动记事本并查找其窗口
    # hwnd, class_name = start_and_find_window("notepad.exe", "记事本") 