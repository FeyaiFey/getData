import win32gui
import win32con
import time
import os

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
    print("\n所有可见窗口:")
    print("-" * 80)
    print(f"{'窗口标题':<40} {'类名':<30} {'句柄'}")
    print("-" * 80)
    for window in windows:
        print(f"{window['title'][:38]:<40} {window['class'][:28]:<30} {window['hwnd']}")

def find_window_by_title(title_part, timeout=10):
    """
    查找包含指定标题部分的窗口句柄
    
    Args:
        title_part: 窗口标题的部分内容
        timeout: 超时时间（秒）
        
    Returns:
        tuple: (hwnd, class_name) 窗口句柄和类名，如果未找到返回 (None, None)
    """
    start_time = time.time()
    while time.time() - start_time < timeout:
        windows = get_all_windows()
        for window in windows:
            if title_part.lower() in window['title'].lower():
                return window['hwnd'], window['class']
        time.sleep(0.5)
    return None, None

def start_and_find_window(file_path, title_part=None, timeout=10):
    """
    启动应用程序并获取其窗口句柄
    
    Args:
        file_path: 要启动的文件路径
        title_part: 窗口标题的部分内容（如果为None，则使用文件名）
        timeout: 查找窗口的超时时间（秒）
        
    Returns:
        tuple: (hwnd, class_name) 窗口句柄和类名，如果未找到返回 (None, None)
    """
    # 如果没有指定标题部分，使用文件名
    if title_part is None:
        title_part = os.path.basename(file_path)
    
    # 启动应用程序
    os.startfile(file_path)
    
    # 等待并查找窗口
    return find_window_by_title(title_part, timeout)

if __name__ == "__main__":
    # 示例：打印所有窗口信息
    print_window_info()
    
    # 示例：启动记事本并查找其窗口
    # hwnd, class_name = start_and_find_window("notepad.exe", "记事本") 