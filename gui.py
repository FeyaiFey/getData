import win32gui
import pyautogui
import time

def list_all_windows():
    """
    列出所有可见窗口的标题
    """
    def callback(hwnd, windows):
        if win32gui.IsWindowVisible(hwnd):
            window_title = win32gui.GetWindowText(hwnd)
            if window_title:
                windows.append(f"窗口句柄: {hwnd}, 标题: {window_title}")
        return True
    
    windows = []
    win32gui.EnumWindows(callback, windows)
    
    print("当前所有可见窗口:")
    for window in windows:
        print(window)

time.sleep(5)
pyautogui.moveTo(490, 139)
pyautogui.click()
time.sleep(1)
# 使用剪贴板复制粘贴中文,避免输入法问题
import pyperclip
pyperclip.copy("池州华宇")
pyautogui.hotkey('ctrl', 'v')
time.sleep(1)
pyautogui.press('enter')

# if __name__ == "__main__":
#     list_all_windows()
