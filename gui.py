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

# time.sleep(5)
# pyautogui.moveTo(490, 139)
# pyautogui.click()
# time.sleep(1)
# # 使用剪贴板复制粘贴中文,避免输入法问题
# import pyperclip
# pyperclip.copy("池州华宇")
# pyautogui.hotkey('ctrl', 'v')
# time.sleep(1)
# pyautogui.press('enter')

def check_clipboard():
    """
    查看剪贴板当前内容
    """
    import pyperclip
    content = pyperclip.paste()
    print(content)

test_data = {
        "2025-01-15": [
            {"订单号": "HX-20250107012", "数量": 100, "供应商": "池州华宇"},
            {"订单号": "HX-20250107013", "数量": 200, "供应商": "池州华宇"}
        ]
    }

# 方法1: 使用字符串拼接
order_numbers = ""
for item in test_data["2025-01-15"]:
    order_numbers += item["订单号"] + "\n"
order_numbers = order_numbers.rstrip()  # 移除最后一个换行符

# 方法2: 使用列表推导式和join
order_numbers = "\r\n".join(item["订单号"] for item in test_data["2025-01-15"])

# 方法3: 使用map函数
# order_numbers = "\n".join(map(lambda x: x["订单号"], test_data["2025-01-15"]))

import pyperclip
pyperclip.copy(order_numbers)

# check_clipboard()
# if __name__ == "__main__":
#     list_all_windows()
