import pyautogui
import time
import keyboard

# 设置安全防护，将鼠标移动到屏幕左上角将暂停程序
pyautogui.FAILSAFE = True
# 设置所有pyautogui操作的延迟时间
pyautogui.PAUSE = 0.5

def show_basic_mouse_operations():
    """基本鼠标操作演示"""
    print("=== 基本鼠标操作演示 ===")
    
    # 获取屏幕尺寸
    screen_width, screen_height = pyautogui.size()
    print(f"屏幕尺寸: {screen_width} x {screen_height}")
    
    # 获取当前鼠标位置
    current_x, current_y = pyautogui.position()
    print(f"当前鼠标位置: ({current_x}, {current_y})")
    
    # 移动鼠标（绝对位置）
    print("移动鼠标到屏幕中心...")
    pyautogui.moveTo(screen_width/2, screen_height/2, duration=1)
    
    # 移动鼠标（相对位置）
    print("相对移动鼠标 100 像素...")
    pyautogui.moveRel(100, 0, duration=1)  # 向右移动100像素
    
    # 鼠标点击
    print("执行鼠标点击...")
    pyautogui.click()  # 左键单击
    pyautogui.doubleClick()  # 双击
    pyautogui.rightClick()  # 右键单击
    
    # 鼠标拖拽
    print("执行鼠标拖拽...")
    pyautogui.dragRel(100, 0, duration=1)  # 向右拖拽100像素

def show_keyboard_operations():
    """基本键盘操作演示"""
    print("\n=== 基本键盘操作演示 ===")
    
    # 模拟按键
    print("模拟按下 'shift' 键...")
    pyautogui.keyDown('shift')
    pyautogui.keyUp('shift')
    
    # 输入文本
    print("输入文本...")
    pyautogui.write('Hello, PyAutoGUI!', interval=0.1)  # 每个字符之间间隔0.1秒
    
    # 按下热键组合
    print("按下热键组合 Ctrl+A...")
    pyautogui.hotkey('ctrl', 'a')  # 全选
    
def show_screenshot_operations():
    """截图操作演示"""
    print("\n=== 截图操作演示 ===")
    
    # 截取全屏
    print("截取全屏...")
    screenshot = pyautogui.screenshot()
    screenshot.save('full_screen.png')
    
    # 截取指定区域
    print("截取指定区域...")
    region_screenshot = pyautogui.screenshot(region=(0, 0, 300, 300))  # 左上角300x300像素区域
    region_screenshot.save('region_screen.png')
    
def show_image_recognition():
    """图像识别演示"""
    print("\n=== 图像识别演示 ===")
    
    try:
        # 在屏幕上查找图片
        print("尝试在屏幕上查找图片...")
        # 注意：需要准备一个实际存在的图片文件
        # location = pyautogui.locateOnScreen('target.png', confidence=0.9)
        # print(f"找到图片位置: {location}")
        print("提示：要使用图像识别功能，需要准备目标图片文件")
        
    except pyautogui.ImageNotFoundException:
        print("未找到目标图片")

def main():
    print("PyAutoGUI 快速入门演示")
    print("按 'Esc' 键退出程序")
    print("将鼠标移动到屏幕左上角可以触发 failsafe 功能停止程序")
    
    try:
        # 等待用户准备
        print("\n3秒后开始演示...")
        time.sleep(3)
        
        # 运行演示
        show_basic_mouse_operations()
        show_keyboard_operations()
        show_screenshot_operations()
        show_image_recognition()
        
    except KeyboardInterrupt:
        print("\n程序被用户中断")
    finally:
        print("\n演示结束")

if __name__ == "__main__":
    main() 