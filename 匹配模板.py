import pyautogui
import cv2
import numpy as np

def match_template(template_path):
    # 读取模板图片
    template = cv2.imread(template_path, cv2.IMREAD_GRAYSCALE)
    
    # 使用pyautogui截取屏幕
    screenshot = pyautogui.screenshot()
    # 将PIL图像转换为OpenCV格式并转为灰度图
    screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2GRAY)
    
    # 模板匹配
    result = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)
    min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
    return max_loc

template_path = r'D:\PycharmProjects\getData\templates\订单号.png'
