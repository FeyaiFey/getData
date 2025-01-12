import win32gui
import os
import time
import pyautogui
import cv2
import numpy as np
import pyperclip


def open_excel():
    try:
        os.startfile(r'D:\PycharmProjects\getData\templates\erp_template.xlsx')
        # 等待Excel窗口出现
        print("等待Excel窗口出现...")
        time.sleep(5)
        
        # 查找Excel窗口
        excel_hwnd = win32gui.FindWindow('XLMAIN', None)
        
        if not excel_hwnd:
            print("正在重试查找Excel窗口...")
            time.sleep(2)
            excel_hwnd = win32gui.FindWindow('XLMAIN', None)
            
        if excel_hwnd:
            # 将Excel窗口置于最前面
            win32gui.ShowWindow(excel_hwnd, 9)  # SW_RESTORE = 9
            win32gui.SetForegroundWindow(excel_hwnd)
            print("已找到并激活Excel窗口")
        else:
            print("未能找到Excel窗口，请确认是否正常打开")
            return False
        print("正在打开Excel...")
    except Exception as e:
        print(f"打开Excel失败: {str(e)}")
        return False
    # 最大化Excel窗口
    try:
        win32gui.ShowWindow(excel_hwnd, 3)  # SW_MAXIMIZE = 3
        print("Excel窗口已最大化")
    except Exception as e:
        print(f"最大化Excel窗口失败: {str(e)}")
    pyautogui.moveTo(102, 383, duration=0.5)
    pyautogui.click()
    time.sleep(1)
    pyautogui.moveTo(102, 483, duration=0.5)
    pyautogui.click()
    

open_excel()
