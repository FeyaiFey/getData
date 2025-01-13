from services.autogui import AutoGuiProcessor
import pyautogui
import logging
import time

def main():
    try:
        json_file = "downloads/shipping/池州华宇送货单/处理结果.json"
        processor = AutoGuiProcessor(json_file)
        
        # 检查并设置digiwin窗口
        if not processor.check_and_setup_digiwin():
            logging.error("鼎捷ERP设置失败")
            return
            
        # 定位模板
        index_center_x, index_center_y = processor.locate_template('index')
        if index_center_x and index_center_y:
            logging.info(f"模板匹配成功，中心点坐标: ({index_center_x}, {index_center_y})")
        else:
            logging.error("模板匹配失败")
            return

        # 点击模板
        pyautogui.moveTo(index_center_x, index_center_y,duration=1)
        pyautogui.click(index_center_x, index_center_y)
        time.sleep(2)
        if processor.locate_template('ic_design_system'):
            logging.info("已成功点击ic_design_system")
        else:
            logging.error("点击ic_design_system失败")
            return
        
        # 定位模板
        receipt_button_center_x, receipt_button_center_y = processor.locate_template('receipt_button')
        if receipt_button_center_x and receipt_button_center_y:
            logging.info(f"模板匹配成功，中心点坐标: ({receipt_button_center_x}, {receipt_button_center_y})")
        else:
            logging.error("模板匹配失败")
            return
        
        # 点击模板
        pyautogui.moveTo(receipt_button_center_x, receipt_button_center_y,duration=1)
        pyautogui.click(receipt_button_center_x, receipt_button_center_y)
        time.sleep(2)
        if processor.locate_template('receipt_main'):
            logging.info("已成功点击receipt_button")
        else:
            logging.error("点击receipt_button失败")
            return


    except Exception as e:
        logging.error(f"程序执行失败: {str(e)}")
        raise

if __name__ == "__main__":
    main() 