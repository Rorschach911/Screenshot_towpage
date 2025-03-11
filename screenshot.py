from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import time
import os
import pyautogui
from PIL import Image
import io
import win32clipboard
from ppt_handler import add_slide_and_content_pair  # 修正导入语句，只保留正确的函数

def take_screenshots(df, ppt_path, title_text, update_progress, update_link_status, root):
    """
    执行网页截图并添加到PPT的主要函数，每页显示两组数据
    """
    chrome_options = Options()
 # 设置Chrome窗口大小为1290*900
    chrome_options.add_argument('--window-size=1290,900')
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), 
                           options=chrome_options)
    driver.set_window_size(1290, 900)  # 确保窗口大小为1290*900
    
    time.sleep(3)
    total = len(df)
    
    # 创建临时目录用于保存截图
    temp_dir = os.path.join(os.path.dirname(os.path.abspath(ppt_path)), "temp_screenshots")
    if not os.path.exists(temp_dir):
        os.makedirs(temp_dir)
    
    try:
        # 每次处理两条记录
        for i in range(0, total, 2):
            update_progress(i, total)
            
            # 第一组数据
            row1 = df.iloc[i]
            media_name1 = row1['媒体名称']
            publish_time1 = row1['发布时间']
            link1 = row1['链接']
            
            # 访问第一个网页并截图
            driver.get(link1)
            time.sleep(3)
            screenshot_path1 = os.path.join(temp_dir, f"temp_screenshot_{i}.png")
            driver.save_screenshot(screenshot_path1)
            
            # 准备第二组数据（如果存在）
            media_name2 = None
            publish_time2 = None
            link2 = None
            screenshot_path2 = None
            
            if i + 1 < total:
                row2 = df.iloc[i + 1]
                media_name2 = row2['媒体名称']
                publish_time2 = row2['发布时间']
                link2 = row2['链接']
                
                # 访问第二个网页并截图
                driver.get(link2)
                time.sleep(3)
                screenshot_path2 = os.path.join(temp_dir, f"temp_screenshot_{i+1}.png")
                driver.save_screenshot(screenshot_path2)
            
            # 删除错误的函数调用，只保留正确的函数调用
            slide = add_slide_and_content_pair(
                media_name1, publish_time1, title_text, link1,
                media_name2, publish_time2, title_text, link2
            )
            
            if not slide:
                continue
            
            # 添加第一张截图
            if os.path.exists(screenshot_path1):
                left1 = 30  # 0.5英寸
                top = 210   # 2.5英寸（文本框下方）
                width = 430  # 4.5英寸
                height = 300  # 3.5英寸
                slide.Shapes.AddPicture(
                    screenshot_path1, 
                    False, True, 
                    left1, top, 
                    width, height
                )
                os.remove(screenshot_path1)
            
            # 添加第二张截图（如果存在）
            if screenshot_path2 and os.path.exists(screenshot_path2):
                left2 = 490  # 5.5英寸
                slide.Shapes.AddPicture(
                    screenshot_path2, 
                    False, True, 
                    left2, top, 
                    width, height
                )
                os.remove(screenshot_path2)
            
            update_link_status(i)
            if i + 1 < total:
                update_link_status(i + 1)
            root.update()
            
    finally:
        # 清理临时文件
        if os.path.exists(temp_dir):
            for file in os.listdir(temp_dir):
                try:
                    os.remove(os.path.join(temp_dir, file))
                except:
                    pass
            try:
                os.rmdir(temp_dir)
            except:
                pass
        driver.quit()

def copy_to_clipboard(image_path):
    """将图片复制到剪贴板"""
    image = Image.open(image_path)
    output = io.BytesIO()
    image.convert('RGB').save(output, 'BMP')
    data = output.getvalue()[14:]
    output.close()
    
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
    win32clipboard.CloseClipboard()

def paste_to_ppt(row, title_text):
    """将截图粘贴到PPT"""
    # 点击幻灯片中间位置
    pyautogui.click(x=500, y=500)
    time.sleep(0.5)
    
    # 粘贴图片
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)
    
