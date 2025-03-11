import tkinter as tk
from ui import ScreenshotApp
import traceback
import sys

def exception_handler(exc_type, exc_value, exc_traceback):
    """全局异常处理函数"""
    error_msg = ''.join(traceback.format_exception(exc_type, exc_value, exc_traceback))
    print(f"发生错误:\n{error_msg}")
    tk.messagebox.showerror("程序错误", f"发生错误:\n{str(exc_value)}\n\n请检查文件路径是否正确")

if __name__ == "__main__":
    # 设置全局异常处理
    sys.excepthook = exception_handler
    
    root = tk.Tk()
    app = ScreenshotApp(root)
    root.mainloop()