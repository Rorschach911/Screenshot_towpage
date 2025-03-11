import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from screenshot import take_screenshots
import excel_handler
import os

class ScreenshotApp:
    def __init__(self, root):
        self.root = root
        self.root.title("网页截图工具")
        self.root.geometry("600x600")
        self.root.configure(bg='#e6ffe6')

        # 初始化变量
        self.excel_path = tk.StringVar()
        self.ppt_path = tk.StringVar()
        self.title_text = tk.StringVar()  # 添加标题变量
        # 删除发布日期变量
        self.create_widgets()

    def select_excel(self):
        filename = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel files", "*.xlsx;*.xls")]
        )
        if filename:
            self.excel_path.set(filename)

    def select_ppt(self):
        filename = filedialog.askopenfilename(
            title="选择PPT文件",
            filetypes=[("PowerPoint files", "*.pptx")]
        )
        if filename:
            self.ppt_path.set(filename)

    def create_widgets(self):
        # Excel文件选择
        tk.Label(self.root, text="导入Excel", bg='#e6ffe6').grid(row=0, column=0, padx=10, pady=10)
        tk.Entry(self.root, textvariable=self.excel_path, width=50).grid(row=0, column=1, padx=5)
        tk.Button(self.root, text="打开", command=self.select_excel).grid(row=0, column=2, padx=5)

        # PPT文件选择
        tk.Label(self.root, text="导出PPT", bg='#e6ffe6').grid(row=1, column=0, padx=10, pady=10)
        tk.Entry(self.root, textvariable=self.ppt_path, width=50).grid(row=1, column=1, padx=5)
        tk.Button(self.root, text="打开", command=self.select_ppt).grid(row=1, column=2, padx=5)

        # 添加标题输入框
        tk.Label(self.root, text="页面标题", bg='#e6ffe6').grid(row=2, column=0, padx=10, pady=10)
        tk.Entry(self.root, textvariable=self.title_text, width=50).grid(row=2, column=1, padx=5)

        # 执行按钮移到第3行 (原来是第4行)
        tk.Button(self.root, text="执 行", command=self.start_process, width=20).grid(row=3, column=1, pady=20)

        # 进度条移到第4行 (原来是第5行)
        self.progress_label = tk.Label(self.root, text="执行进度", bg='#e6ffe6')
        self.progress_label.grid(row=4, column=0, padx=10)
        
        self.progress = ttk.Progressbar(self.root, length=400, mode='determinate')
        self.progress.grid(row=4, column=1, padx=5)
        
        self.progress_text = tk.Label(self.root, text="0%", bg='#e6ffe6')
        self.progress_text.grid(row=4, column=2)

        # 链接列表标签移到第5行 (原来是第6行)
        tk.Label(self.root, text="链接列表", bg='#e6ffe6').grid(row=5, column=0, padx=10, pady=5)
        
        # 文本框移到第6行 (原来是第7行)
        frame = tk.Frame(self.root)
        frame.grid(row=6, column=0, columnspan=3, padx=10, pady=5, sticky='nsew')
        
        self.link_list = tk.Text(frame, height=15, width=80)
        scrollbar = tk.Scrollbar(frame, command=self.link_list.yview)
        self.link_list.configure(yscrollcommand=scrollbar.set)
        
        self.link_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 配置网格权重使链接列表可以随窗口调整大小
        self.root.grid_rowconfigure(4, weight=1)  # 更新行号
        self.root.grid_columnconfigure(1, weight=1)

    def update_progress(self, current, total):
        progress = (current + 1) / total * 100
        self.progress['value'] = progress
        self.progress_text['text'] = f"{int(progress)}%"
        self.root.update()

    def update_link_status(self, index, status="已完成"):
        start_idx = f"{index+1}."
        line_start = self.link_list.search(start_idx, "1.0", tk.END)
        if line_start:
            line_end = self.link_list.search("\n", line_start, tk.END)
            self.link_list.insert(line_end, f" [{status}]")
            self.link_list.see(line_start)

    def start_process(self):
        if not self.excel_path.get() or not self.ppt_path.get():
            tk.messagebox.showerror("错误", "请先选择Excel和PPT文件")
            return

        try:
            # 读取Excel文件
            df = excel_handler.read_excel(self.excel_path.get())
            
            # 检查必要的列是否存在
            required_columns = ["媒体名称", "链接", "发布时间"]
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                tk.messagebox.showerror("错误", f"Excel文件中缺少以下列: {', '.join(missing_columns)}")
                return
            
            # 清空并显示链接列表
            self.link_list.delete(1.0, tk.END)
            for index, row in df.iterrows():
                self.link_list.insert(tk.END, f"{index+1}. {row['媒体名称']} - {row['链接']}\n")
            
            # 打开PPT
            os.startfile(self.ppt_path.get())
            
            # 执行截图过程
            take_screenshots(
                df, 
                self.ppt_path.get(), 
                self.title_text.get(),
                self.update_progress,
                self.update_link_status,
                self.root
            )
            
            tk.messagebox.showinfo("完成", "截图已完成并保存到PPT中")
            
        except Exception as e:
            tk.messagebox.showerror("错误", f"执行过程中出现错误：{str(e)}")