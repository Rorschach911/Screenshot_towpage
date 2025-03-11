# 网页截图工具
每页幻灯片存2组网页截图和文字信息

## 项目简介

这是一个自动化网页截图工具，可以根据Excel文件中提供的链接地址自动打开网页，进行截图，并将截图与相关信息一起添加到PowerPoint演示文稿中。该工具特别适合需要批量收集网页内容并制作演示文稿的场景，如媒体监测、市场调研等。

## 功能特点

- 从Excel文件中批量读取网页链接及相关信息
- 自动打开网页并进行截图
- 将截图和相关信息（媒体名称、发布时间、标题、链接）添加到PPT中
- 实时显示处理进度和状态
- 友好的图形用户界面
- 自动调整PPT中的字体大小和图片尺寸

## 系统要求

- Windows操作系统
- Python 3.6+
- Microsoft PowerPoint
- Google Chrome浏览器

## 安装依赖

```bash
pip install pandas openpyxl python-pptx selenium webdriver-manager pyautogui pillow pywin32
```

## 使用方法

1. 运行程序：

```bash
python main.py
```

2. 在界面中选择包含网页链接的Excel文件（必须包含"媒体名称"、"发布时间"和"链接"列）
3. 选择或创建要保存截图的PPT文件
4. 输入页面标题
5. 点击"执行"按钮开始处理
6. 程序会自动打开Chrome浏览器访问每个链接，截图并添加到PPT中
7. 处理过程中可以在界面下方查看进度和状态

## Excel文件格式要求

Excel文件必须包含以下列：
- 媒体名称：网站或媒体的名称
- 发布时间：内容的发布时间
- 链接：需要截图的网页URL

示例：

| 媒体名称 | 发布时间 | 链接 |
|---------|---------|-----|
| 新浪网 | 2023-01-01 | https://www.sina.com.cn |
| 腾讯网 | 2023-01-02 | https://www.qq.com |

## 项目结构

- `main.py` - 程序入口点
- `ui.py` - 用户界面相关代码
- `screenshot.py` - 截图功能相关代码
- `ppt_handler.py` - PPT处理相关代码
- `excel_handler.py` - Excel处理相关代码

## 技术实现

- 使用 Selenium 进行网页自动化
- 使用 python-pptx 处理 PPT 文件
- 使用 pandas 处理 Excel 数据
- 使用 tkinter 构建图形界面
- 使用 win32com 进行 PPT 操作

## 注意事项

1. 确保 Excel 文件格式正确，包含所有必需列
2. 处理过程中请勿关闭 Chrome 浏览器
3. 确保有足够的磁盘空间存储临时文件
4. 建议定期保存 PPT 文件
5. 处理大量链接时可能需要较长时间

## 常见问题

1. **程序无法启动**
   - 确保已安装所有依赖包
   - 检查Python版本是否兼容

2. **无法打开Chrome浏览器**
   - 确保已安装Chrome浏览器
   - 检查webdriver是否与Chrome版本匹配

3. **PPT无法保存**
   - 确保PowerPoint未被其他程序占用
   - 检查保存路径是否有写入权限

4. **Excel文件读取错误**
   - 确保Excel文件格式正确
   - 检查是否包含所有必需的列

## 更新日志

### v1.0.0
- 实现基础的网页截图功能
- 支持 Excel 批量导入
- 添加进度显示功能
- 实现自动 PPT 生成
