# 该文件可进行.docx文件的批量处理，包括删除页眉、页脚和添加页码
# 打包代码
# pyinstaller --name=FileModifier --distpath=dist --clean --onefile --noconsole --add-data "src;src" --hidden-import=docx --hidden-import=docx.shared src\gui\FileModifier.py


import tkinter as tk
from tkinter import filedialog, ttk, messagebox  # 确保导入了messagebox模块
import sys
import os

# 添加工作目录到系统路径，确保src模块可导入
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

# 修正了从src.utils导入WordDocument的路径
from src.utils.file_utils import WordDocument

# 新增：设置 DPI 感知模式
try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(1)
except:
    pass


class MainWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("文件修改工具")
        self.root.geometry("400x200")
        
        # 新增：调用_center_window方法来居中窗口
        self._center_window()

    # 新增：定义_center_window方法来计算并设置窗口位置
    def _center_window(self):
        # 设置样式
        style = ttk.Style()
        style.configure("TButton", font=("微软雅黑", 12), padding=10)
        style.configure("TFrame", background="#f0f0f0")
        
        # 创建按钮框架
        self.button_frame = ttk.Frame(self.root)
        self.button_frame.pack(pady=20, expand=True, fill=tk.BOTH)
        
        # 创建“选择文件”按钮
        self.file_button = ttk.Button(self.button_frame, text="选择文件", command=self.select_file)
        self.file_button.pack(side=tk.LEFT, padx=10, expand=True, fill=tk.BOTH)
        
        # 创建“选择文件夹”按钮
        self.folder_button = ttk.Button(self.button_frame, text="选择文件夹", command=self.select_folder)
        self.folder_button.pack(side=tk.LEFT, padx=10, expand=True, fill=tk.BOTH)
    
    def select_file(self):
        file_paths = filedialog.askopenfilenames(filetypes=[("Word files", "*.docx")])
        if file_paths:
            for file_path in file_paths:
                print(f"选择的文件: {file_path}")
                # 创建Word文档对象
                word_doc = WordDocument(file_path)
                # 打开文档
                if word_doc.open():
                    # 删除页眉
                    word_doc.remove_headers()
                    # 删除页脚
                    word_doc.remove_footers()
                    # 添加页码
                    word_doc.add_page_numbers()
                    # 保存修改（这里选择另存为，避免覆盖原文件）
                    word_doc.save()
                    print(f"处理完成，新文件已保存至：{word_doc.file_path}")
            # 弹出一次性完成通知
            messagebox.showinfo("完成", "所选文件已全部处理完成。")
    
    def select_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            print(f"选择的文件夹: {folder_path}")
            # 遍历文件夹及其子目录下的所有.docx文件
            for root, dirs, files in os.walk(folder_path):
                for file in files:
                    if file.endswith(".docx"):
                        file_path = os.path.join(root, file)
                        print(f"正在处理文件: {file_path}")
                        # 创建Word文档对象
                        word_doc = WordDocument(file_path)
                        # 打开文档
                        if word_doc.open():
                            # 删除页眉
                            word_doc.remove_headers()
                            # 删除页脚
                            word_doc.remove_footers()
                            # 添加页码
                            word_doc.add_page_numbers()
                            # 保存修改（这里选择另存为，避免覆盖原文件）
                            word_doc.save()
                            print(f"处理完成，新文件已保存至：{word_doc.file_path}")
            # 弹出一次性完成通知
            messagebox.showinfo("完成", "文件夹内所有文件已全部处理完成。")

if __name__ == "__main__":
    root = tk.Tk()
    app = MainWindow(root)
    root.mainloop()
