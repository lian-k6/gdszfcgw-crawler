import os
import tkinter as tk
from tkinter import filedialog, scrolledtext, ttk
import sys


class FileNameTool:
    def __init__(self, root):
        self.root = root
        self.root.title("文件名提取与批量重命名工具")
        self.root.geometry("700x600")
        self.root.resizable(True, True)

        # 设置样式
        self.style = ttk.Style()
        self.style.configure("TButton", font=("微软雅黑", 10))
        self.style.configure("TLabel", font=("微软雅黑", 10))
        self.style.configure("TEntry", font=("微软雅黑", 10))

        # 创建界面组件
        self.create_widgets()

    def create_widgets(self):
        # 源文件夹选择
        frame_source = ttk.Frame(self.root, padding="10")
        frame_source.pack(fill=tk.X, padx=10)

        ttk.Label(frame_source, text="源文件夹（提取文件名）:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.entry_source = ttk.Entry(frame_source)
        self.entry_source.grid(row=0, column=1, sticky=tk.EW, pady=5, padx=5)
        ttk.Button(frame_source, text="浏览...", command=self.browse_source).grid(row=0, column=2, padx=5)

        frame_source.columnconfigure(1, weight=1)

        # TXT文件路径选择
        frame_txt = ttk.Frame(self.root, padding="10")
        frame_txt.pack(fill=tk.X, padx=10)

        ttk.Label(frame_txt, text="TXT文件路径（保存/读取）:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.entry_txt = ttk.Entry(frame_txt)
        self.entry_txt.grid(row=0, column=1, sticky=tk.EW, pady=5, padx=5)
        ttk.Button(frame_txt, text="浏览...", command=self.browse_txt).grid(row=0, column=2, padx=5)

        frame_txt.columnconfigure(1, weight=1)

        # 目标文件夹选择
        frame_target = ttk.Frame(self.root, padding="10")
        frame_target.pack(fill=tk.X, padx=10)

        ttk.Label(frame_target, text="目标文件夹（需重命名）:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.entry_target = ttk.Entry(frame_target)
        self.entry_target.grid(row=0, column=1, sticky=tk.EW, pady=5, padx=5)
        ttk.Button(frame_target, text="浏览...", command=self.browse_target).grid(row=0, column=2, padx=5)

        frame_target.columnconfigure(1, weight=1)

        # 操作按钮
        frame_buttons = ttk.Frame(self.root, padding="10")
        frame_buttons.pack(fill=tk.X, padx=10)

        ttk.Button(frame_buttons, text="提取文件名并写入TXT", command=self.extract_and_write).pack(side=tk.LEFT, padx=5)
        ttk.Button(frame_buttons, text="从TXT读取并批量重命名", command=self.read_and_rename).pack(side=tk.LEFT, padx=5)
        ttk.Button(frame_buttons, text="清空日志", command=self.clear_log).pack(side=tk.RIGHT, padx=5)

        # 日志显示区域
        frame_log = ttk.Frame(self.root, padding="10")
        frame_log.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        ttk.Label(frame_log, text="操作日志:").pack(anchor=tk.W)
        self.log_area = scrolledtext.ScrolledText(frame_log, wrap=tk.WORD, height=18, font=("微软雅黑", 9))
        self.log_area.pack(fill=tk.BOTH, expand=True, pady=5)
        self.log_area.config(state=tk.DISABLED)

        # 重定向控制台输出到日志区域
        self.redirect_stdout()

    def redirect_stdout(self):
        """将print输出重定向到日志区域"""

        class TextRedirector:
            def __init__(self, text_widget):
                self.text_widget = text_widget

            def write(self, string):
                self.text_widget.config(state=tk.NORMAL)
                self.text_widget.insert(tk.END, string)
                self.text_widget.see(tk.END)  # 自动滚动到底部
                self.text_widget.config(state=tk.DISABLED)

            def flush(self):
                pass

        sys.stdout = TextRedirector(self.log_area)

    def browse_source(self):
        folder = filedialog.askdirectory(title="选择源文件夹")
        if folder:
            self.entry_source.delete(0, tk.END)
            self.entry_source.insert(0, folder)

    def browse_txt(self):
        file = filedialog.asksaveasfilename(
            title="选择或创建TXT文件",
            defaultextension=".txt",
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")]
        )
        if file:
            self.entry_txt.delete(0, tk.END)
            self.entry_txt.insert(0, file)

    def browse_target(self):
        folder = filedialog.askdirectory(title="选择目标文件夹")
        if folder:
            self.entry_target.delete(0, tk.END)
            self.entry_target.insert(0, folder)

    def get_file_names(self, folder_path):
        """获取文件夹中所有文件的文件名（不含子文件夹），并打印具体文件"""
        if not os.path.exists(folder_path):
            print(f"错误：文件夹 '{folder_path}' 不存在")
            return []

        file_names = []
        print(f"正在扫描文件夹: {folder_path}")
        print("发现的文件:")
        for item in os.listdir(folder_path):
            item_path = os.path.join(folder_path, item)
            try:
                # 检查是否为文件（排除文件夹）
                if os.path.isfile(item_path):
                    file_names.append(item)
                    print(f"  - {item}")  # 打印具体文件名
                else:
                    print(f"  跳过文件夹: {item}")
            except Exception as e:
                print(f"  访问{item}失败（可能无权限）: {str(e)}")

        # 按文件名排序
        file_names.sort()
        print(f"共提取到 {len(file_names)} 个文件")
        return file_names

    def write_names_to_txt(self, file_names, txt_path):
        """将文件名列表写入TXT文件，并验证写入内容"""
        try:
            with open(txt_path, 'w', encoding='utf-8') as f:
                for name in file_names:
                    f.write(f"{name}\n")
            print(f"\n成功写入TXT文件: {txt_path}")
            print("TXT文件内容:")
            with open(txt_path, 'r', encoding='utf-8') as f:
                for line in f:
                    print(f"  {line.strip()}")  # 打印TXT中的内容
            return True
        except Exception as e:
            print(f"写入TXT失败：{str(e)}")
            return False

    def read_names_from_txt(self, txt_path):
        """从TXT文件读取文件名列表，并打印读取结果"""
        if not os.path.exists(txt_path):
            print(f"错误：TXT文件 '{txt_path}' 不存在")
            return []

        try:
            with open(txt_path, 'r', encoding='utf-8') as f:
                names = [line.strip() for line in f.readlines() if line.strip()]
            print(f"\n从TXT读取到的文件名:")
            for name in names:
                print(f"  - {name}")  # 打印读取到的文件名
            print(f"共读取到 {len(names)} 个文件名")
            return names
        except Exception as e:
            print(f"读取TXT失败：{str(e)}")
            return []

    def rename_target_files(self, target_folder, new_names):
        """用新文件名列表重命名目标文件夹中的文件，打印目标文件情况"""
        if not os.path.exists(target_folder):
            print(f"错误：目标文件夹 '{target_folder}' 不存在")
            return

        # 获取目标文件夹中的文件并打印
        target_files = []
        print(f"\n正在扫描目标文件夹: {target_folder}")
        print("目标文件夹中的文件:")
        for item in os.listdir(target_folder):
            item_path = os.path.join(target_folder, item)
            try:
                if os.path.isfile(item_path):
                    target_files.append(item)
                    print(f"  - {item}")  # 打印目标文件
                else:
                    print(f"  跳过文件夹: {item}")
            except Exception as e:
                print(f"  访问{item}失败（可能无权限）: {str(e)}")

        target_files.sort()
        print(f"目标文件夹共发现 {len(target_files)} 个文件")

        # 检查数量匹配
        if len(target_files) != len(new_names):
            print(f"错误：目标文件数量（{len(target_files)}）与新文件名数量（{len(new_names)}）不匹配")
            return

        # 批量重命名
        success_count = 0
        print("\n开始重命名:")
        for old_name, new_name in zip(target_files, new_names):
            old_path = os.path.join(target_folder, old_name)
            new_path = os.path.join(target_folder, new_name)

            if old_name == new_name:
                print(f"  跳过：{old_name}（与新名相同）")
                success_count += 1
                continue

            try:
                os.rename(old_path, new_path)
                print(f"  成功：{old_name} -> {new_name}")
                success_count += 1
            except Exception as e:
                print(f"  失败（{old_name}）：{str(e)}")

        print(f"\n重命名完成，成功 {success_count}/{len(target_files)} 个文件")

    def extract_and_write(self):
        """提取文件名并写入TXT"""
        source_folder = self.entry_source.get().strip()
        txt_path = self.entry_txt.get().strip()

        if not source_folder:
            print("请选择源文件夹")
            return

        if not txt_path:
            print("请选择TXT文件路径")
            return

        print("\n" + "=" * 40)
        print("===== 开始提取文件名并写入TXT =====")
        source_names = self.get_file_names(source_folder)
        if source_names:
            self.write_names_to_txt(source_names, txt_path)
        print("===== 操作结束 =====")
        print("=" * 40 + "\n")

    def read_and_rename(self):
        """从TXT读取并批量重命名"""
        txt_path = self.entry_txt.get().strip()
        target_folder = self.entry_target.get().strip()

        if not txt_path:
            print("请选择TXT文件路径")
            return

        if not target_folder:
            print("请选择目标文件夹")
            return

        print("\n" + "=" * 40)
        print("===== 开始从TXT读取并批量重命名 =====")
        new_names = self.read_names_from_txt(txt_path)
        if new_names:
            self.rename_target_files(target_folder, new_names)
        print("===== 操作结束 =====")
        print("=" * 40 + "\n")

    def clear_log(self):
        """清空日志区域"""
        self.log_area.config(state=tk.NORMAL)
        self.log_area.delete(1.0, tk.END)
        self.log_area.config(state=tk.DISABLED)


if __name__ == "__main__":
    root = tk.Tk()
    app = FileNameTool(root)
    root.mainloop()