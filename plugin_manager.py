import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox
import os
import importlib.util
import sys
from pathlib import Path
import json


class PluginManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("插件管理应用")
        self.root.geometry("800x600")
        self.root.minsize(600, 400)

        # 插件目录设置
        self.plugins_dir = "plugins"
        self.plugin_config_file = "plugin_config.json"

        # 确保插件目录存在
        if not os.path.exists(self.plugins_dir):
            os.makedirs(self.plugins_dir)

        # 加载插件配置
        self.plugin_config = self.load_plugin_config()

        # 创建UI
        self.create_ui()

        # 加载插件
        self.load_plugins()

    def create_ui(self):
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 创建左侧插件列表框架
        left_frame = ttk.LabelFrame(main_frame, text="插件列表", padding="10")
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=False, padx=(0, 10))

        # 插件列表
        self.plugin_listbox = tk.Listbox(left_frame, width=30, height=25)
        self.plugin_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.plugin_listbox.bind('<<ListboxSelect>>', self.on_plugin_select)

        # 滚动条
        scrollbar = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=self.plugin_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.plugin_listbox.config(yscrollcommand=scrollbar.set)

        # 插件操作按钮
        button_frame = ttk.Frame(left_frame, padding="5")
        button_frame.pack(fill=tk.X, pady=10)

        ttk.Button(button_frame, text="添加插件", command=self.add_plugin).pack(fill=tk.X, pady=2)
        ttk.Button(button_frame, text="刷新列表", command=self.refresh_plugins).pack(fill=tk.X, pady=2)
        ttk.Button(button_frame, text="插件设置", command=self.plugin_settings).pack(fill=tk.X, pady=2)

        # 创建右侧插件运行区域
        right_frame = ttk.Frame(main_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # 插件信息区域
        self.plugin_info_frame = ttk.LabelFrame(right_frame, text="插件信息", padding="10")
        self.plugin_info_frame.pack(fill=tk.X, pady=(0, 10))

        self.plugin_name_label = ttk.Label(self.plugin_info_frame, text="未选择插件")
        self.plugin_name_label.pack(anchor=tk.W, pady=(0, 5))

        self.plugin_desc_label = ttk.Label(self.plugin_info_frame, text="", wraplength=500)
        self.plugin_desc_label.pack(anchor=tk.W)

        # 插件输出区域
        output_frame = ttk.LabelFrame(right_frame, text="插件输出", padding="10")
        output_frame.pack(fill=tk.BOTH, expand=True)

        self.output_text = scrolledtext.ScrolledText(output_frame, wrap=tk.WORD, state=tk.DISABLED)
        self.output_text.pack(fill=tk.BOTH, expand=True)

        # 插件控制按钮
        control_frame = ttk.Frame(right_frame, padding="5")
        control_frame.pack(fill=tk.X, pady=10)

        self.run_button = ttk.Button(control_frame, text="运行插件", command=self.run_selected_plugin, state=tk.DISABLED)
        self.run_button.pack(side=tk.RIGHT, padx=5)

        # 创建菜单栏
        self.create_menu()

    def create_menu(self):
        menubar = tk.Menu(self.root)

        # 文件菜单
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="设置插件目录", command=self.set_plugins_directory)
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.root.quit)
        menubar.add_cascade(label="文件", menu=file_menu)

        # 帮助菜单
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="关于", command=self.show_about)
        help_menu.add_command(label="使用帮助", command=self.show_help)
        menubar.add_cascade(label="帮助", menu=help_menu)

        self.root.config(menu=menubar)

    def load_plugin_config(self):
        """加载插件配置文件"""
        if os.path.exists(self.plugin_config_file):
            with open(self.plugin_config_file, 'r') as f:
                try:
                    return json.load(f)
                except json.JSONDecodeError:
                    messagebox.showerror("错误", "插件配置文件损坏，将使用默认配置")
        return {"plugins_dir": self.plugins_dir, "enabled_plugins": {}}

    def save_plugin_config(self):
        """保存插件配置文件"""
        with open(self.plugin_config_file, 'w') as f:
            json.dump(self.plugin_config, f, indent=2)

    def load_plugins(self):
        """加载插件目录中的所有插件"""
        self.plugins = {}
        self.plugin_listbox.delete(0, tk.END)

        if not os.path.exists(self.plugins_dir):
            messagebox.showwarning("警告", f"插件目录不存在: {self.plugins_dir}")
            return

        # 遍历插件目录
        for item in os.listdir(self.plugins_dir):
            item_path = os.path.join(self.plugins_dir, item)

            # 检查是否是Python文件或目录
            if item.endswith('.py') and not item.startswith('__'):
                self._load_python_plugin(item, item_path)
            elif os.path.isdir(item_path) and os.path.exists(os.path.join(item_path, '__init__.py')):
                self._load_package_plugin(item, item_path)

    def _load_python_plugin(self, name, path):
        """加载单个Python文件作为插件"""
        try:
            # 导入模块
            spec = importlib.util.spec_from_file_location(name, path)
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)

            # 检查是否是有效的插件
            if hasattr(module, 'Plugin'):
                plugin_class = module.Plugin
                plugin_instance = plugin_class()

                # 获取插件信息
                plugin_name = getattr(plugin_instance, 'name', name[:-3])
                plugin_desc = getattr(plugin_instance, 'description', "无描述信息")

                # 存储插件
                self.plugins[plugin_name] = {
                    'instance': plugin_instance,
                    'path': path,
                    'description': plugin_desc
                }

                # 添加到列表
                self.plugin_listbox.insert(tk.END, plugin_name)

        except Exception as e:
            self.log_message(f"加载插件 {name} 失败: {str(e)}")

    def _load_package_plugin(self, name, path):
        """加载Python包作为插件"""
        try:
            # 将插件目录添加到系统路径
            sys.path.append(os.path.dirname(path))

            # 导入模块
            module = importlib.import_module(name)

            # 检查是否是有效的插件
            if hasattr(module, 'Plugin'):
                plugin_class = module.Plugin
                plugin_instance = plugin_class()

                # 获取插件信息
                plugin_name = getattr(plugin_instance, 'name', name)
                plugin_desc = getattr(plugin_instance, 'description', "无描述信息")

                # 存储插件
                self.plugins[plugin_name] = {
                    'instance': plugin_instance,
                    'path': path,
                    'description': plugin_desc
                }

                # 添加到列表
                self.plugin_listbox.insert(tk.END, plugin_name)

        except Exception as e:
            self.log_message(f"加载插件 {name} 失败: {str(e)}")
        finally:
            # 从系统路径中移除
            if os.path.dirname(path) in sys.path:
                sys.path.remove(os.path.dirname(path))

    def on_plugin_select(self, event):
        """当选择插件时更新信息"""
        selection = self.plugin_listbox.curselection()
        if not selection:
            return

        index = selection[0]
        plugin_name = self.plugin_listbox.get(index)

        if plugin_name in self.plugins:
            plugin = self.plugins[plugin_name]
            self.plugin_name_label.config(text=f"插件名称: {plugin_name}")
            self.plugin_desc_label.config(text=f"描述: {plugin['description']}")
            self.run_button.config(state=tk.NORMAL)
        else:
            self.run_button.config(state=tk.DISABLED)

    def run_selected_plugin(self):
        """运行选中的插件"""
        selection = self.plugin_listbox.curselection()
        if not selection:
            return

        index = selection[0]
        plugin_name = self.plugin_listbox.get(index)

        if plugin_name in self.plugins:
            self.clear_output()
            self.log_message(f"开始运行插件: {plugin_name}")

            try:
                plugin_instance = self.plugins[plugin_name]['instance']

                # 检查插件是否有run方法，并且接受output_callback参数
                if hasattr(plugin_instance, 'run'):
                    # 运行插件，并传入日志回调函数
                    result = plugin_instance.run(output_callback=self.log_message)
                    if result is not None:
                        self.log_message(f"插件运行结果: {result}")
                    self.log_message(f"插件 {plugin_name} 运行完成")
                else:
                    self.log_message(f"错误: 插件 {plugin_name} 没有实现run方法")

            except Exception as e:
                self.log_message(f"插件运行出错: {str(e)}")

    def log_message(self, message):
        """在输出区域显示消息"""
        self.output_text.config(state=tk.NORMAL)
        self.output_text.insert(tk.END, message + "\n")
        self.output_text.see(tk.END)
        self.output_text.config(state=tk.DISABLED)
        self.root.update_idletasks()  # 刷新界面

    def clear_output(self):
        """清空输出区域"""
        self.output_text.config(state=tk.NORMAL)
        self.output_text.delete(1.0, tk.END)
        self.output_text.config(state=tk.DISABLED)

    def add_plugin(self):
        """添加新插件"""
        file_paths = filedialog.askopenfilenames(
            title="选择插件文件",
            filetypes=[("Python files", "*.py"), ("All files", "*.*")]
        )

        if not file_paths:
            return

        for file_path in file_paths:
            try:
                # 获取文件名
                file_name = os.path.basename(file_path)
                dest_path = os.path.join(self.plugins_dir, file_name)

                # 复制文件到插件目录
                with open(file_path, 'rb') as src, open(dest_path, 'wb') as dst:
                    dst.write(src.read())

                self.log_message(f"插件 {file_name} 添加成功")

            except Exception as e:
                self.log_message(f"添加插件失败: {str(e)}")

        # 刷新插件列表
        self.refresh_plugins()

    def refresh_plugins(self):
        """刷新插件列表"""
        self.log_message("刷新插件列表...")
        self.load_plugins()
        self.log_message("插件列表刷新完成")

    def plugin_settings(self):
        """插件设置"""
        selection = self.plugin_listbox.curselection()
        if not selection:
            messagebox.showinfo("提示", "请先选择一个插件")
            return

        index = selection[0]
        plugin_name = self.plugin_listbox.get(index)

        if plugin_name in self.plugins:
            plugin_instance = self.plugins[plugin_name]['instance']

            # 检查插件是否有settings方法
            if hasattr(plugin_instance, 'settings'):
                try:
                    plugin_instance.settings()
                    self.log_message(f"打开 {plugin_name} 插件设置")
                except Exception as e:
                    self.log_message(f"打开插件设置出错: {str(e)}")
            else:
                messagebox.showinfo("提示", f"插件 {plugin_name} 没有设置选项")

    def set_plugins_directory(self):
        """设置插件目录"""
        new_dir = filedialog.askdirectory(title="选择插件目录")
        if new_dir:
            self.plugins_dir = new_dir
            self.plugin_config["plugins_dir"] = new_dir
            self.save_plugin_config()
            self.refresh_plugins()
            self.log_message(f"插件目录已设置为: {new_dir}")

    def show_about(self):
        """显示关于信息"""
        messagebox.showinfo(
            "关于",
            "插件管理应用 v1.0\n\n一个用于管理和运行多个小程序插件的桌面应用"
        )

    def show_help(self):
        """显示帮助信息"""
        help_window = tk.Toplevel(self.root)
        help_window.title("使用帮助")
        help_window.geometry("600x400")
        help_window.resizable(True, True)

        help_text = scrolledtext.ScrolledText(help_window, wrap=tk.WORD, state=tk.DISABLED)
        help_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        help_content = """
插件管理应用使用帮助:

1. 添加插件:
   - 点击"添加插件"按钮，选择Python文件(.py)添加到应用中
   - 插件会被复制到当前插件目录

2. 运行插件:
   - 在左侧列表中选择一个插件
   - 点击"运行插件"按钮启动选中的插件
   - 插件输出会显示在右侧输出区域

3. 插件要求:
   - 插件必须是Python文件(.py)
   - 插件中必须定义一个名为Plugin的类
   - Plugin类必须实现run()方法，这是插件的入口点
   - 可选：实现settings()方法用于插件设置
   - 可选：定义name和description属性提供插件信息

4. 插件示例:
   class Plugin:
       name = "我的插件"
       description = "这是一个插件示例"

       def run(self, output_callback=None):
           if output_callback:
               output_callback("插件正在运行...")
           # 插件逻辑代码
           return "运行成功"

       def settings(self):
           # 插件设置逻辑
           pass
        """

        help_text.config(state=tk.NORMAL)
        help_text.insert(tk.END, help_content)
        help_text.config(state=tk.DISABLED)

        ttk.Button(help_window, text="关闭", command=help_window.destroy).pack(pady=10)


if __name__ == "__main__":
    root = tk.Tk()
    app = PluginManagerApp(root)
    root.mainloop()
