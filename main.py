import logging
import os
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from tkinterdnd2 import DND_FILES, TkinterDnD


class FigConverter(TkinterDnD.Tk):
    """
    主应用程序类 - 处理图像格式转换
    """
    def __init__(self):
        super().__init__()
        
        # 设置窗口属性
        self.title("图片格式转换工具")
        self.geometry("600x500")
        self.minsize(500, 400)
        
        # 支持的输出文件类型
        self.file_types = {
            "PNG": "png",
            # "JPEG": "jpg",  # 已移除JPEG输出格式
            "SVG": "svg",   # 已移除SVG输出格式
            "PDF": "pdf",
            "EPS": "eps",
            "EMF": "emf",
        }
        
        # 定义位图和矢量图格式
        self.bitmap_formats = ['.png', '.jpg', '.jpeg', '.tiff', '.bmp', '.gif']
        self.vector_formats = ['.svg', '.pdf', '.eps', '.ps', '.emf']
        
        # 定义输出格式类型（位图/矢量图）
        self.output_format_types = {
            'PNG': 'bitmap',
            'TIFF': 'bitmap',
            'SVG': 'vector',
            'PDF': 'vector',
            'EPS': 'vector',
            'EMF': 'vector'
        }
        
        # 选择的输出文件类型
        self.selected_types = {}
        
        # 要转换的文件列表
        self.files_to_convert = []
        
        # DPI设置，默认为300
        self.dpi_value = tk.IntVar(value=300)
        
        # 创建GUI组件
        self._create_widgets()
        
        # 设置日志
        self._setup_logging()
        
        # 检查Inkscape
        self.inkscape_path = None
        self._check_inkscape()
    
    def _create_widgets(self):
        """创建GUI组件"""
        # 创建主框架
        main_frame = ttk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 创建顶部区域 - 输出格式选择
        format_outer_frame = ttk.Frame(main_frame)
        format_outer_frame.pack(fill=tk.X, padx=5, pady=5)
        
        format_frame = ttk.LabelFrame(format_outer_frame, text="选择输出格式")
        format_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 添加重置格式按钮到右侧
        reset_format_button = ttk.Button(
            format_outer_frame, 
            text="重置格式", 
            command=self._reset_format_options
        )
        reset_format_button.pack(side=tk.RIGHT, padx=5, pady=5)
        
        # 在网格中添加复选框
        row, col = 0, 0
        for file_type in self.file_types:
            var = tk.BooleanVar(value=False)
            self.selected_types[file_type] = var
            chk = ttk.Checkbutton(format_frame, text=file_type, variable=var, command=self._update_button_state)
            chk.grid(row=row, column=col, sticky=tk.W, padx=5, pady=2)
            col += 1
            if col > 4:  # 每行显示4个选项
                col = 0
                row += 1
        
        # 创建DPI设置框架
        dpi_frame = ttk.LabelFrame(main_frame, text="位图DPI设置 (仅对PNG、TIFF等位图格式生效)")
        dpi_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # 添加DPI滑动条
        ttk.Label(dpi_frame, text="DPI:").grid(row=0, column=0, padx=5, pady=5)
        dpi_scale = ttk.Scale(dpi_frame, from_=72, to=600, variable=self.dpi_value, 
                             orient=tk.HORIZONTAL, length=200)
        dpi_scale.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W+tk.E)
        
        # 添加DPI数值显示和输入框
        dpi_entry = ttk.Entry(dpi_frame, textvariable=self.dpi_value, width=5)
        dpi_entry.grid(row=0, column=2, padx=5, pady=5)
        ttk.Label(dpi_frame, text="(72-600)").grid(row=0, column=3, padx=5, pady=5, sticky=tk.W)
        
        # 创建拖放区域
        drop_frame = ttk.LabelFrame(main_frame, text="拖拽文件到此处")
        drop_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 创建一个可以显示文件列表的文本区域
        self.file_list = ScrolledText(drop_frame, height=10)
        self.file_list.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.file_list.insert(tk.END, "请拖拽文件到此处或点击\"添加文件\"按钮...\n")
        self.file_list.config(state=tk.DISABLED)
        
        # 配置拖拽事件
        self.file_list.drop_target_register(DND_FILES)
        self.file_list.dnd_bind('<<Drop>>', self._on_drop)
        
        # 创建按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # 添加按钮
        ttk.Button(button_frame, text="添加文件", command=self._add_files).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="清除列表", command=self._clear_files).pack(side=tk.LEFT, padx=5)
        
        # 转换按钮
        self.convert_button = ttk.Button(button_frame, text="开始转换", command=self._start_conversion)
        self.convert_button.pack(side=tk.RIGHT, padx=5)
        
        # 初始状态下禁用转换按钮
        self.convert_button.config(state=tk.DISABLED)
        
        # 创建状态栏
        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(self, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 创建进度条
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, padx=5, pady=5)
    
    def _setup_logging(self):
        """设置日志"""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            handlers=[
                logging.StreamHandler(),
                logging.FileHandler("fig_converter.log")
            ]
        )
        self.logger = logging.getLogger("FigConverter")
    
    def _check_inkscape(self):
        """检查是否安装了Inkscape"""
        self.status_var.set("检查Inkscape安装...")
        self.logger.info("检查Inkscape安装")
        
        try:
            # 尝试查找Inkscape可执行文件
            if os.name == 'nt':  # Windows
                # 检查常见的安装路径
                common_paths = [
                    r"C:\Program Files\Inkscape\bin\inkscape.exe",
                    r"C:\Program Files (x86)\Inkscape\bin\inkscape.exe",
                    r"C:\Program Files\Inkscape\inkscape.exe",
                    r"C:\Program Files (x86)\Inkscape\inkscape.exe"
                ]
                
                # 检查环境变量PATH中的inkscape
                try:
                    result = subprocess.run(["where", "inkscape"], 
                                           capture_output=True, 
                                           text=True, 
                                           encoding='utf-8',
                                           check=True)
                    paths = result.stdout.strip().split('\n')
                    if paths:
                        for path in paths:
                            if os.path.exists(path):
                                self.inkscape_path = path
                                self.logger.info(f"从PATH中找到Inkscape: {path}")
                                self.status_var.set(f"已找到Inkscape: {path}")
                                return
                except subprocess.CalledProcessError:
                    self.logger.warning("在PATH中未找到Inkscape")
                
                # 检查常见路径
                for path in common_paths:
                    if os.path.exists(path):
                        self.inkscape_path = path
                        self.logger.info(f"找到Inkscape: {path}")
                        self.status_var.set(f"已找到Inkscape: {path}")
                        return
            else:  # Linux/Mac
                try:
                    result = subprocess.run(["which", "inkscape"], 
                                          capture_output=True, 
                                          text=True, 
                                          encoding='utf-8',
                                          check=True)
                    path = result.stdout.strip()
                    if path:
                        self.inkscape_path = path
                        self.logger.info(f"找到Inkscape: {path}")
                        self.status_var.set(f"已找到Inkscape: {path}")
                        return
                except subprocess.CalledProcessError:
                    self.logger.warning("未找到Inkscape")
            
            # 如果没有找到Inkscape
            self.logger.error("未找到Inkscape，请确保已安装")
            messagebox.showerror(
                "错误", 
                "未找到Inkscape。请安装Inkscape并确保其在系统PATH中，或手动选择Inkscape可执行文件。"
            )
            self._select_inkscape_manually()
            
        except Exception as e:
            self.logger.error(f"检查Inkscape时出错: {str(e)}")
            self.status_var.set("检查Inkscape失败，请手动选择")
            self._select_inkscape_manually()
    
    def _select_inkscape_manually(self):
        """手动选择Inkscape可执行文件"""
        self.logger.info("请求用户手动选择Inkscape可执行文件")
        if os.name == 'nt':  # Windows
            file_types = [("Inkscape 可执行文件", "inkscape.exe"), ("所有文件", "*.*")]
        else:  # Linux/Mac
            file_types = [("Inkscape 可执行文件", "inkscape"), ("所有文件", "*.*")]
        
        path = filedialog.askopenfilename(
            title="选择Inkscape可执行文件",
            filetypes=file_types
        )
        
        if path:
            self.inkscape_path = path
            self.logger.info(f"手动选择的Inkscape路径: {path}")
            self.status_var.set(f"已设置Inkscape: {path}")
        else:
            self.logger.warning("用户取消了Inkscape选择")
            self.status_var.set("未设置Inkscape路径，部分功能可能不可用")
    
    def _on_drop(self, event):
        """处理文件拖放事件"""
        # 解析拖放的文件路径
        files = self._parse_drop_data(event.data)
        if files:
            self._add_files_to_list(files)
    
    def _parse_drop_data(self, data):
        """解析拖放数据，提取文件路径列表"""
        # 如果是Windows，需要处理花括号和引号
        if os.name == 'nt':
            # 移除可能的花括号
            if data.startswith('{') and data.endswith('}'):
                data = data[1:-1]
            # 分割多个文件（如果有的话）
            files = []
            for item in data.split('} {'):
                # 处理每个路径，移除引号
                item = item.strip('"')
                files.append(item)
            return files
        else:
            # Linux/Mac格式
            return data.split()
        
    def _add_files_to_list(self, file_paths):
        """添加文件到列表中"""
        # 添加合法的图像文件到转换列表
        valid_extensions = ['.svg', '.png', '.jpg', '.jpeg', '.pdf', '.eps', '.ps', '.emf', '.tiff', '.bmp', '.gif']
        added_files = []
        
        # 不再在添加文件时自动重置格式选项
        
        for file_path in file_paths:
            path = Path(file_path)
            if path.is_file() and path.suffix.lower() in valid_extensions:
                if file_path not in self.files_to_convert:
                    self.files_to_convert.append(file_path)
                    added_files.append(file_path)
                    # 根据文件类型禁用不兼容的输出格式
                    self._update_format_options(path.suffix.lower())
            else:
                self.logger.warning(f"不支持的文件类型或文件不存在: {file_path}")
        
        # 更新UI显示
        if added_files:
            self.file_list.config(state=tk.NORMAL)
            # 如果这是第一个文件，清除提示文本
            if self.file_list.get(1.0, tk.END).strip() == "请拖拽文件到此处或点击\"添加文件\"按钮...":
                self.file_list.delete(1.0, tk.END)
            
            # 添加新文件到列表
            for file_path in added_files:
                self.file_list.insert(tk.END, f"{file_path}\n")
            
            self.file_list.config(state=tk.DISABLED)
            self.status_var.set(f"已添加 {len(added_files)} 个文件，共 {len(self.files_to_convert)} 个文件待转换")
            
            # 更新按钮状态
            self._update_button_state()
        
    def _add_files(self):
        """添加文件到转换列表"""
        # 打开文件选择对话框
        file_types = [
            ("支持的图像文件", "*.svg *.png *.jpg *.jpeg *.pdf *.eps *.ps *.emf"),
            ("SVG文件", "*.svg"),
            ("PNG文件", "*.png"),
            ("JPEG文件", "*.jpg *.jpeg"),
            ("PDF文件", "*.pdf"),
            ("EPS文件", "*.eps"),
            ("PS文件", "*.ps"),
            ("EMF文件", "*.emf"),
            ("所有文件", "*.*")
        ]
        
        files = filedialog.askopenfilenames(
            title="选择要转换的图像文件",
            filetypes=file_types
        )
        
        if files:
            self._add_files_to_list(files)
    
    def _clear_files(self):
        """清除文件列表"""
        self.files_to_convert = []
        self.file_list.config(state=tk.NORMAL)
        self.file_list.delete(1.0, tk.END)
        self.file_list.insert(tk.END, "请拖拽文件到此处或点击\"添加文件\"按钮...\n")
        self.file_list.config(state=tk.DISABLED)
        self.status_var.set("文件列表已清除")
        
        # 更新按钮状态
        self._update_button_state()
    
    def _reset_format_options(self):
        """重置所有格式选项为可选状态"""
        for format_name in self.file_types:
            # 获取对应的复选框
            self.selected_types[format_name].set(False)
            
            # 遍历主窗口中的所有复选框，找到对应的复选框并启用它
            for child in self.winfo_children():
                if isinstance(child, ttk.Frame):
                    for sub_child in child.winfo_children():
                        if isinstance(sub_child, ttk.Frame):
                            for inner_child in sub_child.winfo_children():
                                if isinstance(inner_child, ttk.LabelFrame) and inner_child["text"] == "选择输出格式":
                                    for widget in inner_child.winfo_children():
                                        if isinstance(widget, ttk.Checkbutton) and widget["text"] == format_name:
                                            widget.config(state=tk.NORMAL)
        
        # 更新转换按钮状态
        self._update_button_state()
        
        # 显示状态信息
        self.status_var.set("格式选项已重置")
    
    def _update_format_options(self, file_extension):
        """根据文件类型禁用不兼容的输出格式"""
        # 判断文件是位图还是矢量图
        is_bitmap = file_extension.lower() in self.bitmap_formats
        
        # 对于位图文件，禁用矢量图输出格式
        if is_bitmap:
            for format_name, format_type in self.output_format_types.items():
                if format_type == 'vector':
                    # 找到并禁用对应的复选框
                    for child in self.winfo_children():
                        if isinstance(child, ttk.Frame):
                            for sub_child in child.winfo_children():
                                if isinstance(sub_child, ttk.LabelFrame) and sub_child["text"] == "选择输出格式":
                                    for widget in sub_child.winfo_children():
                                        if isinstance(widget, ttk.Checkbutton) and widget["text"] == format_name:
                                            # 禁用复选框并取消选择
                                            widget.config(state=tk.DISABLED)
                                            self.selected_types[format_name].set(False)
        
        # DPI设置框架始终显示，不再根据文件类型隐藏
    
    def _update_button_state(self):
        """更新转换按钮状态"""
        # 检查是否有选择的输出格式
        has_selected_format = any(var.get() for var in self.selected_types.values())
        
        # 启用或禁用转换按钮
        if has_selected_format:
            self.convert_button.config(state=tk.NORMAL)
        else:
            self.convert_button.config(state=tk.DISABLED)
    
    def _start_conversion(self):
        """开始转换流程"""
        # 检查是否有文件要转换
        if not self.files_to_convert:
            messagebox.showwarning("警告", "没有选择要转换的文件")
            return
        
        # 检查是否选择了输出格式
        selected_formats = [fmt for fmt, var in self.selected_types.items() if var.get()]
        if not selected_formats:
            messagebox.showwarning("警告", "请选择至少一种输出格式")
            return
        
        # 检查Inkscape是否可用
        if not self.inkscape_path:
            messagebox.showerror("错误", "Inkscape路径未设置，无法执行转换")
            return
            
        # 创建一个新线程执行转换，以免阻塞UI
        conversion_thread = threading.Thread(
            target=self._execute_conversion,
            args=(self.files_to_convert.copy(), selected_formats)
        )
        conversion_thread.daemon = True
        conversion_thread.start()
        
        self.status_var.set("开始转换...")
        self.progress_var.set(0)
    
    def _execute_conversion(self, files, formats):
        """执行实际的文件转换"""
        try:
            # 首先计算实际需要转换的任务数量（排除相同格式）
            total_tasks = 0
            skipped_tasks = []
            
            for file_path in files:
                input_file = Path(file_path)
                for format_name in formats:
                    format_extension = self.file_types[format_name]
                    if input_file.suffix.lower() == f".{format_extension}":
                        skipped_tasks.append((file_path, format_name))
                    else:
                        total_tasks += 1
            
            total_files = len(files)
            total_formats = len(formats)
            completed_tasks = 0
            
            if len(skipped_tasks) > 0:
                self.logger.info(f"跳过 {len(skipped_tasks)} 个相同格式的转换任务")
            
            # 如果没有实际需要转换的任务
            if total_tasks == 0:
                self.logger.info("没有需要转换的任务，所有选择的格式与源文件格式相同")
                self.status_var.set("没有需要转换的任务")
                self.progress_var.set(100)
                self.after(0, lambda: messagebox.showinfo("完成", "没有需要转换的任务，所有选择的格式与源文件格式相同"))
                return
                
            self.logger.info(f"开始转换 {total_files} 个文件到 {total_formats} 种格式，实际执行 {total_tasks} 个任务")
            
            for file_path in files:
                input_file = Path(file_path)
                base_name = input_file.stem
                output_dir = input_file.parent
                
                for format_name in formats:
                    try:
                        format_extension = self.file_types[format_name]
                        output_file = output_dir / f"{base_name}.{format_extension}"
                        
                        # 检查是否与源文件格式相同
                        if input_file.suffix.lower() == f".{format_extension}":
                            self.logger.info(f"跳过相同格式转换: {input_file.name} 已经是 {format_name} 格式")
                            continue
                        
                        # 更新状态
                        status_msg = f"转换 {input_file.name} 到 {format_name}..."
                        self.status_var.set(status_msg)
                        self.logger.info(status_msg)
                        
                        # 构建Inkscape命令
                        cmd = [
                            self.inkscape_path,
                            str(input_file),
                            f"--export-type={format_extension}",
                            f"--export-filename={str(output_file)}"
                        ]
                        
                        # 如果是导出位图格式，添加DPI设置
                        if format_extension in ['png', 'tiff']:
                            cmd.append(f"--export-dpi={self.dpi_value.get()}")
                        
                        # 执行命令
                        self.logger.debug(f"执行命令: {' '.join(cmd)}")
                        # 修改：不使用text=True，而是手动处理二进制输出
                        result = subprocess.run(cmd, capture_output=True, text=False)
                        
                        if result.returncode == 0:
                            self.logger.info(f"成功转换: {output_file}")
                        else:
                            # 尝试用UTF-8解码错误信息，如果失败则使用'replace'策略来替换无法解码的字符
                            stderr = result.stderr.decode('utf-8', errors='replace') if result.stderr else ""
                            self.logger.error(f"转换失败: {stderr}")
                            # 在UI线程中显示错误信息
                            stderr = result.stderr.decode('utf-8', errors='replace') if result.stderr else ""
                            self.after(0, lambda stderr=stderr, fname=input_file.name, fmt=format_name: messagebox.showerror(
                                "转换错误", 
                                f"转换 {fname} 到 {fmt} 时出错:\n{stderr}"
                            ))
                    except Exception as e:
                        self.logger.error(f"转换 {input_file.name} 到 {format_name} 时出错: {str(e)}")
                    finally:
                        # 更新进度
                        completed_tasks += 1
                        if total_tasks > 0:  # 避免除以零
                            progress = (completed_tasks / total_tasks) * 100
                            self.progress_var.set(progress)
            
            # 完成所有转换后
            self.logger.info("所有转换任务已完成")
            self.status_var.set("转换完成")
            # 显示完成消息
            self.after(0, lambda: messagebox.showinfo("完成", "所有文件已转换完成"))
            
        except Exception as e:
            self.logger.error(f"转换过程中出错: {str(e)}")
            self.status_var.set("转换过程中出错")
            self.after(0, lambda: messagebox.showerror("错误", f"转换过程中出错: {str(e)}"))
        finally:
            # 确保进度条完成
            self.progress_var.set(100)


def main():
    app = FigConverter()
    app.mainloop()


if __name__ == "__main__":
    main()
