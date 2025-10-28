import os
import sys
import shutil
import csv
import time
import configparser
import queue
import threading
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from tkinter import font

# 版本和版权信息
VERSION = "V1.0.2"
COPYRIGHT = "Tobin © 2025"

# 全局队列：用于子线程与GUI主线程通信（传递日志和进度）
log_queue = queue.Queue()
progress_queue = queue.Queue()


def get_root_path():
    """获取程序根目录（exe所在目录，支持PyInstaller打包后路径）"""
    if hasattr(sys, '_MEIPASS'):
        return os.path.dirname(os.path.abspath(sys.executable))
    else:
        return os.path.dirname(os.path.abspath(__file__))


def clean_filename(name):
    """清理文件名：去除特定后缀和标识符，统一转为小写"""
    if name.endswith("L"):
        name = name[:-1]
    if "L(" in name:
        parts = name.split("L(")
        name = parts[0]
    return name.lower()


def load_configuration(config_path):
    """加载配置文件"""
    if not os.path.exists(config_path):
        print(f"🔥 配置文件不存在: {config_path}")
        print("请确保 config.ini 文件与exe在同一目录下")
        return None

    try:
        config = configparser.ConfigParser()
        config.optionxform = lambda option: option  # 保留大小写
        config.read(config_path, encoding='utf-8')

        # 读取路径配置
        drive_letter = config.get("Paths", "drive_letter").strip()
        target_dir_name = config.get("Paths", "target_dir_name")
        original_list_filename = config.get("Paths", "original_list_file")
        log_filename = config.get("Paths", "log_file")

        # 构建路径
        root_path = get_root_path()
        source_dirs = []
        for key in config["SourceDirectories"]:
            relative_path = config.get("SourceDirectories", key)
            full_path = f"{drive_letter}:\\{relative_path}"
            source_dirs.append(full_path)
        
        target_dir = os.path.join(root_path, target_dir_name)
        list_file = os.path.join(root_path, original_list_filename)
        log_file = os.path.join(root_path, log_filename)

        # 读取性能设置
        max_workers = config.getint("Settings", "max_workers", fallback=12)
        retry_attempts = config.getint("Settings", "retry_attempts", fallback=3)

        print(f"✅ 配置加载成功:")
        print(f"   驱动器盘符: {drive_letter}")
        print(f"   源目录数量: {len(source_dirs)}")
        print(f"   目标目录: {target_dir}")
        print(f"   待处理列表: {list_file}")
        print(f"   日志文件: {log_file}")
        print(f"   最大线程数: {max_workers}")
        print(f"   重试次数: {retry_attempts}")

        return {
            "source_dirs": source_dirs,
            "target_dir": target_dir,
            "list_file": list_file,
            "log_file": log_file,
            "max_workers": max_workers,
            "retry_attempts": retry_attempts,
            "original_list_filename": original_list_filename  # 保存原始列表文件名
        }
    except Exception as e:
        print(f"🔥 配置文件解析失败: {str(e)}")
        print("请检查 config.ini 文件格式是否正确")
        return None


def cleanup_target_directory(target_dir):
    """清理目标目录中的.step文件"""
    print("🧹 正在清理目标目录...")
    clean_count = 0

    if not os.path.exists(target_dir):
        print(f"📁 目标目录不存在，将创建: {target_dir}")
        os.makedirs(target_dir, exist_ok=True)
        return

    for file in os.listdir(target_dir):
        if file.lower().endswith(".step"):
            try:
                file_path = os.path.join(target_dir, file)
                os.remove(file_path)
                clean_count += 1
            except Exception as e:
                print(f"⚠️ 删除旧文件失败: {file} - {str(e)}")
    
    print(f"✅ 已清理 {clean_count} 个旧文件")


def build_file_index(source_dirs):
    """构建文件索引"""
    print("⏳ 正在构建全局文件索引...")
    index = defaultdict(list)
    start_time = time.time()
    total_files = 0

    for src_dir in source_dirs:
        try:
            if not os.path.exists(src_dir):
                print(f"⚠️ 路径不存在或无权限访问: {src_dir}")
                continue

            with os.scandir(src_dir) as entries:
                for entry in entries:
                    if entry.is_file() and entry.name.lower().endswith(".step"):
                        base_name = os.path.splitext(entry.name)[0]
                        clean_base = clean_filename(base_name)
                        prefix_key = clean_base[:4] if len(clean_base) >= 4 else clean_base
                        index[prefix_key].append((clean_base, entry.name, src_dir))
                        total_files += 1
        except Exception as e:
            print(f"⚠️ 目录扫描失败: {src_dir} - {str(e)}")

    index_time = time.time() - start_time
    print(f"✅ 索引构建完成: {len(index)} 个前缀组, {total_files} 个文件, 耗时 {index_time:.2f}秒")
    return index


def read_original_file_list(list_file):
    """读取待处理文件列表"""
    try:
        with open(list_file, "r", encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            all_lines = [row[0].strip() for row in reader if row and row[0].strip()]
        
        print(f"📋 待处理文件数: {len(all_lines)}")
        return all_lines
    except Exception as e:
        print(f"🔥 CSV 文件读取失败: {str(e)}")
        print(f"请检查文件是否存在且格式正确: {list_file}")
        return None


def process_item(item, target_dir, index, retry_attempts):
    """处理单个文件复制"""
    original_name, search_name = item
    dst_file = os.path.join(target_dir, f"{original_name}.STEP")
    prefix_key = search_name[:4] if len(search_name) >= 4 else search_name

    if prefix_key in index:
        for clean_base, src_filename, src_dir in index[prefix_key]:
            if clean_base.startswith(search_name):
                for attempt in range(retry_attempts):
                    try:
                        src_path = os.path.join(src_dir, src_filename)
                        shutil.copy2(src_path, dst_file)
                        return {
                            "status": "success",
                            "original": original_name,
                            "copied": src_filename,
                            "source": src_dir
                        }
                    except Exception as e:
                        if attempt < retry_attempts - 1:
                            time.sleep(2 **attempt)
                        else:
                            return {
                                "status": "error",
                                "original": original_name,
                                "copied": f"复制失败: {str(e)}",
                                "source": src_dir
                            }

    return {
        "status": "not_found",
        "original": original_name,
        "copied": "未找到",
        "source": ""
    }


def write_result_log(log_file, result_log):
    """写入日志文件"""
    print("📝 正在写入日志文件...")
    try:
        with open(log_file, "w", encoding="utf-8", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["原始文件名", "实际复制文件名", "来源路径"])
            for res in result_log:
                writer.writerow([res["original"], res["copied"], res["source"]])
        print(f"✅ 日志已保存至: {log_file}")
        return True
    except Exception as e:
        print(f"⚠️ 日志文件写入失败: {str(e)}")
        return False


def worker(config, progress_callback):
    """后台工作线程：执行完整的复制流程"""
    program_start_time = time.time()
    result_log = []
    found_count = 0
    not_found_count = 0
    copy_errors = 0

    if not config:
        progress_queue.put(("complete", False))
        return

    source_dirs = config["source_dirs"]
    target_dir = config["target_dir"]
    list_file = config["list_file"]
    log_file = config["log_file"]
    max_workers = config["max_workers"]
    retry_attempts = config["retry_attempts"]

    # 2. 清理目标目录
    cleanup_target_directory(target_dir)

    # 3. 构建文件索引
    index = build_file_index(source_dirs)

    # 4. 读取待处理列表
    original_files = read_original_file_list(list_file)
    if not original_files or len(original_files) == 0:
        print("🔥 无待处理文件，程序退出")
        progress_queue.put(("complete", False))
        return
    total_files = len(original_files)
    progress_queue.put(("max", total_files))  # 设置进度条最大值

    # 5. 预处理搜索名
    search_items = [(orig, clean_filename(orig)) for orig in original_files]

    # 6. 多线程复制
    print("📦 开始并行复制文件...")
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = [executor.submit(process_item, item, target_dir, index, retry_attempts) 
                  for item in search_items]
        
        for idx, future in enumerate(as_completed(futures)):
            result = future.result()
            result_log.append(result)

            # 更新统计
            if result["status"] == "success":
                found_count += 1
            elif result["status"] == "not_found":
                not_found_count += 1
            elif result["status"] == "error":
                copy_errors += 1

            # 更新进度（+1表示当前完成数）
            progress_queue.put(("update", idx + 1))

    # 7. 写入日志
    write_result_log(log_file, result_log)

    # 8. 输出统计
    total_time = time.time() - program_start_time
    print("\n" + "=" * 60)
    print("📊 处理统计报告")
    print("=" * 60)
    print(f"  总文件数: {total_files}")
    print(f"  ✅ 成功复制: {found_count} ({found_count/max(1, total_files):.1%})")
    print(f"  ❌ 未找到: {not_found_count} ({not_found_count/max(1, total_files):.1%})")
    print(f"  ⚠️ 复制错误: {copy_errors}")
    print(f"⏱️ 总耗时: {total_time:.1f}秒 | 平均速度: {total_files / max(1, total_time):.1f} 文件/秒")
    print("=" * 60)

    # 警告检查
    failure_rate = (not_found_count + copy_errors) / max(1, total_files)
    if failure_rate > 0.5:
        print(f"\n⚠️ 警告: 超过50%的文件处理失败 ({failure_rate:.1%})！")
        print("可能的原因:")
        print("  - 网络驱动器连接异常")
        print("  - 源目录路径不正确")
        print("  - 文件命名不匹配")
        print("请检查配置文件和网络连接状态")

    print(f"\n🎉 程序执行完成！")
    progress_queue.put(("complete", True))


class StdoutRedirector:
    """重定向stdout到GUI的Text组件"""
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, message):
        log_queue.put(message)

    def flush(self):
        pass  # 必须实现flush方法，否则会报错


class BatchCopyGUI(tk.Tk):
    """3D文件批量复制GUI界面"""
    def __init__(self):
        super().__init__()
        self.title(f"3D文件批量复制工具 {VERSION}")
        self.geometry("900x650")
        self.resizable(True, True)
        self.config_path = None  # 配置文件路径
        self.list_file_path = None  # 原始清单文件路径
        self.config_data = None  # 加载的配置数据
        
        # 设置样式
        self._setup_styles()
        
        # 初始化界面组件
        self._init_widgets()
        
        # 重定向stdout到日志框
        self._redirect_stdout()
        
        # 启动队列监听（处理日志和进度更新）
        self._listen_queues()
        
        # 尝试自动加载配置文件和原始清单
        self._auto_load_files()

    def _setup_styles(self):
        """设置界面样式"""
        self.style = ttk.Style()
        
        # 配置主题
        self.style.theme_use('clam')
        
        # 配置按钮样式
        self.style.configure('TButton', 
                            font=('微软雅黑', 10),
                            padding=6)
        self.style.map('TButton',
                      foreground=[('active', 'blue'), ('pressed', 'red')],
                      background=[('active', '#e0e0e0')])
        
        # 配置标签样式
        self.style.configure('TLabel', 
                            font=('微软雅黑', 10),
                            padding=4)
        
        # 配置进度条样式
        self.style.configure('TProgressbar',
                            thickness=15)
        
        # 配置标题标签样式
        self.style.configure('Header.TLabel',
                            font=('微软雅黑', 12, 'bold'),
                            foreground='#2c3e50',
                            padding=8)

    def _init_widgets(self):
        """初始化GUI组件"""
        # 创建主框架，添加内边距
        main_frame = ttk.Frame(self, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 1. 标题和版权信息
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(header_frame, text="3D文件批量复制工具", style='Header.TLabel').pack(side=tk.LEFT)
        ttk.Label(header_frame, text=f"{COPYRIGHT} | {VERSION}", font=('微软雅黑', 9)).pack(side=tk.RIGHT)
        
        # 2. 文件选择区
        file_frame = ttk.LabelFrame(main_frame, text="文件设置", padding=10)
        file_frame.pack(fill=tk.X, pady=(0, 15))
        
        # 配置文件选择
        config_frame = ttk.Frame(file_frame)
        config_frame.pack(fill=tk.X, pady=(0, 8))
        
        ttk.Label(config_frame, text="配置文件:").pack(side=tk.LEFT, padx=(0, 10))
        
        self.config_path_var = tk.StringVar()
        self.config_entry = ttk.Entry(config_frame, textvariable=self.config_path_var, state='readonly', width=60)
        self.config_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        ttk.Button(config_frame, text="浏览...", command=self._select_config).pack(side=tk.RIGHT)
        
        # 原始清单文件选择
        list_frame = ttk.Frame(file_frame)
        list_frame.pack(fill=tk.X)
        
        ttk.Label(list_frame, text="原始清单:").pack(side=tk.LEFT, padx=(0, 10))
        
        self.list_file_var = tk.StringVar()
        self.list_entry = ttk.Entry(list_frame, textvariable=self.list_file_var, state='readonly', width=60)
        self.list_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        ttk.Button(list_frame, text="浏览...", command=self._select_list_file).pack(side=tk.RIGHT)
        
        # 3. 操作按钮区
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.start_btn = ttk.Button(
            btn_frame, text="开始批量复制", command=self._start_process, state=tk.DISABLED
        )
        self.start_btn.pack(side=tk.LEFT, padx=5)
        
        self.view_log_btn = ttk.Button(
            btn_frame, text="查看日志文件", command=self._view_log, state=tk.DISABLED
        )
        self.view_log_btn.pack(side=tk.LEFT, padx=5)
        
        self.clear_log_btn = ttk.Button(
            btn_frame, text="清空日志", command=self._clear_log
        )
        self.clear_log_btn.pack(side=tk.RIGHT, padx=5)
        
        # 4. 进度条
        self.progress_frame = ttk.Frame(main_frame)
        self.progress_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.progress_label = ttk.Label(self.progress_frame, text="进度:")
        self.progress_label.pack(side=tk.LEFT)
        
        self.progress_bar = ttk.Progressbar(
            self.progress_frame, orient=tk.HORIZONTAL, length=650, mode="determinate"
        )
        self.progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        # 5. 日志显示区
        log_frame = ttk.LabelFrame(main_frame, text="处理日志", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = scrolledtext.ScrolledText(
            log_frame, wrap=tk.WORD, state=tk.DISABLED, font=("Consolas", 10)
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # 添加分隔线增强视觉效果
        ttk.Separator(main_frame, orient='horizontal').pack(fill=tk.X, pady=10)

    def _redirect_stdout(self):
        """重定向stdout到日志Text组件"""
        self.stdout_redirector = StdoutRedirector(self.log_text)
        sys.stdout = self.stdout_redirector

    def _listen_queues(self):
        """监听日志队列和进度队列，更新GUI"""
        # 处理日志队列
        while not log_queue.empty():
            message = log_queue.get_nowait()
            self.log_text.config(state=tk.NORMAL)
            self.log_text.insert(tk.END, message)
            self.log_text.see(tk.END)  # 自动滚动到末尾
            self.log_text.config(state=tk.DISABLED)

        # 处理进度队列
        while not progress_queue.empty():
            msg_type, data = progress_queue.get_nowait()
            if msg_type == "max":
                # 设置进度条最大值
                self.progress_bar["maximum"] = data
                self.progress_bar["value"] = 0
            elif msg_type == "update":
                # 更新进度值
                self.progress_bar["value"] = data
            elif msg_type == "complete":
                # 处理完成，启用按钮
                self.start_btn.config(state=tk.NORMAL)
                self.view_log_btn.config(state=tk.NORMAL)
                if data:
                    messagebox.showinfo("完成", "批量复制任务已完成！")
                else:
                    messagebox.showerror("失败", "任务执行过程中出现错误，请查看日志！")

        # 100ms后再次检查队列
        self.after(100, self._listen_queues)

    def _auto_load_files(self):
        """自动加载配置文件和原始清单"""
        root_path = get_root_path()
        
        # 尝试加载默认配置文件
        default_config_path = os.path.join(root_path, "config.ini")
        if os.path.exists(default_config_path):
            self.config_path = default_config_path
            self.config_path_var.set(os.path.basename(default_config_path))
            self.config_data = load_configuration(default_config_path)
            
            # 如果配置加载成功，尝试加载原始清单
            if self.config_data and "original_list_filename" in self.config_data:
                default_list_path = os.path.join(root_path, self.config_data["original_list_filename"])
                if os.path.exists(default_list_path):
                    self.list_file_path = default_list_path
                    self.list_file_var.set(os.path.basename(default_list_path))
                    self.start_btn.config(state=tk.NORMAL)
                else:
                    print(f"⚠️ 未找到默认原始清单文件: {default_list_path}")
        else:
            print("⚠️ 未找到默认配置文件，请手动选择")

    def _select_config(self):
        """选择配置文件"""
        default_dir = get_root_path()
        default_path = os.path.join(default_dir, "config.ini") if not self.config_path else self.config_path

        path = filedialog.askopenfilename(
            title="选择config.ini配置文件",
            initialdir=default_dir,
            initialfile=os.path.basename(default_path),
            filetypes=[("INI文件", "*.ini"), ("所有文件", "*.*")]
        )

        if path:
            self.config_path = path
            self.config_path_var.set(os.path.basename(path))
            self.config_data = load_configuration(path)
            
            # 尝试从配置中获取原始清单文件名并自动加载
            if self.config_data and "original_list_filename" in self.config_data:
                list_path = os.path.join(get_root_path(), self.config_data["original_list_filename"])
                if os.path.exists(list_path):
                    self.list_file_path = list_path
                    self.list_file_var.set(os.path.basename(list_path))
            
            # 检查是否可以启用开始按钮
            self._check_start_button_state()

    def _select_list_file(self):
        """选择原始清单文件"""
        default_dir = get_root_path()
        default_filename = self.config_data["original_list_filename"] if (self.config_data and "original_list_filename" in self.config_data) else "Original file list.csv"
        default_path = os.path.join(default_dir, default_filename)

        path = filedialog.askopenfilename(
            title="选择原始清单文件",
            initialdir=default_dir,
            initialfile=os.path.basename(default_path),
            filetypes=[("CSV文件", "*.csv"), ("所有文件", "*.*")]
        )

        if path:
            self.list_file_path = path
            self.list_file_var.set(os.path.basename(path))
            self._check_start_button_state()

    def _check_start_button_state(self):
        """检查是否可以启用开始按钮"""
        if self.config_path and self.list_file_path and self.config_data:
            # 更新配置中的清单文件路径
            self.config_data["list_file"] = self.list_file_path
            self.start_btn.config(state=tk.NORMAL)
        else:
            self.start_btn.config(state=tk.DISABLED)

    def _start_process(self):
        """启动批量复制（后台线程）"""
        if not self.config_path or not self.list_file_path or not self.config_data:
            messagebox.showwarning("警告", "请确保已选择配置文件和原始清单！")
            return

        # 禁用按钮，防止重复点击
        self.start_btn.config(state=tk.DISABLED)
        self.view_log_btn.config(state=tk.DISABLED)
        self.progress_bar["value"] = 0

        # 启动后台工作线程
        threading.Thread(
            target=worker,
            args=(self.config_data, self._update_progress),
            daemon=True
        ).start()

    def _update_progress(self, current, total):
        """更新进度（兼容旧逻辑，实际通过队列更新）"""
        pass

    def _view_log(self):
        """打开日志文件"""
        if self.config_data and "log_file" in self.config_data:
            log_path = self.config_data["log_file"]
            if os.path.exists(log_path):
                os.startfile(log_path)
            else:
                messagebox.showerror("错误", f"日志文件不存在: {log_path}")
        else:
            messagebox.showerror("错误", "未加载配置信息，无法确定日志文件位置")

    def _clear_log(self):
        """清空日志框"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)

    def destroy(self):
        """关闭窗口时恢复stdout"""
        sys.stdout = sys.__stdout__
        super().destroy()


if __name__ == "__main__":
    app = BatchCopyGUI()
    app.mainloop()