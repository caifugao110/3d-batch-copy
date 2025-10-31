import os
import sys
import shutil
import csv
import time
import configparser
import queue
import threading
import json
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
import customtkinter as ctk
from tkinter import filedialog, messagebox
import webbrowser
import subprocess
import platform

# 版本和版权信息
VERSION = "V1.1.6"
COPYRIGHT = "Tobin © 2025"
PROJECT_URL = "https://github.com/caifugao110/3d-batch-copy"

# 全局队列：用于子线程与GUI主线程通信
log_queue = queue.Queue()
progress_queue = queue.Queue()

# 默认主题设置 - 修复主题默认值为customtkinter支持的"blue"
DEFAULT_APPEARANCE_MODE = "light"  # "dark", "light", "system"
DEFAULT_COLOR_THEME = "blue"     # "blue", "green", "dark-blue" (customtkinter支持的主题)

# 初始化主题
ctk.set_appearance_mode(DEFAULT_APPEARANCE_MODE)
ctk.set_default_color_theme(DEFAULT_COLOR_THEME)


def get_root_path():
    """获取程序根目录（exe所在目录，支持PyInstaller打包后路径）"""
    if hasattr(sys, '_MEIPASS'):
        return os.path.dirname(os.path.abspath(sys.executable))
    else:
        return os.path.dirname(os.path.abspath(__file__))


def clean_filename(name):
    """清理文件名：去除特定后缀和标识符，统一转为小写"""
    # 新增处理：包含 "-L(" 则分割取前面部分
    if "-L(" in name:
        parts = name.split("-L(")
        name = parts[0]
    # 新增处理：以 "-L" 结尾则分割取前面部分
    if name.endswith("-L"):
        parts = name.split("-L")
        name = parts[0]
    # 原有处理：以 "L" 结尾
    if name.endswith("L"):
        name = name[:-1]
    # 原有处理：包含 "L("
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
        
        # 读取重命名选项配置
        rename_option = config.getboolean("Settings", "rename_files", fallback=False)

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
        print(f"   映射的驱动器盘符: {drive_letter}")
        print(f"   源目录数量: {len(source_dirs)}")
        print(f"   目标目录: {target_dir}")
        print(f"   待处理列表: {list_file}")
        print(f"   复制日志文件: {log_file}")
        print(f"   最大线程数: {max_workers}")
        print(f"   重试次数: {retry_attempts}")
        print(f"   按清单重命名: {'是' if rename_option else '否'}")

        return {
            "source_dirs": source_dirs,
            "target_dir": target_dir,
            "list_file": list_file,
            "log_file": log_file,
            "max_workers": max_workers,
            "retry_attempts": retry_attempts,
            "original_list_filename": original_list_filename,
            "drive_letter": drive_letter,
            "target_dir_name": target_dir_name,
            "log_filename": log_filename,
            "config_path": config_path,
            "rename_files": rename_option  # 添加重命名选项
        }
    except Exception as e:
        print(f"🔥 配置文件解析失败: {str(e)}")
        print("请检查 config.ini 文件格式是否正确")
        return None


def save_configuration(config_path, config_data):
    """保存配置文件"""
    try:
        config = configparser.ConfigParser()
        config.optionxform = lambda option: option
        
        # 保存基本配置
        config["Paths"] = {
            "drive_letter": config_data.get("drive_letter", "D"),
            "target_dir_name": config_data.get("target_dir_name", "Target"),
            "original_list_file": config_data.get("original_list_filename", "Original file list.csv"),
            "log_file": config_data.get("log_filename", "log.csv")
        }
        
        config["Settings"] = {
            "max_workers": str(config_data.get("max_workers", 12)),
            "retry_attempts": str(config_data.get("retry_attempts", 3)),
            "rename_files": str(config_data.get("rename_files", False)).lower()  # 保存重命名选项
        }
        
        # 保存源目录
        config["SourceDirectories"] = {}
        for idx, src_dir in enumerate(config_data.get("source_dirs", []), 1):
            # 提取相对路径（假设格式为 D:\path\to\dir）
            if ":\\" in src_dir:
                rel_path = src_dir.split(":\\", 1)[1]
            else:
                rel_path = src_dir
            config["SourceDirectories"][f"source_{idx}"] = rel_path
        
        with open(config_path, "w", encoding="utf-8") as f:
            config.write(f)
        
        print(f"✅ 配置已保存至: {config_path}")
        return True
    except Exception as e:
        print(f"🔥 配置保存失败: {str(e)}")
        return False


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
    """构建文件索引（支持递归）"""
    print("⏳ 正在构建全局文件索引（包含子目录）...")
    index = defaultdict(list)
    start_time = time.time()
    total_files = 0

    for src_dir in source_dirs:
        try:
            if not os.path.exists(src_dir):
                print(f"⚠️ 路径不存在或无权限访问: {src_dir}")
                continue
            # 使用 os.walk 递归遍历
            for root, dirs, files in os.walk(src_dir):
                for file in files:
                    if file.lower().endswith(".step"):
                        full_path = os.path.join(root, file)
                        base_name = os.path.splitext(file)[0]
                        clean_base = clean_filename(base_name)
                        prefix_key = clean_base[:4] if len(clean_base) >= 4 else clean_base
                        # 注意：这里存储的 src_dir 是实际的 root，即文件所在的具体目录
                        index[prefix_key].append((clean_base, file, root))
                        total_files += 1
        except Exception as e:
            print(f"⚠️ 目录扫描失败: {src_dir} - {str(e)}")

    index_time = time.time() - start_time
    print(f"✅ 索引构建完成: {len(index)} 个前缀组, {total_files} 个文件, 耗时 {index_time:.2f}秒")
    return index

def read_original_file_list(list_file):
    """读取待处理文件列表（仅支持CSV和TXT格式）"""
    try:
        _, ext = os.path.splitext(list_file)
        ext = ext.lower()
        
        if ext == '.csv':
            with open(list_file, "r", encoding="utf-8-sig") as f:
                reader = csv.reader(f)
                all_lines = [row[0].strip() for row in reader if row and row[0].strip()]
        elif ext == '.txt':
            with open(list_file, "r", encoding="utf-8-sig") as f:
                all_lines = [line.strip() for line in f if line.strip()]
        else:
            print(f"⚠️ 不支持的文件格式: {ext}，仅支持CSV和TXT文件")
            return None
        
        print(f"📋 待处理文件数: {len(all_lines)}")
        return all_lines
    except Exception as e:
        print(f"🔥 文件读取失败: {str(e)}")
        print(f"请检查文件是否存在且格式正确: {list_file}")
        return None


def process_item(item, target_dir, index, retry_attempts, stop_event, rename_files):
    """处理单个文件复制，添加stop_event参数用于终止，添加rename_files参数控制是否重命名"""
    original_name, search_name = item
    # 根据是否需要重命名决定目标文件名
    if rename_files:
        dst_file = os.path.join(target_dir, f"{original_name}.STEP")
    else:
        # 先不确定文件名，找到源文件后再确定
        dst_file = None
    prefix_key = search_name[:4] if len(search_name) >= 4 else search_name

    # 检查是否需要停止
    if stop_event.is_set():
        return {
            "status": "cancelled",
            "original": original_name,
            "copied": "操作已取消",
            "source": ""
        }

    if prefix_key in index:
        for clean_base, src_filename, src_dir in index[prefix_key]:
            if clean_base == search_name:  # 完全匹配
                for attempt in range(retry_attempts):
                    # 检查是否需要停止
                    if stop_event.is_set():
                        return {
                            "status": "cancelled",
                            "original": original_name,
                            "copied": "操作已取消",
                            "source": ""
                        }
                        
                    try:
                        src_path = os.path.join(src_dir, src_filename)
                        # 如果不需要重命名，使用源文件的文件名
                        if not rename_files and dst_file is None:
                            dst_file = os.path.join(target_dir, src_filename)
                        shutil.copy2(src_path, dst_file)
                        return {
                            "status": "success",
                            "original": original_name,
                            "copied": src_filename,
                            "source": src_dir,
                            "renamed_to": original_name if rename_files else None
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


def worker(config, progress_callback, stop_event):
    """后台工作线程：执行完整的复制流程，添加stop_event参数"""
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
    rename_files = config.get("rename_files", False)  # 获取重命名选项

    # 清理目标目录
    cleanup_target_directory(target_dir)

    # 构建文件索引
    index = build_file_index(source_dirs)

    # 读取待处理列表
    original_files = read_original_file_list(list_file)
    if not original_files or len(original_files) == 0:
        print("🔥 无待处理文件，程序退出")
        progress_queue.put(("complete", False))
        return
    total_files = len(original_files)
    progress_queue.put(("max", total_files))

    # 预处理搜索名
    search_items = [(orig, clean_filename(orig)) for orig in original_files]

    # 多线程复制
    print(f"📦 开始并行复制文件... {'(将按清单重命名)' if rename_files else ''}")
    executor = ThreadPoolExecutor(max_workers=max_workers)
    # 传递rename_files参数
    futures = [executor.submit(process_item, item, target_dir, index, retry_attempts, stop_event, rename_files) 
              for item in search_items]
    
    try:
        for idx, future in enumerate(as_completed(futures)):
            # 检查是否需要停止
            if stop_event.is_set():
                executor.shutdown(wait=False)
                print("⏹️ 正在终止所有复制任务...")
                break
                
            result = future.result()
            result_log.append(result)

            # 更新统计
            if result["status"] == "success":
                found_count += 1
            elif result["status"] == "not_found":
                not_found_count += 1
            elif result["status"] == "error":
                copy_errors += 1

            # 更新进度，包括统计信息
            progress_queue.put((
                "update", 
                idx + 1, 
                found_count, 
                not_found_count + copy_errors,
                (idx + 1) / max(1, time.time() - program_start_time)
            ))
    finally:
        # 确保执行器被关闭
        if not executor._shutdown:
            executor.shutdown(wait=False)

    # 如果是被取消的，不写入完整日志
    if not stop_event.is_set():
        # 写入日志
        write_result_log(log_file, result_log)

        # 输出统计
        total_time = time.time() - program_start_time
        print("\n" + "=" * 60)
        print("📊 处理统计报告")
        print("=" * 60)
        print(f"  总文件数: {total_files}")
        print(f"  ✅ 成功复制: {found_count} ({found_count/max(1, total_files):.1%})")
        print(f"  ❌ 未找到: {not_found_count} ({not_found_count/max(1, total_files):.1%})")
        print(f"  ⚠️ 复制错误: {copy_errors}")
        print(f"⏱️ 总耗时: {total_time:.1f}秒 | 平均速度: {total_files / max(1, total_time):.1f} 文件/秒")
        print(f"  重命名模式: {'启用' if rename_files else '禁用'}")
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
    else:
        print("\n⏹️ 任务已被用户终止")

    progress_queue.put(("complete", not stop_event.is_set()))


def write_result_log(log_file, result_log):
    """写入复制日志文件，使用utf-8-sig编码解决Office乱码问题"""
    print("📝 正在写入复制日志文件...")
    try:
        # 使用gbk编码，解决Office打开乱码问题
        with open(log_file, "w", encoding="gbk", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["原始文件名", "实际复制文件名", "来源路径", "重命名状态"])
            for res in result_log:
                renamed_status = "是" if res.get("renamed_to") else "否"
                writer.writerow([res["original"], res["copied"], res["source"], renamed_status])
        print(f"✅ 复制日志已保存至: {log_file}")
        return True
    except Exception as e:
        print(f"⚠️ 复制日志文件写入失败: {str(e)}")
        return False


class StdoutRedirector:
    """重定向stdout到GUI的Text组件"""
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, message):
        log_queue.put(message)

    def flush(self):
        pass


class SettingsWindow(ctk.CTkToplevel):
    """配置管理窗口"""
    def __init__(self, parent, config_data, on_save_callback):
        super().__init__(parent)
        self.title("配置管理")
        self.geometry("700x800")
        self.config_data = config_data.copy() if config_data else {}
        self.on_save_callback = on_save_callback
        self.original_drive_letter = self.config_data.get("drive_letter", "D")
        
        # 使窗口模态化
        self.transient(parent)
        self.grab_set()
        
        self._init_widgets()
        
        # 绑定驱动器盘符变化事件
        self.drive_entry.bind("<FocusOut>", self._on_drive_letter_change)
        self.drive_entry.bind("<Return>", self._on_drive_letter_change)
        
    def _init_widgets(self):
        """初始化配置窗口组件"""
        # 主框架
        main_frame = ctk.CTkScrollableFrame(self, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=15, pady=15)
        
        # 基本设置区
        basic_frame = ctk.CTkFrame(main_frame, fg_color=("gray86", "gray17"))
        basic_frame.pack(fill="x", pady=(0, 15))
        
        ctk.CTkLabel(basic_frame, text="基本设置", font=("微软雅黑", 16, "bold")).pack(anchor="w", padx=15, pady=(10, 5))
        
        # 驱动器盘符
        drive_frame = ctk.CTkFrame(basic_frame, fg_color="transparent")
        drive_frame.pack(fill="x", padx=15, pady=5)
        ctk.CTkLabel(drive_frame, text="\\\\192.168.160.2\\生产管理部3d\\3D 资料 映射的盘符:", width=100, font=("微软雅黑", 12)).pack(side="left")
        self.drive_entry = ctk.CTkEntry(drive_frame, width=200, font=("微软雅黑", 12))
        self.drive_entry.pack(side="left", padx=(10, 0))
        self.drive_entry.insert(0, self.config_data.get("drive_letter", "D"))
        
        # 目标目录名
        target_frame = ctk.CTkFrame(basic_frame, fg_color="transparent")
        target_frame.pack(fill="x", padx=15, pady=5)
        ctk.CTkLabel(target_frame, text="目标目录名:", width=100, font=("微软雅黑", 12)).pack(side="left")
        self.target_entry = ctk.CTkEntry(target_frame, width=200, font=("微软雅黑", 12))
        self.target_entry.pack(side="left", padx=(10, 0))
        self.target_entry.insert(0, self.config_data.get("target_dir_name", "Target"))
        
        # 原始清单文件名
        list_frame = ctk.CTkFrame(basic_frame, fg_color="transparent")
        list_frame.pack(fill="x", padx=15, pady=5)
        ctk.CTkLabel(list_frame, text="原始清单文件:", width=100, font=("微软雅黑", 12)).pack(side="left")
        self.list_entry = ctk.CTkEntry(list_frame, width=200, font=("微软雅黑", 12))
        self.list_entry.pack(side="left", padx=(10, 0))
        self.list_entry.insert(0, self.config_data.get("original_list_filename", "Original file list.csv"))
        
        # 复制日志文件名
        log_frame = ctk.CTkFrame(basic_frame, fg_color="transparent")
        log_frame.pack(fill="x", padx=15, pady=(5, 10))
        ctk.CTkLabel(log_frame, text="复制日志文件名:", width=100, font=("微软雅黑", 12)).pack(side="left")
        self.log_entry = ctk.CTkEntry(log_frame, width=200, font=("微软雅黑", 12))
        self.log_entry.pack(side="left", padx=(10, 0))
        self.log_entry.insert(0, self.config_data.get("log_filename", "log.csv"))
        
        # 性能设置区
        perf_frame = ctk.CTkFrame(main_frame, fg_color=("gray86", "gray17"))
        perf_frame.pack(fill="x", pady=(0, 15))
        
        ctk.CTkLabel(perf_frame, text="性能设置", font=("微软雅黑", 16, "bold")).pack(anchor="w", padx=15, pady=(10, 5))
        
        # 最大线程数
        workers_frame = ctk.CTkFrame(perf_frame, fg_color="transparent")
        workers_frame.pack(fill="x", padx=15, pady=5)
        ctk.CTkLabel(workers_frame, text="最大线程数:", width=100, font=("微软雅黑", 12)).pack(side="left")
        self.workers_entry = ctk.CTkEntry(workers_frame, width=200, font=("微软雅黑", 12))
        self.workers_entry.pack(side="left", padx=(10, 0))
        self.workers_entry.insert(0, str(self.config_data.get("max_workers", 12)))
        
        # 重试次数
        retry_frame = ctk.CTkFrame(perf_frame, fg_color="transparent")
        retry_frame.pack(fill="x", padx=15, pady=5)
        ctk.CTkLabel(retry_frame, text="重试次数:", width=100, font=("微软雅黑", 12)).pack(side="left")
        self.retry_entry = ctk.CTkEntry(retry_frame, width=200, font=("微软雅黑", 12))
        self.retry_entry.pack(side="left", padx=(10, 0))
        self.retry_entry.insert(0, str(self.config_data.get("retry_attempts", 3)))
        
        # 添加重命名选项
        rename_frame = ctk.CTkFrame(perf_frame, fg_color="transparent")
        rename_frame.pack(fill="x", padx=15, pady=(5, 10))
        self.rename_var = ctk.BooleanVar(value=self.config_data.get("rename_files", False))
        ctk.CTkCheckBox(
            rename_frame, 
            text="按照清单重命名3D文件", 
            variable=self.rename_var,
            font=("微软雅黑", 12)
        ).pack(anchor="w")
        
        # 源目录管理区
        source_frame = ctk.CTkFrame(main_frame, fg_color=("gray86", "gray17"))
        source_frame.pack(fill="both", expand=True, pady=(0, 15))
        
        ctk.CTkLabel(source_frame, text="源目录管理", font=("微软雅黑", 16, "bold")).pack(anchor="w", padx=15, pady=(10, 5))
        
        # 源目录列表框
        list_frame_inner = ctk.CTkFrame(source_frame, fg_color="transparent")
        list_frame_inner.pack(fill="both", expand=True, padx=15, pady=5)
        
        self.source_listbox = ctk.CTkTextbox(list_frame_inner, height=150, width=400, font=("微软雅黑", 12))
        self.source_listbox.pack(fill="both", expand=True)
        
        # 初始化源目录列表
        source_dirs = self.config_data.get("source_dirs", [])
        for src_dir in source_dirs:
            self.source_listbox.insert("end", src_dir + "\n")
        
        # 源目录操作按钮
        btn_frame = ctk.CTkFrame(source_frame, fg_color="transparent")
        btn_frame.pack(fill="x", padx=15, pady=(5, 10))
        
        ctk.CTkButton(btn_frame, text="添加目录", width=100, font=("微软雅黑", 12), command=self._add_source_dir).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="删除选中", width=100, font=("微软雅黑", 12), command=self._remove_source_dir).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="清空全部", width=100, font=("微软雅黑", 12), command=self._clear_source_dirs).pack(side="left", padx=5)
        
        # 底部按钮
        footer_frame = ctk.CTkFrame(self, fg_color="transparent")
        footer_frame.pack(fill="x", padx=15, pady=15)
        
        ctk.CTkButton(footer_frame, text="保存配置", width=150, font=("微软雅黑", 12), command=self._save_config).pack(side="right", padx=5)
        ctk.CTkButton(footer_frame, text="取消", width=150, font=("微软雅黑", 12), command=self.destroy).pack(side="right", padx=5)
    
    def _on_drive_letter_change(self, event=None):
        """当驱动器盘符改变时，更新源目录中的盘符"""
        new_drive = self.drive_entry.get().strip().upper()
        if not new_drive or new_drive == self.original_drive_letter:
            return
            
        # 获取当前所有源目录
        content = self.source_listbox.get("1.0", "end").strip()
        if not content:
            self.original_drive_letter = new_drive
            return
            
        # 替换每个目录的盘符
        updated_dirs = []
        for line in content.split("\n"):
            line = line.strip()
            if line and ":\\" in line:
                # 分割盘符和路径
                _, path = line.split(":\\", 1)
                updated_dir = f"{new_drive}:\\{path}"
                updated_dirs.append(updated_dir)
            else:
                updated_dirs.append(line)
        
        # 更新列表框内容
        self.source_listbox.delete("1.0", "end")
        for dir_path in updated_dirs:
            self.source_listbox.insert("end", dir_path + "\n")
            
        self.original_drive_letter = new_drive
    
    def _add_source_dir(self):
        """添加源目录"""
        dir_path = filedialog.askdirectory(title="选择源目录")
        if dir_path:
            self.source_listbox.insert("end", dir_path + "\n")
    
    def _remove_source_dir(self):
        """删除源目录"""
        try:
            # 获取当前光标位置所在的行
            current_line = self.source_listbox.index("insert").split(".")[0]
            self.source_listbox.delete(f"{current_line}.0", f"{current_line}.end+1c")
        except:
            messagebox.showwarning("警告", "请先选择要删除的目录")
    
    def _clear_source_dirs(self):
        """清空所有源目录"""
        if messagebox.askyesno("确认", "确定要清空所有源目录吗？"):
            self.source_listbox.delete("1.0", "end")
    
    def _save_config(self):
        """保存配置"""
        try:
            # 验证输入
            max_workers = int(self.workers_entry.get())
            retry_attempts = int(self.retry_entry.get())
            
            if max_workers < 1 or retry_attempts < 1:
                messagebox.showerror("错误", "线程数和重试次数必须大于0")
                return
            
            # 收集源目录
            source_dirs = []
            content = self.source_listbox.get("1.0", "end").strip()
            if content:
                source_dirs = [line.strip() for line in content.split("\n") if line.strip()]
            
            # 更新配置数据
            self.config_data["drive_letter"] = self.drive_entry.get().strip()
            self.config_data["target_dir_name"] = self.target_entry.get().strip()
            self.config_data["original_list_filename"] = self.list_entry.get().strip()
            self.config_data["log_filename"] = self.log_entry.get().strip()
            self.config_data["max_workers"] = max_workers
            self.config_data["retry_attempts"] = retry_attempts
            self.config_data["source_dirs"] = source_dirs
            self.config_data["rename_files"] = self.rename_var.get()  # 保存重命名选项
            
            # 调用保存回调
            if self.on_save_callback:
                self.on_save_callback(self.config_data)
            
            messagebox.showinfo("成功", "配置已保存并重新加载")
            self.destroy()
        except ValueError:
            messagebox.showerror("错误", "线程数和重试次数必须是整数")


class BatchCopyGUI(ctk.CTk):
    """3D文件批量复制GUI界面 - CustomTkinter 版本"""
    def __init__(self):
        super().__init__()
        self.title(f"3D文件批量复制工具 {VERSION}")
        self.geometry("1200x800")
        self.minsize(1000, 700)
        
        # 配置变量
        self.config_path = None
        self.list_file_path = None
        self.config_data = None
        self.running = False
        self.task_history = []  # 任务历史记录
        self.stop_event = threading.Event()  # 添加停止事件
        self.worker_thread = None  # 保存工作线程引用
        self.total_files = 0
        self.success_count = 0
        self.failure_count = 0
        self.start_time = 0
        
        # 初始化界面
        self._init_widgets()
        
        # 重定向stdout
        self._redirect_stdout()
        
        # 启动队列监听
        self._listen_queues()
        
        # 自动加载配置文件
        self.after(200, self._auto_load_files)
    
    def _init_widgets(self):
        """初始化GUI组件"""
        # 创建主容器
        main_container = ctk.CTkFrame(self)
        main_container.pack(fill="both", expand=True, padx=0, pady=0)
        
        # 顶部标题栏
        header_frame = ctk.CTkFrame(main_container, fg_color=("gray90", "gray20"), height=100)
        header_frame.pack(fill="x", padx=0, pady=0)
        header_frame.pack_propagate(False)
        
        # 标题和主题选择
        title_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        title_frame.pack(fill="x", padx=20, pady=10)
        
        # 标题
        title_label = ctk.CTkLabel(
            title_frame, 
            text="3D文件批量复制工具",
            font=("微软雅黑", 26, "bold"),
            text_color=("#1f77b4", "#64b5f6")
        )
        title_label.pack(anchor="w", side="left")
        
        # 主题选择
        theme_frame = ctk.CTkFrame(title_frame, fg_color="transparent")
        theme_frame.pack(anchor="e", side="right")
        
        ctk.CTkLabel(
            theme_frame,
            text="主题:",
            font=("微软雅黑", 12),
            text_color=("gray50", "gray70")
        ).pack(side="left", padx=(0, 10))
        
        self.appearance_mode_optionemenu = ctk.CTkOptionMenu(
            theme_frame,
            values=["light", "dark", "system"],
            command=self._change_appearance_mode_event,
            font=("微软雅黑", 12),
            width=120
        )
        self.appearance_mode_optionemenu.set(DEFAULT_APPEARANCE_MODE)
        self.appearance_mode_optionemenu.pack(side="left", padx=(0, 10))
        
        self.color_theme_optionemenu = ctk.CTkOptionMenu(
            theme_frame,
            values=["blue", "green", "dark-blue"],  # 仅保留customtkinter支持的主题
            command=self._change_color_theme_event,
            font=("微软雅黑", 12),
            width=120
        )
        self.color_theme_optionemenu.set(DEFAULT_COLOR_THEME)
        self.color_theme_optionemenu.pack(side="left")
        
        # 版本和链接
        info_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        info_frame.pack(anchor="w", padx=20, pady=(5, 10))
        
        ctk.CTkLabel(
            info_frame,
            text=f"{COPYRIGHT} | {VERSION}",
            font=("微软雅黑", 12),
            text_color=("gray50", "gray70")
        ).pack(side="left", padx=(0, 20))
        
        # GitHub链接
        github_btn = ctk.CTkButton(
            info_frame,
            text="📌 GitHub地址",
            width=120,
            height=30,
            font=("微软雅黑", 12),
            command=lambda: webbrowser.open(PROJECT_URL)
        )
        github_btn.pack(side="left", padx=5)
        
        # 使用说明链接
        help_btn = ctk.CTkButton(
            info_frame,
            text="❓ 使用说明",
            width=120,
            height=30,
            font=("微软雅黑", 12),
            command=lambda: webbrowser.open("https://caifugao110.github.io/3d-batch-copy/")
        )
        help_btn.pack(side="left", padx=5)
        
        # 主内容区
        content_frame = ctk.CTkFrame(main_container, fg_color="transparent")
        content_frame.pack(fill="both", expand=True, padx=15, pady=15)
        
        # 左侧面板（文件选择和操作）
        left_panel = ctk.CTkFrame(content_frame, fg_color=("gray86", "gray17"))
        left_panel.pack(side="left", fill="both", expand=False, padx=(0, 10))
        
        # 文件选择区
        file_section = ctk.CTkFrame(left_panel, fg_color="transparent")
        file_section.pack(fill="x", padx=15, pady=15)
        
        ctk.CTkLabel(file_section, text="文件设置", font=("微软雅黑", 14, "bold")).pack(anchor="w", pady=(0, 10))
        
        # 配置文件
        config_label_frame = ctk.CTkFrame(file_section, fg_color="transparent")
        config_label_frame.pack(fill="x", pady=(0, 5))
        ctk.CTkLabel(config_label_frame, text="📋 配置文件:", font=("微软雅黑", 12)).pack(anchor="w")
        
        config_btn_frame = ctk.CTkFrame(file_section, fg_color="transparent")
        config_btn_frame.pack(fill="x", pady=(0, 10))
        
        self.config_label = ctk.CTkLabel(
            config_btn_frame,
            text="未选择",
            font=("微软雅黑", 11),
            text_color=("gray50", "gray70"),
            fg_color=("gray95", "gray25"),
            corner_radius=5,
            padx=10,
            pady=8
        )
        self.config_label.pack(fill="x", side="left", expand=True, padx=(0, 5))
        
        ctk.CTkButton(
            config_btn_frame,
            text="浏览",
            width=80,
            font=("微软雅黑", 12),
            command=self._select_config
        ).pack(side="left")
        
        # 原始清单
        list_label_frame = ctk.CTkFrame(file_section, fg_color="transparent")
        list_label_frame.pack(fill="x", pady=(0, 5))
        ctk.CTkLabel(list_label_frame, text="📄 原始清单:", font=("微软雅黑", 12)).pack(anchor="w")
        
        list_btn_frame = ctk.CTkFrame(file_section, fg_color="transparent")
        list_btn_frame.pack(fill="x", pady=(0, 10))
        
        self.list_label = ctk.CTkLabel(
            list_btn_frame,
            text="未选择",
            font=("微软雅黑", 11),
            text_color=("gray50", "gray70"),
            fg_color=("gray95", "gray25"),
            corner_radius=5,
            padx=10,
            pady=8
        )
        self.list_label.pack(fill="x", side="left", expand=True, padx=(0, 5))
        
        ctk.CTkButton(
            list_btn_frame,
            text="浏览",
            width=80,
            font=("微软雅黑", 12),
            command=self._select_list_file
        ).pack(side="left")
        
        # 添加重命名选项复选框
        self.rename_checkbox_var = ctk.BooleanVar(value=False)
        self.rename_checkbox = ctk.CTkCheckBox(
            file_section,
            text="按照清单重命名3D文件",
            variable=self.rename_checkbox_var,
            font=("微软雅黑", 12),
            command=self._on_rename_checkbox_change
        )
        self.rename_checkbox.pack(anchor="w", pady=(10, 0))
        
        # 操作按钮区
        btn_section = ctk.CTkFrame(left_panel, fg_color="transparent")
        btn_section.pack(fill="x", padx=15, pady=(0, 15))
        
        ctk.CTkLabel(btn_section, text="操作", font=("微软雅黑", 14, "bold")).pack(anchor="w", pady=(0, 10))
        
        self.start_btn = ctk.CTkButton(
            btn_section,
            text="🚀 开始批量复制",
            font=("微软雅黑", 13, "bold"),
            height=45,
            state="disabled",
            command=self._start_process
        )
        self.start_btn.pack(fill="x", pady=(0, 8))
        
        # 添加停止处理按钮
        self.stop_btn = ctk.CTkButton(
            btn_section,
            text="⏹️ 停止处理",
            font=("微软雅黑", 13, "bold"),
            height=45,
            state="disabled",
            fg_color="#e74c3c",  # 红色
            hover_color="#c0392b",
            command=self._stop_process
        )
        self.stop_btn.pack(fill="x", pady=(0, 8))
        
        ctk.CTkButton(
            btn_section,
            text="⚙️ 配置管理",
            font=("微软雅黑", 13),
            height=40,
            command=self._open_settings
        ).pack(fill="x", pady=(0, 8))
        
        ctk.CTkButton(
            btn_section,
            text="📂 打开目标目录",
            font=("微软雅黑", 13),
            height=40,
            state="disabled",
            command=self._open_target_dir
        ).pack(fill="x", pady=(0, 8))
        self.open_target_btn = btn_section.winfo_children()[-1]
        
        ctk.CTkButton(
            btn_section,
            text="📊 查看复制日志",
            font=("微软雅黑", 13),
            height=40,
            state="disabled",
            command=self._view_log
        ).pack(fill="x", pady=(0, 8))
        self.view_log_btn = btn_section.winfo_children()[-1]
        
        ctk.CTkButton(
            btn_section,
            text="🗑️ 清空日志框",
            font=("微软雅黑", 13),
            height=40,
            command=self._clear_log
        ).pack(fill="x")
        
        # 右侧面板（进度和日志）
        right_panel = ctk.CTkFrame(content_frame, fg_color="transparent")
        right_panel.pack(side="left", fill="both", expand=True)
        
        # 进度区
        progress_section = ctk.CTkFrame(right_panel, fg_color=("gray86", "gray17"))
        progress_section.pack(fill="x", padx=0, pady=(0, 10))
        
        ctk.CTkLabel(progress_section, text="📈 处理进度", font=("微软雅黑", 14, "bold")).pack(anchor="w", padx=15, pady=(10, 5))
        
        # 进度条
        progress_bar_frame = ctk.CTkFrame(progress_section, fg_color="transparent")
        progress_bar_frame.pack(fill="x", padx=15, pady=(0, 5))
        
        self.progress_bar = ctk.CTkProgressBar(progress_bar_frame, height=30)
        self.progress_bar.pack(fill="x", side="left", expand=True, padx=(0, 10))
        self.progress_bar.set(0)
        
        self.progress_percent = ctk.CTkLabel(
            progress_bar_frame,
            text="0%",
            font=("微软雅黑", 13, "bold"),
            width=60
        )
        self.progress_percent.pack(side="left")
        
        # 统计信息
        stats_frame = ctk.CTkFrame(progress_section, fg_color="transparent")
        stats_frame.pack(fill="x", padx=15, pady=(0, 10))
        
        self.stats_label = ctk.CTkLabel(
            stats_frame,
            text="已处理: 0 | 成功: 0 | 失败: 0 | 速度: 0 文件/秒",
            font=("微软雅黑", 12),
            text_color=("gray50", "gray70")
        )
        self.stats_label.pack(anchor="w")
        
        # 日志区
        log_section = ctk.CTkFrame(right_panel, fg_color=("gray86", "gray17"))
        log_section.pack(fill="both", expand=True)
        
        ctk.CTkLabel(log_section, text="📝 处理日志", font=("微软雅黑", 14, "bold")).pack(anchor="w", padx=15, pady=(10, 5))
        
        # 日志文本框
        log_text_frame = ctk.CTkFrame(log_section, fg_color="transparent")
        log_text_frame.pack(fill="both", expand=True, padx=15, pady=(0, 10))
        
        self.log_textbox = ctk.CTkTextbox(
            log_text_frame,
            wrap="word",
            font=("微软雅黑", 14),
            state="disabled"
        )
        self.log_textbox.pack(fill="both", expand=True, side="left")
        
        # 滚动条
        scrollbar = ctk.CTkScrollbar(
            log_text_frame,
            command=self.log_textbox.yview
        )
        scrollbar.pack(side="right", fill="y")
        self.log_textbox.configure(yscrollcommand=scrollbar.set)
    
    def _on_rename_checkbox_change(self):
        """重命名选项变更时立即保存配置"""
        if self.config_data and self.config_path:
            self.config_data["rename_files"] = self.rename_checkbox_var.get()
            save_configuration(self.config_path, self.config_data)
            # 重新加载配置以确保一致性
            self.config_data = load_configuration(self.config_path)
    
    def _change_appearance_mode_event(self, new_appearance_mode: str):
        ctk.set_appearance_mode(new_appearance_mode)
    
    def _change_color_theme_event(self, new_color_theme: str):
        ctk.set_default_color_theme(new_color_theme)
    
    def _redirect_stdout(self):
        """重定向标准输出到日志文本框"""
        self.original_stdout = sys.stdout
        sys.stdout = StdoutRedirector(self.log_textbox)
    
    def _listen_queues(self):
        """监听日志和进度队列，更新GUI"""
        # 处理日志队列
        while not log_queue.empty():
            message = log_queue.get()
            self.log_textbox.configure(state="normal")
            self.log_textbox.insert("end", message)
            self.log_textbox.see("end")
            self.log_textbox.configure(state="disabled")
        
        # 处理进度队列
        while not progress_queue.empty():
            item = progress_queue.get()
            if item[0] == "max":
                self.total_files = item[1]
                self.success_count = 0
                self.failure_count = 0
                self.start_time = time.time()
                self.progress_bar.set(0)
                self.progress_percent.configure(text="0%")
                self.stats_label.configure(text="已处理: 0 | 成功: 0 | 失败: 0 | 速度: 0 文件/秒")
            elif item[0] == "update":
                current = item[1]
                self.success_count = item[2]
                self.failure_count = item[3]
                speed = item[4]
                
                if self.total_files > 0:
                    percentage = (current / self.total_files) * 100
                    self.progress_bar.set(percentage / 100)
                    self.progress_percent.configure(text=f"{int(percentage)}%")
                
                self.stats_label.configure(
                    text=f"已处理: {current} | 成功: {self.success_count} | 失败: {self.failure_count} | 速度: {speed:.1f} 文件/秒"
                )
            elif item[0] == "complete":
                self.running = False
                self.start_btn.configure(state="normal")
                self.stop_btn.configure(state="disabled")
                
                # 如果任务完成且成功，更新按钮状态
                if item[1]:
                    self.open_target_btn.configure(state="normal")
                    self.view_log_btn.configure(state="normal")
        
        # 继续监听
        self.after(100, self._listen_queues)
    
    def _auto_load_files(self):
        """自动加载默认配置文件"""
        root_path = get_root_path()
        default_config = os.path.join(root_path, "config.ini")
        
        if os.path.exists(default_config):
            self.config_path = default_config
            self.config_label.configure(text=os.path.basename(default_config))
            self.config_data = load_configuration(default_config)
            
            if self.config_data:
                # 更新重命名选项
                self.rename_checkbox_var.set(self.config_data.get("rename_files", False))
                
                # 检查默认的清单文件是否存在
                default_list = self.config_data.get("list_file")
                if os.path.exists(default_list):
                    self.list_file_path = default_list
                    self.list_label.configure(text=os.path.basename(default_list))
                
                # 启用开始按钮
                self.start_btn.configure(state="normal")
    
    def _select_config(self):
        """选择配置文件"""
        file_path = filedialog.askopenfilename(
            title="选择配置文件",
            filetypes=[("INI文件", "*.ini"), ("所有文件", "*.*")]
        )
        
        if file_path:
            self.config_path = file_path
            self.config_label.configure(text=os.path.basename(file_path))
            self.config_data = load_configuration(file_path)
            
            if self.config_data:
                # 更新重命名选项
                self.rename_checkbox_var.set(self.config_data.get("rename_files", False))
                
                # 检查是否有对应的清单文件
                list_file = self.config_data.get("list_file")
                if os.path.exists(list_file):
                    self.list_file_path = list_file
                    self.list_label.configure(text=os.path.basename(list_file))
                
                # 启用开始按钮
                self.start_btn.configure(state="normal")
            else:
                self.start_btn.configure(state="disabled")
    
    def _select_list_file(self):
        """选择原始清单文件"""
        file_path = filedialog.askopenfilename(
            title="选择原始清单文件",
            filetypes=[("CSV文件", "*.csv"), ("TXT文件", "*.txt"), ("所有文件", "*.*")]
        )
        
        if file_path:
            self.list_file_path = file_path
            self.list_label.configure(text=os.path.basename(file_path))
            
            # 如果有配置数据，更新配置中的清单文件名
            if self.config_data and self.config_path:
                self.config_data["original_list_filename"] = os.path.basename(file_path)
                save_configuration(self.config_path, self.config_data)
            
            # 如果配置已加载，启用开始按钮
            if self.config_data:
                self.start_btn.configure(state="normal")
    
    def _open_settings(self):
        """打开配置管理窗口"""
        if not self.config_data:
            self.config_data = {
                "drive_letter": "D",
                "target_dir_name": "Target",
                "original_list_filename": "Original file list.csv",
                "log_filename": "log.csv",
                "max_workers": 12,
                "retry_attempts": 3,
                "source_dirs": [],
                "rename_files": False
            }
        
        # 创建配置窗口
        settings_window = SettingsWindow(
            self, 
            self.config_data,
            self._on_settings_saved
        )
        settings_window.focus()
    
    def _on_settings_saved(self, config_data):
        """配置保存后的回调函数"""
        if not self.config_path:
            # 如果之前没有配置文件路径，使用默认路径
            root_path = get_root_path()
            self.config_path = os.path.join(root_path, "config.ini")
        
        # 保存配置
        save_configuration(self.config_path, config_data)
        
        # 重新加载配置
        self.config_data = load_configuration(self.config_path)
        
        if self.config_data:
            # 更新界面上的重命名选项
            self.rename_checkbox_var.set(self.config_data.get("rename_files", False))
            
            # 更新清单文件显示
            list_file = self.config_data.get("list_file")
            if os.path.exists(list_file):
                self.list_file_path = list_file
                self.list_label.configure(text=os.path.basename(list_file))
            
            # 启用开始按钮
            self.start_btn.configure(state="normal")
            
            # 更新配置文件显示
            self.config_label.configure(text=os.path.basename(self.config_path))
    
    def _start_process(self):
        """开始批量复制过程"""
        if not self.config_data or not self.list_file_path:
            messagebox.showwarning("警告", "请先选择配置文件和清单文件")
            return
        
        # 检查目标目录是否存在，不存在则创建
        target_dir = self.config_data.get("target_dir")
        if not os.path.exists(target_dir):
            os.makedirs(target_dir, exist_ok=True)
        
        # 更新按钮状态
        self.running = True
        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.open_target_btn.configure(state="disabled")
        self.view_log_btn.configure(state="disabled")
        
        # 重置进度
        self.progress_bar.set(0)
        self.progress_percent.configure(text="0%")
        self.stats_label.configure(text="已处理: 0 | 成功: 0 | 失败: 0 | 速度: 0 文件/秒")
        
        # 重置停止事件
        self.stop_event.clear()
        
        # 启动工作线程
        self.worker_thread = threading.Thread(
            target=worker,
            args=(self.config_data, self._update_progress, self.stop_event),
            daemon=True
        )
        self.worker_thread.start()
    
    def _stop_process(self):
        """停止批量复制过程"""
        if messagebox.askyesno("确认", "确定要停止当前操作吗？"):
            self.stop_event.set()
            self.stop_btn.configure(state="disabled")
    
    def _update_progress(self, current, total):
        """更新进度条（已集成到队列处理中）"""
        pass
    
    def _open_target_dir(self):
        """打开目标目录"""
        if self.config_data:
            target_dir = self.config_data.get("target_dir")
            if os.path.exists(target_dir):
                if platform.system() == "Windows":
                    os.startfile(target_dir)
                elif platform.system() == "Darwin":
                    subprocess.Popen(["open", target_dir])
                else:
                    subprocess.Popen(["xdg-open", target_dir])
            else:
                messagebox.showwarning("警告", f"目标目录不存在: {target_dir}")
    
    def _view_log(self):
        """查看复制日志"""
        if self.config_data:
            log_file = self.config_data.get("log_file")
            if os.path.exists(log_file):
                if platform.system() == "Windows":
                    os.startfile(log_file)
                elif platform.system() == "Darwin":
                    subprocess.Popen(["open", log_file])
                else:
                    subprocess.Popen(["xdg-open", log_file])
            else:
                messagebox.showwarning("警告", f"日志文件不存在: {log_file}")
    
    def _clear_log(self):
        """清空日志文本框"""
        self.log_textbox.configure(state="normal")
        self.log_textbox.delete("1.0", "end")
        self.log_textbox.configure(state="disabled")
    
    def on_closing(self):
        """窗口关闭事件处理"""
        if self.running:
            if messagebox.askyesno("确认", "当前正在处理文件，确定要退出吗？"):
                self.stop_event.set()
                self.destroy()
        else:
            self.destroy()


if __name__ == "__main__":
    app = BatchCopyGUI()
    app.protocol("WM_DELETE_WINDOW", app.on_closing)
    app.mainloop()