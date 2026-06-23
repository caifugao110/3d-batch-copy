from __future__ import annotations

import configparser
import csv
import os
import queue
import re
import shutil
import subprocess
import sys
import threading
import time
import webbrowser
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from tkinter import filedialog, messagebox
import tkinter as tk

import requests
import ttkbootstrap as ttk
from ttkbootstrap.constants import BOTH, LEFT, RIGHT, X, YES


def configure_standard_streams() -> None:
    for stream_name in ("stdout", "stderr"):
        stream = getattr(sys, stream_name, None)
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass


configure_standard_streams()


def bundled_path(name: str) -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / name
    return Path(__file__).resolve().parent / name


def load_project_metadata() -> dict[str, str]:
    pyproject_path = bundled_path("pyproject.toml")
    metadata = {
        "name": "3d-batch-copy",
        "version": "2.0.1",
        "author": "Tobin",
        "homepage": "https://github.com/caifugao110/3d-batch-copy",
    }
    if not pyproject_path.exists():
        return metadata

    text = pyproject_path.read_text(encoding="utf-8")
    try:
        import tomllib

        project = tomllib.loads(text).get("project", {})
        urls = project.get("urls", {})
        authors = project.get("authors", [])
        metadata["name"] = project.get("name", metadata["name"])
        metadata["version"] = project.get("version", metadata["version"])
        if authors:
            metadata["author"] = authors[0].get("name", metadata["author"])
        metadata["homepage"] = urls.get("Homepage", metadata["homepage"])
        return metadata
    except Exception:
        pass

    patterns = {
        "name": r'(?m)^name\s*=\s*"([^"]+)"',
        "version": r'(?m)^version\s*=\s*"([^"]+)"',
        "author": r'authors\s*=\s*\[\{\s*name\s*=\s*"([^"]+)"',
        "homepage": r'(?m)^Homepage\s*=\s*"([^"]+)"',
    }
    for key, pattern in patterns.items():
        match = re.search(pattern, text)
        if match:
            metadata[key] = match.group(1)
    return metadata


PROJECT_METADATA = load_project_metadata()
PROJECT_NAME = PROJECT_METADATA["name"]
__version__ = PROJECT_METADATA["version"]
__author__ = PROJECT_METADATA["author"]
__homepage__ = PROJECT_METADATA["homepage"]

VERSION = f"V{__version__}"
COPYRIGHT = f"{__author__} © 2026"
PROJECT_URL = __homepage__

log_queue: queue.Queue[str] = queue.Queue()
progress_queue: queue.Queue[tuple] = queue.Queue()


def project_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def get_root_path() -> str:
    """获取程序根目录（exe所在目录，支持PyInstaller打包后路径）。"""
    return str(project_root())


def asset_path(name: str) -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / "assets" / name
    return project_root() / "assets" / name


ASSET_ICON = asset_path("app.ico")


def open_path(path: str | Path) -> None:
    target = str(path)
    if sys.platform.startswith("win"):
        os.startfile(target)  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", target])
    else:
        subprocess.Popen(["xdg-open", target])


def center_window(window: tk.Toplevel, parent: tk.Misc, width: int, height: int) -> None:
    parent.update_idletasks()
    x = parent.winfo_rootx() + max((parent.winfo_width() - width) // 2, 0)
    y = parent.winfo_rooty() + max((parent.winfo_height() - height) // 2, 0)
    window.geometry(f"{width}x{height}+{x}+{y}")


def get_update_logs(count=5):
    """从Gitee Releases API获取最近count条更新记录。"""
    api_url = "https://gitee.com/api/v5/repos/caifugao110/3d-batch-copy/releases"
    headers = {
        "Authorization": "token a09da64c1d9e9c7420a18dfd838890b0",
    }
    try:
        response = requests.get(api_url, headers=headers, timeout=5)
        response.raise_for_status()
        releases = response.json()

        version_pattern = re.compile(r"v?(\d+\.\d+\.\d+)", re.IGNORECASE)
        updates = []

        for release in releases:
            tag_name = release.get("tag_name", "")
            match = version_pattern.search(tag_name)
            if match:
                version_str = match.group(1)
                version_tuple = tuple(map(int, version_str.split(".")))
                changelog = "暂无更新说明"
                try:
                    commit_url = f"https://gitee.com/api/v5/repos/caifugao110/3d-batch-copy/commits/{tag_name}"
                    commit_resp = requests.get(commit_url, headers=headers, timeout=5)
                    commit_resp.raise_for_status()
                    commit_data = commit_resp.json()
                    changelog = commit_data.get("commit", {}).get("message", "").strip() or "暂无更新说明"
                except Exception:
                    body = release.get("body", "")
                    match_info = re.search(r"最后提交信息为.*?[:：]\s*(.*)", body, re.DOTALL)
                    if match_info:
                        extracted = match_info.group(1).strip()
                        if extracted:
                            changelog = extracted
                created_at = release.get("created_at", "")[:10] if release.get("created_at") else ""
                updates.append(
                    {
                        "version": tag_name,
                        "version_tuple": version_tuple,
                        "changelog": changelog,
                        "date": created_at,
                    }
                )

        updates.sort(key=lambda x: x["version_tuple"], reverse=True)
        return [{k: v for k, v in item.items() if k != "version_tuple"} for item in updates[:count]]

    except Exception as e:
        print(f"⚠️ 获取更新日志失败: {str(e)}")
        return []


def clean_filename(name):
    """清理文件名: 去除特定后缀和标识符，统一转为小写。"""
    if "-L(" in name:
        parts = name.split("-L(")
        name = parts[0]
    if name.endswith("-L"):
        parts = name.split("-L")
        name = parts[0]
    if name.endswith("L"):
        name = name[:-1]
    if "L(" in name:
        parts = name.split("L(")
        name = parts[0]
    return name.lower()


def load_configuration(config_path):
    """加载配置文件。"""
    if not os.path.exists(config_path):
        print(f"🔥 配置文件不存在: {config_path}")
        print("⚠️ 请确保 config.ini 文件与exe在同一目录下")
        return None

    try:
        config = configparser.ConfigParser()
        config.optionxform = lambda option: option
        config.read(config_path, encoding="utf-8")

        target_dir_name = config.get("Paths", "target_dir_name")
        original_list_filename = config.get("Paths", "original_list_file")
        log_filename = config.get("Paths", "log_file")
        rename_option = config.getboolean("Settings", "rename_files", fallback=False)
        include_xt = config.getboolean("Settings", "include_xt_format", fallback=False)

        root_path = get_root_path()
        source_dirs = []
        for key in config["SourceDirectories"]:
            full_path = config.get("SourceDirectories", key)
            source_dirs.append(full_path)

        target_dir = os.path.join(root_path, target_dir_name)
        list_file = os.path.join(root_path, original_list_filename)
        log_file = os.path.join(root_path, log_filename)

        max_workers = config.getint("Settings", "max_workers", fallback=12)
        retry_attempts = config.getint("Settings", "retry_attempts", fallback=3)

        print("✅ 配置加载成功:")
        print(f"   源目录数量: {len(source_dirs)}")
        print(f"   目标目录: {target_dir}")
        print(f"   待处理列表: {list_file}")
        print(f"   复制日志文件: {log_file}")
        print(f"   最大线程数: {max_workers}")
        print(f"   重试次数: {retry_attempts}")
        print(f"   按清单重命名: {'是' if rename_option else '否'}")
        print(f"   包含 XT 格式: {'是' if include_xt else '否'}")

        return {
            "source_dirs": source_dirs,
            "target_dir": target_dir,
            "list_file": list_file,
            "log_file": log_file,
            "max_workers": max_workers,
            "retry_attempts": retry_attempts,
            "original_list_filename": original_list_filename,
            "target_dir_name": target_dir_name,
            "log_filename": log_filename,
            "config_path": config_path,
            "rename_files": rename_option,
            "include_xt_format": include_xt,
        }
    except Exception as e:
        print(f"🔥 配置文件解析失败: {str(e)}")
        print("⚠️ 请检查 config.ini 文件格式是否正确")
        return None


def save_configuration(config_path, config_data):
    """保存配置文件。"""
    try:
        config = configparser.ConfigParser()
        config.optionxform = lambda option: option

        config["Paths"] = {
            "target_dir_name": config_data.get("target_dir_name", "Target"),
            "original_list_file": config_data.get("original_list_filename", "Original file list.txt"),
            "log_file": config_data.get("log_filename", "log.csv"),
        }

        config["Settings"] = {
            "max_workers": str(config_data.get("max_workers", 12)),
            "retry_attempts": str(config_data.get("retry_attempts", 3)),
            "rename_files": str(config_data.get("rename_files", False)).lower(),
            "include_xt_format": str(config_data.get("include_xt_format", False)).lower(),
        }

        config["SourceDirectories"] = {}
        for idx, src_dir in enumerate(config_data.get("source_dirs", []), 1):
            config["SourceDirectories"][f"source_{idx}"] = src_dir

        with open(config_path, "w", encoding="utf-8") as f:
            config.write(f)

        print(f"✅ 配置已保存至: {config_path}")
        return True
    except Exception as e:
        print(f"🔥 配置保存失败: {str(e)}")
        return False


def apply_runtime_paths(config_data):
    """根据当前程序根目录补齐运行时使用的文件路径。"""
    root_path = get_root_path()
    config_data["target_dir"] = os.path.join(root_path, config_data.get("target_dir_name", "Target"))
    config_data["list_file"] = os.path.join(root_path, config_data.get("original_list_filename", "Original file list.txt"))
    config_data["log_file"] = os.path.join(root_path, config_data.get("log_filename", "log.csv"))
    return config_data


def cleanup_target_directory(target_dir, include_xt=False):
    """清理目标目录中的step/stp和（可选）xt文件（不区分大小写）。"""
    print("🧹 正在清理目标目录...")
    clean_count = 0

    if not os.path.exists(target_dir):
        print(f"📁 目标目录不存在，将创建: {target_dir}")
        os.makedirs(target_dir, exist_ok=True)
        return

    for file in os.listdir(target_dir):
        lower = file.lower()
        try:
            if lower.endswith(".step") or lower.endswith(".stp"):
                file_path = os.path.join(target_dir, file)
                os.remove(file_path)
                clean_count += 1
            elif include_xt and (lower.endswith(".xt") or lower.endswith(".x_t")):
                file_path = os.path.join(target_dir, file)
                os.remove(file_path)
                clean_count += 1
        except Exception as e:
            print(f"⚠️ 删除旧文件失败: {file} - {str(e)}")

    print(f"✅ 已清理 {clean_count} 个旧文件")


def is_xt_variant(filename):
    """判断文件名是否属于XT变体（大小写不敏感）。"""
    lower = filename.lower()
    return lower.endswith(".xt") or lower.endswith(".x_t")


def is_step_variant(filename):
    """判断是否为STEP类（.step 或 .stp 不区分大小写）。"""
    lower = filename.lower()
    return lower.endswith(".step") or lower.endswith(".stp")


def build_file_index(source_dirs, include_xt=False):
    """构建文件索引（支持递归），支持可选包含 XT 格式以及 .stp。"""
    print("⏳ 正在构建全局文件索引（包含子目录）...")
    index = defaultdict(list)
    start_time = time.time()
    total_files = 0

    for src_dir in source_dirs:
        try:
            if not os.path.exists(src_dir):
                print(f"⚠️ 路径不存在或无权限访问: {src_dir}")
                continue
            for root, _, files in os.walk(src_dir):
                for file in files:
                    if is_step_variant(file) or (include_xt and is_xt_variant(file)):
                        base_name = os.path.splitext(file)[0]
                        clean_base = clean_filename(base_name)
                        prefix_key = clean_base[:4] if len(clean_base) >= 4 else clean_base
                        index[prefix_key].append((clean_base, file, root))
                        total_files += 1
        except Exception as e:
            print(f"⚠️ 目录扫描失败: {src_dir} - {str(e)}")

    index_time = time.time() - start_time
    print(f"✅ 索引构建完成: {len(index)} 个前缀组, {total_files} 个文件, 耗时 {index_time:.2f}秒")
    return index


def read_original_file_list(list_file):
    """读取待处理文件列表（支持CSV和TXT格式）。"""
    try:
        _, ext = os.path.splitext(list_file)
        ext = ext.lower()

        if ext == ".csv":
            with open(list_file, "r", encoding="utf-8-sig") as f:
                reader = csv.reader(f)
                all_lines = [row[0].strip() for row in reader if row and row[0].strip()]
        elif ext == ".txt":
            with open(list_file, "r", encoding="utf-8-sig") as f:
                all_lines = [line.strip() for line in f if line.strip()]
        else:
            print(f"⚠️ 不支持的文件格式: {ext}, 仅支持CSV和TXT文件")
            return None

        print(f"📋 待处理文件数: {len(all_lines)}")
        return all_lines
    except Exception as e:
        print(f"🔥 文件读取失败: {str(e)}")
        print(f"⚠️ 请检查文件是否存在且格式正确: {list_file}")
        return None


def process_item(item, target_dir, index, retry_attempts, stop_event, rename_files):
    """处理单个文件复制。"""
    original_name, search_name = item
    dst_file = None
    prefix_key = search_name[:4] if len(search_name) >= 4 else search_name

    if stop_event.is_set():
        return {
            "status": "cancelled",
            "original": original_name,
            "copied": "操作已取消",
            "source": "",
        }

    if prefix_key in index:
        for clean_base, src_filename, src_dir in index[prefix_key]:
            if clean_base == search_name:
                for attempt in range(retry_attempts):
                    if stop_event.is_set():
                        return {
                            "status": "cancelled",
                            "original": original_name,
                            "copied": "操作已取消",
                            "source": "",
                        }

                    try:
                        src_path = os.path.join(src_dir, src_filename)
                        src_ext = os.path.splitext(src_filename)[1]
                        if rename_files:
                            ext_upper = src_ext.upper()
                            dst_file = os.path.join(target_dir, f"{original_name}{ext_upper}")
                        else:
                            if dst_file is None:
                                dst_file = os.path.join(target_dir, src_filename)
                        shutil.copy2(src_path, dst_file)
                        return {
                            "status": "success",
                            "original": original_name,
                            "copied": os.path.basename(dst_file),
                            "source": src_dir,
                            "renamed_to": original_name if rename_files else None,
                        }
                    except Exception as e:
                        if attempt < retry_attempts - 1:
                            time.sleep(2**attempt)
                        else:
                            return {
                                "status": "error",
                                "original": original_name,
                                "copied": f"复制失败: {str(e)}",
                                "source": src_dir,
                            }

    return {
        "status": "not_found",
        "original": original_name,
        "copied": "未找到",
        "source": "",
    }


def worker(config, progress_callback, stop_event):
    """后台工作线程：执行完整的复制流程。"""
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
    rename_files = config.get("rename_files", False)
    include_xt = config.get("include_xt_format", False)

    cleanup_target_directory(target_dir, include_xt=include_xt)
    index = build_file_index(source_dirs, include_xt=include_xt)

    original_files = read_original_file_list(list_file)
    if not original_files or len(original_files) == 0:
        print("🔥 无待处理文件，程序退出")
        progress_queue.put(("complete", False))
        return
    total_files = len(original_files)
    progress_queue.put(("max", total_files))

    search_items = [(orig, clean_filename(orig)) for orig in original_files]

    print(f"📦 开始并行复制文件... {'(将按清单重命名)' if rename_files else ''} {'(包含XT)' if include_xt else ''}")
    executor = ThreadPoolExecutor(max_workers=max_workers)
    futures = [
        executor.submit(process_item, item, target_dir, index, retry_attempts, stop_event, rename_files)
        for item in search_items
    ]

    try:
        for idx, future in enumerate(as_completed(futures)):
            if stop_event.is_set():
                executor.shutdown(wait=False)
                print("⏹️ 正在终止所有复制任务...")
                break

            result = future.result()
            result_log.append(result)

            if result["status"] == "success":
                found_count += 1
            elif result["status"] == "not_found":
                not_found_count += 1
            elif result["status"] == "error":
                copy_errors += 1

            progress_queue.put(
                (
                    "update",
                    idx + 1,
                    found_count,
                    not_found_count + copy_errors,
                    (idx + 1) / max(1, time.time() - program_start_time),
                )
            )
    finally:
        executor.shutdown(wait=False)

    if not stop_event.is_set():
        write_result_log(log_file, result_log)

        total_time = time.time() - program_start_time
        print("\n" + "=" * 60)
        print("📊 处理统计报告")
        print("=" * 60)
        print(f"📊   总文件数: {total_files}")
        print(f"✅   成功复制: {found_count} ({found_count / max(1, total_files):.1%})")
        print(f"❌   未找到: {not_found_count} ({not_found_count / max(1, total_files):.1%})")
        print(f"⚠️   复制错误: {copy_errors}")
        print(f"⏱️   总耗时: {total_time:.1f}秒 | 平均速度: {total_files / max(1, total_time):.1f} 文件/秒")
        print(f"🔧   重命名模式: {'启用' if rename_files else '禁用'}")
        print(f"🔧   包含 XT: {'是' if include_xt else '否'}")
        print("=" * 60)

        failure_rate = (not_found_count + copy_errors) / max(1, total_files)
        if failure_rate > 0.5:
            print(f"\n⚠️ 警告: 超过50%的文件处理失败 ({failure_rate:.1%})！")
            print("⚠️ 可能的原因:")
            print("⚠️   - 网络驱动器连接异常")
            print("⚠️   - 源目录路径不正确")
            print("⚠️   - 文件名不匹配")
            print("⚠️ 请检查配置文件和网络连接状态")

        print("\n🎉 程序执行完成！")
    else:
        print("\n⏹️ 任务已被用户终止")

    progress_queue.put(("complete", not stop_event.is_set()))


def write_result_log(log_file, result_log):
    """写入复制日志文件，使用GBK编码兼容Excel。"""
    print("📝 正在写入复制日志文件...")
    try:
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
    """重定向stdout到GUI日志队列。"""

    def __init__(self):
        self.buffer = []

    def write(self, message):
        self.buffer.append(message)
        if "\n" in message or len("".join(self.buffer)) > 1000:
            self.flush()

    def flush(self):
        if self.buffer:
            log_queue.put("".join(self.buffer))
            self.buffer = []


class SettingsWindow(ttk.Toplevel):
    """配置管理窗口。"""

    def __init__(self, parent, config_data, on_save_callback):
        super().__init__(parent)
        self.title("配置管理")
        self.geometry("760x760")
        self.minsize(660, 620)
        self.config_data = config_data.copy() if config_data else {}
        self.on_save_callback = on_save_callback

        self.transient(parent)
        self.grab_set()
        center_window(self, parent, 760, 760)
        self._build_ui()

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        main = ttk.Frame(self, padding=16)
        main.grid(row=0, column=0, sticky="nsew")
        main.columnconfigure(0, weight=1)
        main.rowconfigure(2, weight=1)

        basic = ttk.Labelframe(main, text="基本设置", padding=12)
        basic.grid(row=0, column=0, sticky="ew")
        basic.columnconfigure(1, weight=1)

        self.target_entry = self._add_entry_row(basic, 0, "目标目录名", self.config_data.get("target_dir_name", "Target"))
        self.list_entry = self._add_entry_row(
            basic,
            1,
            "原始清单文件",
            self.config_data.get("original_list_filename", "Original file list.txt"),
        )
        self.log_entry = self._add_entry_row(basic, 2, "复制日志文件名", self.config_data.get("log_filename", "log.csv"))

        perf = ttk.Labelframe(main, text="性能设置", padding=12)
        perf.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        perf.columnconfigure(1, weight=1)

        self.workers_entry = self._add_entry_row(perf, 0, "最大线程数", str(self.config_data.get("max_workers", 12)))
        self.retry_entry = self._add_entry_row(perf, 1, "重试次数", str(self.config_data.get("retry_attempts", 3)))

        self.rename_var = tk.BooleanVar(value=self.config_data.get("rename_files", False))
        self.include_xt_var = tk.BooleanVar(value=self.config_data.get("include_xt_format", False))
        ttk.Checkbutton(
            perf,
            text="按照清单重命名3D文件",
            variable=self.rename_var,
            bootstyle="round-toggle",
        ).grid(row=2, column=0, columnspan=2, sticky="w", pady=(8, 0))
        ttk.Checkbutton(
            perf,
            text="包含 XT 格式3D文件",
            variable=self.include_xt_var,
            bootstyle="round-toggle",
        ).grid(row=3, column=0, columnspan=2, sticky="w", pady=(6, 0))

        source = ttk.Labelframe(main, text="源目录管理", padding=12)
        source.grid(row=2, column=0, sticky="nsew", pady=(12, 0))
        source.columnconfigure(0, weight=1)
        source.rowconfigure(0, weight=1)

        text_outer = ttk.Frame(source)
        text_outer.grid(row=0, column=0, sticky="nsew")
        text_outer.columnconfigure(0, weight=1)
        text_outer.rowconfigure(0, weight=1)

        self.source_text = tk.Text(
            text_outer,
            height=12,
            wrap="none",
            font=("Microsoft YaHei UI", 10),
            relief="solid",
            borderwidth=1,
        )
        self.source_text.grid(row=0, column=0, sticky="nsew")
        yscroll = ttk.Scrollbar(text_outer, orient=tk.VERTICAL, command=self.source_text.yview)
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll = ttk.Scrollbar(text_outer, orient=tk.HORIZONTAL, command=self.source_text.xview)
        xscroll.grid(row=1, column=0, sticky="ew")
        self.source_text.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

        for src_dir in self.config_data.get("source_dirs", []):
            self.source_text.insert(tk.END, src_dir + "\n")

        source_buttons = ttk.Frame(source)
        source_buttons.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        ttk.Button(source_buttons, text="添加目录", bootstyle="secondary-outline", command=self._add_source_dir).pack(
            side=LEFT, padx=(0, 8)
        )
        ttk.Button(source_buttons, text="删除当前行", bootstyle="secondary-outline", command=self._remove_source_dir).pack(
            side=LEFT, padx=(0, 8)
        )
        ttk.Button(source_buttons, text="清空全部", bootstyle="secondary-outline", command=self._clear_source_dirs).pack(
            side=LEFT
        )

        footer = ttk.Frame(self, padding=(16, 0, 16, 16))
        footer.grid(row=1, column=0, sticky="ew")
        ttk.Button(footer, text="取消", bootstyle="secondary-outline", command=self.destroy).pack(side=RIGHT)
        ttk.Button(footer, text="保存配置", bootstyle="success", command=self._save_config).pack(side=RIGHT, padx=(0, 8))

    def _add_entry_row(self, parent, row: int, label: str, value: str):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=(0, 10), pady=4)
        entry = ttk.Entry(parent)
        entry.grid(row=row, column=1, sticky="ew", pady=4)
        entry.insert(0, value)
        return entry

    def _add_source_dir(self):
        dir_path = filedialog.askdirectory(title="选择源目录")
        if dir_path:
            self.source_text.insert(tk.END, dir_path + "\n")

    def _remove_source_dir(self):
        try:
            start = self.source_text.index("insert linestart")
            end = self.source_text.index("insert lineend +1c")
            self.source_text.delete(start, end)
        except Exception:
            messagebox.showwarning("警告", "请先选择要删除的目录")

    def _clear_source_dirs(self):
        if messagebox.askyesno("确认", "确定要清空所有源目录吗？"):
            self.source_text.delete("1.0", tk.END)

    def _save_config(self):
        try:
            max_workers = int(self.workers_entry.get())
            retry_attempts = int(self.retry_entry.get())

            if max_workers < 1 or retry_attempts < 1:
                messagebox.showerror("错误", "线程数和重试次数必须大于0")
                return

            content = self.source_text.get("1.0", "end").strip()
            source_dirs = [line.strip() for line in content.split("\n") if line.strip()] if content else []

            self.config_data["target_dir_name"] = self.target_entry.get().strip()
            self.config_data["original_list_filename"] = self.list_entry.get().strip()
            self.config_data["log_filename"] = self.log_entry.get().strip()
            self.config_data["max_workers"] = max_workers
            self.config_data["retry_attempts"] = retry_attempts
            self.config_data["source_dirs"] = source_dirs
            self.config_data["rename_files"] = self.rename_var.get()
            self.config_data["include_xt_format"] = self.include_xt_var.get()
            apply_runtime_paths(self.config_data)

            if self.on_save_callback:
                self.on_save_callback(self.config_data)

            messagebox.showinfo("成功", "配置已保存并重新加载")
            self.destroy()
        except ValueError:
            messagebox.showerror("错误", "线程数和重试次数必须是整数")


class ListManagerWindow(ttk.Toplevel):
    """清单管理窗口。"""

    def __init__(self, parent, list_file_path, on_save_callback):
        super().__init__(parent)
        self.title("清单管理")
        self.geometry("820x620")
        self.minsize(680, 500)
        self.list_file_path = list_file_path
        self.on_save_callback = on_save_callback

        self.transient(parent)
        self.grab_set()
        center_window(self, parent, 820, 620)
        self._build_ui()
        self._load_file_content()
        self.protocol("WM_DELETE_WINDOW", self._on_closing)

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        header = ttk.Frame(self, padding=(16, 14, 16, 8))
        header.grid(row=0, column=0, sticky="ew")
        ttk.Label(
            header,
            text=f"编辑清单文件: {os.path.basename(self.list_file_path)}",
            font=("Microsoft YaHei UI", 15, "bold"),
        ).pack(anchor="w")

        text_outer = ttk.Frame(self, padding=(16, 0, 16, 12))
        text_outer.grid(row=1, column=0, sticky="nsew")
        text_outer.columnconfigure(0, weight=1)
        text_outer.rowconfigure(0, weight=1)

        self.text_editor = tk.Text(text_outer, wrap="word", font=("Microsoft YaHei UI", 11), relief="solid", borderwidth=1)
        self.text_editor.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(text_outer, orient=tk.VERTICAL, command=self.text_editor.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.text_editor.configure(yscrollcommand=scrollbar.set)

        footer = ttk.Frame(self, padding=(16, 0, 16, 16))
        footer.grid(row=2, column=0, sticky="ew")
        ttk.Button(footer, text="取消", bootstyle="secondary-outline", command=self.destroy).pack(side=RIGHT)
        ttk.Button(footer, text="保存并退出", bootstyle="success", command=self._save_and_exit).pack(side=RIGHT, padx=(0, 8))

    def _load_file_content(self):
        try:
            if os.path.exists(self.list_file_path):
                with open(self.list_file_path, "r", encoding="utf-8-sig") as f:
                    content = f.read()
                self.text_editor.insert("1.0", content)
        except Exception as e:
            messagebox.showerror("错误", f"加载文件失败: {str(e)}")
            self.destroy()

    def _save_and_exit(self):
        try:
            content = self.text_editor.get("1.0", "end-1c")
            with open(self.list_file_path, "w", encoding="utf-8-sig") as f:
                f.write(content)

            if self.on_save_callback:
                self.on_save_callback()

            messagebox.showinfo("成功", "清单文件已保存")
            self.destroy()
        except Exception as e:
            messagebox.showerror("错误", f"保存文件失败: {str(e)}")

    def _on_closing(self):
        if messagebox.askyesno("确认", "确定要退出吗？未保存的更改将丢失。"):
            self.destroy()


class UpdateLogWindow(ttk.Toplevel):
    """更新日志窗口。"""

    def __init__(self, parent):
        super().__init__(parent)
        self.title("更新日志")
        self.geometry("640x520")
        self.minsize(560, 420)

        self.transient(parent)
        self.grab_set()
        center_window(self, parent, 640, 520)
        self._build_ui()
        self._load_update_logs()

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        header = ttk.Frame(self, padding=(16, 14, 16, 8))
        header.grid(row=0, column=0, sticky="ew")
        ttk.Label(header, text="更新日志", font=("Microsoft YaHei UI", 16, "bold")).pack(anchor="w")

        body = ttk.Frame(self, padding=(16, 0, 16, 12))
        body.grid(row=1, column=0, sticky="nsew")
        body.columnconfigure(0, weight=1)
        body.rowconfigure(0, weight=1)

        self.log_textbox = tk.Text(body, wrap="word", font=("Microsoft YaHei UI", 10), state="disabled", relief="solid")
        self.log_textbox.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(body, orient=tk.VERTICAL, command=self.log_textbox.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_textbox.configure(yscrollcommand=scrollbar.set)

        footer = ttk.Frame(self, padding=(16, 0, 16, 16))
        footer.grid(row=2, column=0, sticky="ew")
        self.loading_var = tk.StringVar(value="正在获取更新日志...")
        ttk.Label(footer, textvariable=self.loading_var, bootstyle="secondary").pack(side=LEFT)
        ttk.Button(footer, text="关闭", bootstyle="primary", command=self.destroy).pack(side=RIGHT)

    def _load_update_logs(self):
        def fetch_logs():
            logs = get_update_logs(5)
            self.after(0, lambda: self._display_logs(logs))

        threading.Thread(target=fetch_logs, daemon=True).start()

    def _display_logs(self, logs):
        if not self.winfo_exists():
            return
        self.loading_var.set("")
        self.log_textbox.configure(state="normal")
        self.log_textbox.delete("1.0", tk.END)

        if not logs:
            self.log_textbox.insert(tk.END, "无法获取更新日志，请检查网络连接。")
            self.log_textbox.configure(state="disabled")
            return

        for log in logs:
            version = log.get("version", "")
            date = log.get("date", "")
            changelog = log.get("changelog", "")

            self.log_textbox.insert(tk.END, f"【版本 {version}】")
            self.log_textbox.insert(tk.END, f" ({date})\n" if date else "\n")
            self.log_textbox.insert(tk.END, "=" * 50 + "\n")
            self.log_textbox.insert(tk.END, f"{changelog}\n\n")

        self.log_textbox.configure(state="disabled")


class HelpWindow(ttk.Toplevel):
    """使用说明窗口 — 从 Gitee 加载 README.md 内容。"""

    README_URL = "https://gitee.com/caifugao110/3d-batch-copy/raw/master/README.md"

    def __init__(self, parent):
        super().__init__(parent)
        self.title("使用说明")
        self.geometry("780x620")
        self.minsize(640, 480)

        self.transient(parent)
        self.grab_set()
        center_window(self, parent, 780, 620)
        self._build_ui()
        self._load_readme()

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        header = ttk.Frame(self, padding=(16, 14, 16, 8))
        header.grid(row=0, column=0, sticky="ew")
        ttk.Label(header, text="使用说明", font=("Microsoft YaHei UI", 16, "bold")).pack(anchor="w")

        body = ttk.Frame(self, padding=(16, 0, 16, 12))
        body.grid(row=1, column=0, sticky="nsew")
        body.columnconfigure(0, weight=1)
        body.rowconfigure(0, weight=1)

        self.textbox = tk.Text(body, wrap="word", font=("Microsoft YaHei UI", 10), state="disabled", relief="solid")
        self.textbox.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(body, orient=tk.VERTICAL, command=self.textbox.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.textbox.configure(yscrollcommand=scrollbar.set)

        footer = ttk.Frame(self, padding=(16, 0, 16, 16))
        footer.grid(row=2, column=0, sticky="ew")
        self.loading_var = tk.StringVar(value="正在加载使用说明...")
        ttk.Label(footer, textvariable=self.loading_var, bootstyle="secondary").pack(side=LEFT)
        ttk.Button(footer, text="关闭", bootstyle="primary", command=self.destroy).pack(side=RIGHT)

    def _load_readme(self):
        def fetch():
            try:
                resp = requests.get(self.README_URL, timeout=10)
                resp.raise_for_status()
                content = resp.text
            except Exception as e:
                content = f"无法加载使用说明。\n\n错误信息: {str(e)}\n\n请访问项目主页查看: https://github.com/caifugao110/3d-batch-copy"
            self.after(0, lambda: self._display_content(content))

        threading.Thread(target=fetch, daemon=True).start()

    def _display_content(self, content):
        if not self.winfo_exists():
            return
        self.loading_var.set("")
        self.textbox.configure(state="normal")
        self.textbox.delete("1.0", tk.END)
        self.textbox.insert(tk.END, content)
        self.textbox.configure(state="disabled")


class BatchCopyApp(ttk.Window):
    """3D文件批量复制GUI界面。"""

    def __init__(self):
        super().__init__(themename="yeti")
        self.title(f"3d-batch-copy {VERSION}")
        self.geometry("1240x800")
        self.minsize(1000, 680)
        if ASSET_ICON.exists():
            self.iconbitmap(str(ASSET_ICON))

        self.config_path = None
        self.list_file_path = None
        self.config_data = None
        self.running = False
        self.stop_event = threading.Event()
        self.worker_thread = None
        self.total_files = 0
        self.success_count = 0
        self.failure_count = 0
        self.start_time = 0
        self.original_stdout = sys.stdout
        self._closing = False

        self.theme_var = tk.StringVar(value="yeti")
        self.config_label_var = tk.StringVar(value="未选择")
        self.list_label_var = tk.StringVar(value="未选择")
        self.rename_checkbox_var = tk.BooleanVar(value=False)
        self.include_xt_checkbox_var = tk.BooleanVar(value=False)
        self.progress_percent_var = tk.StringVar(value="0%")
        self.stats_var = tk.StringVar(value="已处理: 0 | 成功: 0 | 失败: 0 | 速度: 0 文件/秒")
        self.status_var = tk.StringVar(value="正在初始化")

        self._build_ui()
        self._redirect_stdout()
        self._listen_queues()
        self.after(200, self._auto_load_files)

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        header = ttk.Frame(self, padding=(18, 14, 18, 8))
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)

        title = ttk.Label(header, text="3d-batch-copy", font=("Microsoft YaHei UI", 22, "bold"))
        title.grid(row=0, column=0, sticky="w")
        meta = ttk.Label(header, text=f"{VERSION} 3D文件批量复制工具", bootstyle="secondary")
        meta.grid(row=1, column=0, sticky="w", pady=(2, 0))

        theme_bar = ttk.Frame(header)
        theme_bar.grid(row=0, column=1, rowspan=2, sticky="e")
        ttk.Label(theme_bar, text="主题").pack(side=LEFT, padx=(0, 8))
        theme_box = ttk.Combobox(
            theme_bar,
            textvariable=self.theme_var,
            values=sorted(self.style.theme_names()),
            width=16,
            state="readonly",
        )
        theme_box.pack(side=LEFT)
        theme_box.bind("<<ComboboxSelected>>", self._change_theme)
        ttk.Button(theme_bar, text="GitHub", bootstyle="secondary-outline", command=lambda: webbrowser.open(PROJECT_URL)).pack(
            side=LEFT, padx=(8, 0)
        )
        ttk.Button(theme_bar, text="使用说明", bootstyle="secondary-outline", command=self._show_help).pack(
            side=LEFT, padx=(8, 0)
        )
        ttk.Button(theme_bar, text="更新日志", bootstyle="secondary-outline", command=self._show_update_log).pack(
            side=LEFT, padx=(8, 0)
        )
        ttk.Button(theme_bar, text="关于", bootstyle="secondary-outline", command=self.show_about).pack(side=LEFT, padx=(8, 0))

        main = ttk.Panedwindow(self, orient=tk.HORIZONTAL)
        main.grid(row=1, column=0, sticky="nsew", padx=16, pady=(0, 12))

        controls = ttk.Frame(main, padding=12)
        main.add(controls, weight=1)
        results = ttk.Frame(main, padding=(10, 12, 12, 12))
        main.add(results, weight=4)

        self._build_controls(controls)
        self._build_results(results)

        footer = ttk.Frame(self, padding=(18, 0, 18, 14))
        footer.grid(row=2, column=0, sticky="ew")
        footer.columnconfigure(0, weight=1)
        ttk.Label(footer, textvariable=self.status_var, bootstyle="secondary").grid(row=0, column=0, sticky="w")

    def _build_controls(self, parent):
        parent.columnconfigure(0, weight=1)

        file_box = ttk.Labelframe(parent, text="文件设置", padding=12)
        file_box.grid(row=0, column=0, sticky="ew")
        file_box.columnconfigure(0, weight=1)

        self._add_file_picker(file_box, "配置文件", self.config_label_var, self._select_config, 0)
        self._add_file_picker(file_box, "原始清单", self.list_label_var, self._select_list_file, 1)

        option_box = ttk.Labelframe(parent, text="选项", padding=12)
        option_box.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        ttk.Checkbutton(
            option_box,
            text="按照清单重命名3D文件",
            variable=self.rename_checkbox_var,
            bootstyle="round-toggle",
            command=self._on_rename_checkbox_change,
        ).pack(anchor="w", pady=3)
        ttk.Checkbutton(
            option_box,
            text="包含 XT 格式3D文件",
            variable=self.include_xt_checkbox_var,
            bootstyle="round-toggle",
            command=self._on_include_xt_change,
        ).pack(anchor="w", pady=3)

        action_box = ttk.Labelframe(parent, text="执行", padding=12)
        action_box.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        action_box.columnconfigure(0, weight=1)

        self.start_btn = ttk.Button(action_box, text="开始批量复制", bootstyle="success", command=self._start_process)
        self.start_btn.grid(row=0, column=0, sticky="ew")
        self.stop_btn = ttk.Button(action_box, text="停止处理", bootstyle="danger", command=self._stop_process, state="disabled")
        self.stop_btn.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(action_box, text="配置管理", bootstyle="secondary-outline", command=self._open_settings).grid(
            row=2, column=0, sticky="ew", pady=(8, 0)
        )
        self.list_manager_btn = ttk.Button(
            action_box,
            text="清单管理",
            bootstyle="secondary-outline",
            command=self._open_list_manager,
            state="disabled",
        )
        self.list_manager_btn.grid(row=3, column=0, sticky="ew", pady=(8, 0))
        self.open_target_btn = ttk.Button(
            action_box,
            text="打开目标目录",
            bootstyle="secondary-outline",
            command=self._open_target_dir,
            state="disabled",
        )
        self.open_target_btn.grid(row=4, column=0, sticky="ew", pady=(8, 0))
        self.view_log_btn = ttk.Button(
            action_box,
            text="查看复制日志",
            bootstyle="secondary-outline",
            command=self._view_log,
            state="disabled",
        )
        self.view_log_btn.grid(row=5, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(action_box, text="清空日志框", bootstyle="secondary-outline", command=self._clear_log).grid(
            row=6, column=0, sticky="ew", pady=(8, 0)
        )

    def _add_file_picker(self, parent, label, variable, command, row):
        group = ttk.Frame(parent)
        group.grid(row=row, column=0, sticky="ew", pady=(0, 10) if row == 0 else (0, 0))
        group.columnconfigure(0, weight=1)
        ttk.Label(group, text=label, font=("Microsoft YaHei UI", 10, "bold")).grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(group, textvariable=variable, bootstyle="secondary", anchor="w", padding=(8, 8), relief="solid").grid(
            row=1, column=0, sticky="ew", pady=(4, 0), padx=(0, 8)
        )
        ttk.Button(group, text="选择", bootstyle="secondary-outline", command=command).grid(row=1, column=1, sticky="e", pady=(4, 0))

    def _build_results(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)

        progress_box = ttk.Labelframe(parent, text="处理进度", padding=12)
        progress_box.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        progress_box.columnconfigure(0, weight=1)

        progress_row = ttk.Frame(progress_box)
        progress_row.grid(row=0, column=0, sticky="ew")
        progress_row.columnconfigure(0, weight=1)
        self.progress_bar = ttk.Progressbar(progress_row, mode="determinate", maximum=100, value=0)
        self.progress_bar.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        ttk.Label(progress_row, textvariable=self.progress_percent_var, width=6, anchor="e").grid(row=0, column=1)
        ttk.Label(progress_box, textvariable=self.stats_var, bootstyle="secondary").grid(row=1, column=0, sticky="w", pady=(8, 0))

        log_box = ttk.Labelframe(parent, text="处理日志", padding=12)
        log_box.grid(row=1, column=0, sticky="nsew")
        log_box.columnconfigure(0, weight=1)
        log_box.rowconfigure(0, weight=1)

        self.log_textbox = tk.Text(
            log_box,
            wrap="word",
            font=("Microsoft YaHei UI", 11),
            state="disabled",
            relief="solid",
            borderwidth=1,
        )
        self.log_textbox.grid(row=0, column=0, sticky="nsew")
        yscroll = ttk.Scrollbar(log_box, orient=tk.VERTICAL, command=self.log_textbox.yview)
        yscroll.grid(row=0, column=1, sticky="ns")
        self.log_textbox.configure(yscrollcommand=yscroll.set)

    def _on_rename_checkbox_change(self):
        if self.config_data and self.config_path:
            self.config_data["rename_files"] = self.rename_checkbox_var.get()
            save_configuration(self.config_path, self.config_data)
            self._clear_log()
            self.config_data = load_configuration(self.config_path)

    def _on_include_xt_change(self):
        if self.config_data and self.config_path:
            self.config_data["include_xt_format"] = self.include_xt_checkbox_var.get()
            save_configuration(self.config_path, self.config_data)
            self._clear_log()
            self.config_data = load_configuration(self.config_path)

    def _change_theme(self, _=None):
        self.style.theme_use(self.theme_var.get())

    def _redirect_stdout(self):
        self.original_stdout = sys.stdout
        sys.stdout = StdoutRedirector()

    def _drain_queues(self):
        max_messages_per_batch = 50
        log_count = 0

        while not log_queue.empty() and log_count < max_messages_per_batch:
            message = log_queue.get()
            self.log_textbox.configure(state="normal")
            self.log_textbox.insert(tk.END, message)
            self.log_textbox.see(tk.END)
            self.log_textbox.configure(state="disabled")
            log_count += 1

        while not progress_queue.empty():
            item = progress_queue.get()
            if item[0] == "max":
                self.total_files = item[1]
                self.success_count = 0
                self.failure_count = 0
                self.start_time = time.time()
                self.progress_bar.configure(value=0)
                self.progress_percent_var.set("0%")
                self.stats_var.set("已处理: 0 | 成功: 0 | 失败: 0 | 速度: 0 文件/秒")
            elif item[0] == "update":
                current = item[1]
                self.success_count = item[2]
                self.failure_count = item[3]
                speed = item[4]

                if self.total_files > 0:
                    percentage = (current / self.total_files) * 100
                    self.progress_bar.configure(value=percentage)
                    self.progress_percent_var.set(f"{int(percentage)}%")

                self.stats_var.set(
                    f"已处理: {current} | 成功: {self.success_count} | 失败: {self.failure_count} | 速度: {speed:.1f} 文件/秒"
                )
            elif item[0] == "complete":
                if hasattr(sys.stdout, "flush"):
                    sys.stdout.flush()

                self.running = False
                self.start_btn.configure(state="normal")
                self.stop_btn.configure(state="disabled")
                self.list_manager_btn.configure(state="normal")
                self.status_var.set("任务完成" if item[1] else "任务已停止或失败")

                if item[1]:
                    self.open_target_btn.configure(state="normal")
                    self.view_log_btn.configure(state="normal")

    def _listen_queues(self):
        if self._closing:
            return
        self._drain_queues()
        delay = 20 if not log_queue.empty() else 100
        self.after(delay, self._listen_queues)

    def _auto_load_files(self):
        self._clear_log()
        root_path = get_root_path()
        default_config = os.path.join(root_path, "config.ini")

        if os.path.exists(default_config):
            self.config_path = default_config
            self.config_label_var.set(os.path.basename(default_config))
            self.config_data = load_configuration(default_config)

            if self.config_data:
                self.rename_checkbox_var.set(self.config_data.get("rename_files", False))
                self.include_xt_checkbox_var.set(self.config_data.get("include_xt_format", False))

                default_list = self.config_data.get("list_file")
                if os.path.exists(default_list):
                    self._clear_log()
                    self.list_file_path = default_list
                    self.list_label_var.set(os.path.basename(default_list))
                    self.list_manager_btn.configure(state="normal")

                self.start_btn.configure(state="normal")
                self.status_var.set("配置已加载")
        else:
            self.status_var.set("请选择配置文件")

    def _select_config(self):
        file_path = filedialog.askopenfilename(title="选择配置文件", filetypes=[("INI文件", "*.ini"), ("所有文件", "*.*")])

        if file_path:
            self._clear_log()
            self.config_path = file_path
            self.config_label_var.set(os.path.basename(file_path))
            self.config_data = load_configuration(file_path)

            if self.config_data:
                self.rename_checkbox_var.set(self.config_data.get("rename_files", False))
                self.include_xt_checkbox_var.set(self.config_data.get("include_xt_format", False))

                list_file = self.config_data.get("list_file")
                if os.path.exists(list_file):
                    self._clear_log()
                    self.list_file_path = list_file
                    self.list_label_var.set(os.path.basename(list_file))
                    self.list_manager_btn.configure(state="normal")

                self.start_btn.configure(state="normal")
                self.status_var.set("配置已加载")
            else:
                self.start_btn.configure(state="disabled")
                self.status_var.set("配置加载失败")

    def _select_list_file(self):
        file_path = filedialog.askopenfilename(
            title="选择原始清单文件",
            filetypes=[("TXT文件", "*.txt"), ("CSV文件", "*.csv"), ("所有文件", "*.*")],
        )

        if file_path:
            self._clear_log()
            self.list_file_path = file_path
            self.list_label_var.set(os.path.basename(file_path))
            self.list_manager_btn.configure(state="normal")

            if self.config_data and self.config_path:
                self.config_data["original_list_filename"] = os.path.basename(file_path)
                self.config_data["list_file"] = file_path
                save_configuration(self.config_path, self.config_data)
                self.config_data = load_configuration(self.config_path)
                if self.config_data:
                    self.config_data["list_file"] = file_path
                    self.rename_checkbox_var.set(self.config_data.get("rename_files", False))
                    self.include_xt_checkbox_var.set(self.config_data.get("include_xt_format", False))

            if self.config_data:
                self.start_btn.configure(state="normal")

    def _open_settings(self):
        if not self.config_data:
            self.config_data = {
                "target_dir_name": "Target",
                "original_list_filename": "Original file list.txt",
                "log_filename": "log.csv",
                "max_workers": 12,
                "retry_attempts": 3,
                "source_dirs": [],
                "rename_files": False,
                "include_xt_format": False,
            }

        settings_window = SettingsWindow(self, self.config_data, self._on_settings_saved)
        settings_window.focus()

    def _on_settings_saved(self, config_data):
        self._clear_log()
        if not self.config_path:
            self.config_path = os.path.join(get_root_path(), "config.ini")

        save_configuration(self.config_path, config_data)
        self.config_data = load_configuration(self.config_path)

        if self.config_data:
            self.rename_checkbox_var.set(self.config_data.get("rename_files", False))
            self.include_xt_checkbox_var.set(self.config_data.get("include_xt_format", False))

            list_file = self.config_data.get("list_file")
            if os.path.exists(list_file):
                self._clear_log()
                self.list_file_path = list_file
                self.list_label_var.set(os.path.basename(list_file))
                self.list_manager_btn.configure(state="normal")

            self.start_btn.configure(state="normal")
            self.config_label_var.set(os.path.basename(self.config_path))
            self.status_var.set("配置已保存")

    def _open_list_manager(self):
        if not self.list_file_path:
            messagebox.showwarning("警告", "请先选择清单文件")
            return

        list_manager_window = ListManagerWindow(self, self.list_file_path, self._on_list_saved)
        list_manager_window.focus()

    def _on_list_saved(self):
        print(f"✅ 清单文件已保存: {self.list_file_path}")
        self._clear_log()
        print("🔄 正在重新加载清单文件...")

        if self.config_data:
            original_files = read_original_file_list(self.list_file_path)
            if original_files:
                print(f"✅ 清单文件重新加载成功，共 {len(original_files)} 个文件")
            else:
                print("⚠️ 清单文件重新加载失败")

    def _start_process(self):
        if not self.config_data or not self.list_file_path:
            messagebox.showwarning("警告", "请先选择配置文件和清单文件")
            return

        self._clear_log()

        target_dir = self.config_data.get("target_dir")
        if not os.path.exists(target_dir):
            os.makedirs(target_dir, exist_ok=True)

        self.running = True
        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.open_target_btn.configure(state="disabled")
        self.view_log_btn.configure(state="disabled")
        self.list_manager_btn.configure(state="disabled")

        self.progress_bar.configure(value=0)
        self.progress_percent_var.set("0%")
        self.stats_var.set("已处理: 0 | 成功: 0 | 失败: 0 | 速度: 0 文件/秒")
        self.status_var.set("任务运行中")
        self.stop_event.clear()

        if self.config_data is not None:
            self.config_data["include_xt_format"] = self.include_xt_checkbox_var.get()
            self.config_data["rename_files"] = self.rename_checkbox_var.get()
            self.config_data["list_file"] = self.list_file_path
            if self.config_path:
                save_configuration(self.config_path, self.config_data)

        self.worker_thread = threading.Thread(
            target=worker,
            args=(self.config_data, self._update_progress, self.stop_event),
            daemon=True,
        )
        self.worker_thread.start()

    def _stop_process(self):
        if messagebox.askyesno("确认", "确定要停止当前操作吗？"):
            self.stop_event.set()
            self.stop_btn.configure(state="disabled")
            self.status_var.set("正在停止任务")

    def _update_progress(self, current, total):
        pass

    def _open_target_dir(self):
        if self.config_data:
            target_dir = self.config_data.get("target_dir")
            if os.path.exists(target_dir):
                open_path(target_dir)
            else:
                messagebox.showwarning("警告", f"目标目录不存在: {target_dir}")

    def _view_log(self):
        if self.config_data:
            log_file = self.config_data.get("log_file")
            if os.path.exists(log_file):
                open_path(log_file)
            else:
                messagebox.showwarning("警告", f"日志文件不存在: {log_file}")

    def _show_update_log(self):
        update_log_window = UpdateLogWindow(self)
        update_log_window.focus()

    def _show_help(self):
        help_window = HelpWindow(self)
        help_window.focus()

    def _clear_log(self):
        try:
            self.log_textbox.configure(state="normal")
            self.log_textbox.delete("1.0", tk.END)
            self.log_textbox.configure(state="disabled")
        except Exception:
            pass

    def show_about(self):
        dialog = ttk.Toplevel(self)
        dialog.title("关于 3d-batch-copy")
        dialog.geometry("500x280")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()
        center_window(dialog, self, 500, 280)

        container = ttk.Frame(dialog, padding=22)
        container.pack(fill=BOTH, expand=YES)
        ttk.Label(container, text="3d-batch-copy", font=("Microsoft YaHei UI", 18, "bold")).pack(anchor="w")
        ttk.Label(container, text=f"版本 {VERSION}", bootstyle="secondary").pack(anchor="w", pady=(4, 0))
        ttk.Label(container, text=f"作者：{__author__}", bootstyle="secondary").pack(anchor="w", pady=(8, 0))
        ttk.Label(container, text="开源协议：MIT", bootstyle="secondary").pack(anchor="w", pady=(4, 0))
        link = ttk.Label(container, text=PROJECT_URL, bootstyle="primary", cursor="hand2")
        link.pack(anchor="w", pady=(14, 0))
        link.bind("<Button-1>", lambda _: webbrowser.open(PROJECT_URL))
        ttk.Button(container, text="关闭", bootstyle="primary", command=dialog.destroy).pack(anchor="e", pady=(24, 0))

    def on_closing(self):
        if self.running:
            if not messagebox.askyesno("确认", "当前正在处理文件，确定要退出吗？"):
                return
            self.stop_event.set()

        self._closing = True
        if hasattr(sys.stdout, "flush"):
            sys.stdout.flush()
        self._drain_queues()
        sys.stdout = self.original_stdout
        self.destroy()


def main() -> None:
    app = BatchCopyApp()
    app.protocol("WM_DELETE_WINDOW", app.on_closing)
    app.mainloop()


if __name__ == "__main__":
    main()
