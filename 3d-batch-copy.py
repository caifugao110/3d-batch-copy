import os
import shutil
import csv
import time
import configparser
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from tqdm import tqdm

def load_configuration():
    """
    加载配置文件并解析设置
    
    Returns:
        config: 配置对象
        source_dirs: 完整的源目录路径列表
        target_dir: 目标目录完整路径
        list_file: 原始文件列表完整路径
        log_file: 日志文件完整路径
        max_workers: 最大线程数
        retry_attempts: 重试次数
    """
    # 获取当前脚本所在目录
    current_dir = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(current_dir, "config.ini")
    
    # 检查配置文件是否存在
    if not os.path.exists(config_path):
        print(f"🔥 配置文件不存在: {config_path}")
        print("请确保 config.ini 文件与脚本在同一目录下")
        exit(1)
    
    try:
        # 创建配置解析器
        config = configparser.ConfigParser()
        # 保留选项的大小写
        config.optionxform = lambda option: option
        config.read(config_path, encoding='utf-8')
        
        # 读取基本路径配置
        drive_letter = config.get("Paths", "drive_letter").strip()
        target_dir_name = config.get("Paths", "target_dir_name")
        original_list_file = config.get("Paths", "original_list_file")
        log_file_name = config.get("Paths", "log_file")
        
        # 构建源目录完整路径列表
        source_dirs = []
        for key in config["SourceDirectories"]:
            relative_path = config.get("SourceDirectories", key)
            full_path = f"{drive_letter}:\\{relative_path}"
            source_dirs.append(full_path)
        
        # 构建完整文件路径
        target_dir = os.path.join(current_dir, target_dir_name)
        list_file = os.path.join(current_dir, original_list_file)
        log_file_path = os.path.join(current_dir, log_file_name)
        
        # 读取性能设置
        max_workers = config.getint("Settings", "max_workers", fallback=12)
        retry_attempts = config.getint("Settings", "retry_attempts", fallback=3)
        
        print(f"✅ 配置加载成功:")
        print(f"   驱动器盘符: {drive_letter}")
        print(f"   源目录数量: {len(source_dirs)}")
        print(f"   目标目录: {target_dir}")
        print(f"   最大线程数: {max_workers}")
        print(f"   重试次数: {retry_attempts}")
        
        return config, source_dirs, target_dir, list_file, log_file_path, max_workers, retry_attempts
        
    except Exception as e:
        print(f"🔥 配置文件解析失败: {e}")
        print("请检查 config.ini 文件格式是否正确")
        exit(1)

def clean_filename(name):
    """
    清理文件名：去除特定后缀和标识符，统一转为小写
    
    Args:
        name: 原始文件名（不含扩展名）
    
    Returns:
        str: 清理后的文件名
    """
    # 去除文件名末尾的 'L'（如果有）
    if name.endswith("L"):
        name = name[:-1]
    
    # 如果文件名中包含 'L('，只保留括号前的部分
    if "L(" in name:
        parts = name.split("L(")
        name = parts[0]
    
    # 返回小写格式以进行不区分大小写的匹配
    return name.lower()

def cleanup_target_directory(target_dir):
    """
    清理目标目录中的所有 .step 文件
    
    Args:
        target_dir: 目标目录路径
    """
    print("🧹 正在清理目标目录...")
    clean_count = 0
    
    # 检查目标目录是否存在
    if not os.path.exists(target_dir):
        print(f"📁 目标目录不存在，将创建: {target_dir}")
        os.makedirs(target_dir, exist_ok=True)
        return
    
    # 遍历目标目录中的所有文件
    for file in os.listdir(target_dir):
        # 只处理 .step 文件（不区分大小写）
        if file.lower().endswith(".step"):
            try:
                file_path = os.path.join(target_dir, file)
                os.remove(file_path)
                clean_count += 1
            except Exception as e:
                print(f"⚠️ 删除旧文件失败: {file} - {e}")
    
    print(f"✅ 已清理 {clean_count} 个旧文件")

def build_file_index(source_dirs):
    """
    构建文件索引，按文件名前4个字符分组以提高搜索效率
    
    Args:
        source_dirs: 源目录路径列表
    
    Returns:
        defaultdict: 文件索引字典，格式为 {prefix: [(clean_name, filename, directory), ...]}
    """
    print("⏳ 正在构建全局文件索引...")
    # 使用 defaultdict 自动初始化空列表
    index = defaultdict(list)
    start_time = time.time()
    total_files = 0
    
    # 遍历所有源目录
    for src_dir in source_dirs:
        try:
            # 检查目录是否存在且可访问
            if not os.path.exists(src_dir):
                print(f"⚠️ 路径不存在或无权限访问: {src_dir}")
                continue
            
            # 使用 os.scandir() 进行高效的目录遍历
            with os.scandir(src_dir) as entries:
                for entry in entries:
                    # 只处理 .step 文件且忽略目录
                    if entry.is_file() and entry.name.lower().endswith(".step"):
                        # 分离文件名和扩展名
                        base_name = os.path.splitext(entry.name)[0]
                        # 清理文件名用于匹配
                        clean_base = clean_filename(base_name)
                        # 使用前4个字符作为索引键（如果文件名足够长）
                        prefix_key = clean_base[:4] if len(clean_base) >= 4 else clean_base
                        # 将文件信息添加到索引中
                        index[prefix_key].append((clean_base, entry.name, src_dir))
                        total_files += 1
                        
        except Exception as e:
            print(f"⚠️ 目录扫描失败: {src_dir} - {e}")
    
    index_time = time.time() - start_time
    print(f"✅ 索引构建完成: {len(index)} 个前缀组, {total_files} 个文件, 耗时 {index_time:.2f}秒")
    
    return index

def read_original_file_list(list_file):
    """
    读取原始文件列表 CSV 文件
    
    Args:
        list_file: 原始文件列表路径
    
    Returns:
        list: 文件名列表
    """
    try:
        with open(list_file, "r", encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            # 读取第一列的非空行
            all_lines = [row[0].strip() for row in reader if row and row[0].strip()]
        
        print(f"📋 待处理文件数: {len(all_lines)}")
        return all_lines
        
    except Exception as e:
        print(f"🔥 CSV 文件读取失败: {e}")
        print(f"请检查文件是否存在且格式正确: {list_file}")
        exit(1)

def process_item(item, target_dir, index, retry_attempts):
    """
    处理单个文件的复制任务
    
    Args:
        item: 包含原始文件名和清理后搜索名的元组
        target_dir: 目标目录路径
        index: 文件索引字典
        retry_attempts: 重试次数
    
    Returns:
        dict: 处理结果信息
    """
    original_name, search_name = item
    # 目标文件路径（统一使用 .STEP 扩展名）
    dst_file = os.path.join(target_dir, f"{original_name}.STEP")
    # 使用前4个字符作为搜索键
    prefix_key = search_name[:4] if len(search_name) >= 4 else search_name

    # 在对应前缀组中查找匹配的文件
    if prefix_key in index:
        for clean_base, src_filename, src_dir in index[prefix_key]:
            # 使用 startswith 进行前缀匹配
            if clean_base.startswith(search_name):
                # 重试机制处理可能的临时文件锁定或网络问题
                for attempt in range(retry_attempts):
                    try:
                        src_path = os.path.join(src_dir, src_filename)
                        # 复制文件并保留元数据
                        shutil.copy2(src_path, dst_file)
                        return {
                            "status": "success",
                            "original": original_name,
                            "copied": src_filename,
                            "source": src_dir
                        }
                    except Exception as e:
                        # 如果不是最后一次尝试，则等待后重试
                        if attempt < retry_attempts - 1:
                            # 指数退避策略：等待时间随尝试次数增加
                            time.sleep(2 ** attempt)
                        else:
                            # 最后一次尝试仍然失败，返回错误信息
                            return {
                                "status": "error",
                                "original": original_name,
                                "copied": f"复制失败: {e}",
                                "source": src_dir
                            }
    
    # 没有找到匹配的文件
    return {
        "status": "not_found",
        "original": original_name,
        "copied": "未找到",
        "source": ""
    }

def write_result_log(log_file, result_log):
    """
    将处理结果写入日志文件
    
    Args:
        log_file: 日志文件路径
        result_log: 结果日志列表
    """
    print("📝 正在写入日志文件...")
    try:
        with open(log_file, "w", encoding="utf-8", newline="") as f:
            writer = csv.writer(f)
            # 写入表头
            writer.writerow(["原始文件名", "实际复制文件名", "来源路径"])
            # 写入每条结果
            for res in result_log:
                writer.writerow([res["original"], res["copied"], res["source"]])
        print(f"✅ 日志已保存至: {log_file}")
    except Exception as e:
        print(f"⚠️ 日志文件写入失败: {e}")

def main():
    """主函数"""
    print("=" * 60)
    print("🔧 3D 文件批量复制工具")
    print("=" * 60)
    
    # 记录程序开始时间
    program_start_time = time.time()
    
    # 加载配置
    config, source_dirs, target_dir, list_file, log_file, max_workers, retry_attempts = load_configuration()
    
    # 清理目标目录
    cleanup_target_directory(target_dir)
    
    # 构建文件索引
    index = build_file_index(source_dirs)
    
    # 读取原始文件列表
    original_files = read_original_file_list(list_file)
    total_files = len(original_files)
    
    # 预处理搜索名
    search_items = [(orig, clean_filename(orig)) for orig in original_files]
    
    # 多线程处理文件复制
    print("📦 开始并行复制文件...")
    result_log = []
    found_count = 0
    not_found_count = 0
    copy_errors = 0
    
    # 使用线程池并行处理
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        # 提交所有任务
        futures = [executor.submit(process_item, item, target_dir, index, retry_attempts) 
                  for item in search_items]
        
        # 使用 tqdm 显示进度条
        for future in tqdm(as_completed(futures), total=len(futures), 
                          desc="⏳ 复制中", unit="文件"):
            result = future.result()
            result_log.append(result)
            
            # 统计结果
            if result["status"] == "success":
                found_count += 1
            elif result["status"] == "not_found":
                not_found_count += 1
            elif result["status"] == "error":
                copy_errors += 1
    
    # 写入结果日志
    write_result_log(log_file, result_log)
    
    # 计算总耗时
    total_time = time.time() - program_start_time
    
    # 输出统计信息
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

if __name__ == "__main__":
    main()