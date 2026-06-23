
<p align="center">
  <h1 align="center">3d-batch-copy</h1>
  <p align="center">
    <img src="https://img.shields.io/github/v/tag/caifugao110/3d-batch-copy?label=version" alt="version">
    <img src="https://img.shields.io/badge/python-%3E%3D3.9-green" alt="python">
    <img src="https://img.shields.io/badge/license-MIT-yellow" alt="license">
    <img src="https://img.shields.io/badge/platform-Windows-lightgrey" alt="platform">
  </p>
  <p align="center">
    <i>3D文件批量复制工具 — 从网络或本地目录中快速查找并批量复制3D模型文件。</i>
  </p>
</p>

---

## 简介

**3d-batch-copy** 是一款基于 Python 的桌面 GUI 工具，专为 OBARA 有调图需求的人员打造，用于从多个网络或本地目录中快速查找并批量复制 3D 模型文件到指定目标目录。

| 项目信息 |  |
| --- | --- |
| 作者 | **Tobin** |
| 项目地址 | [github.com/caifugao110/3d-batch-copy](https://github.com/caifugao110/3d-batch-copy) |
| 使用文档 | [caifugao110.github.io/3d-batch-copy](https://caifugao110.github.io/3d-batch-copy/) |
| 开源协议 | MIT |

---

## 功能特性

### 文件支持
- 支持 **STEP** 格式（`.step`、`.stp`）
- 可选包含 **XT** 格式（`.xt`、`.x_t`）
- 多源目录递归扫描，支持海量文件快速索引

### 复制模式
- 按清单文件自动匹配并批量复制
- 可选择是否将复制的文件统一重命名为清单中的名称
- 多线程并行复制，大幅提升处理速度

### 配置与管理
- 图形化配置管理，无需手动编辑配置文件
- 内置清单管理功能，方便编辑待处理文件列表
- 支持本地路径和局域网 UNC 路径（如 `\\192.168.160.2\生产管理部3d\...`）

### 结果与日志
- 自动生成 CSV 复制日志（GBK 编码，兼容 Excel）
- 完整的处理统计报告，含成功率、速度等
- 支持任务终止功能，可随时停止复制操作

### 界面
- 顶部分别显示标题、版本信息、主题切换、GitHub链接、使用说明、更新日志、关于
- 左侧面板：文件设置、复制选项、操作按钮（开始复制、停止处理、配置管理、清单管理等）
- 右侧面板：处理进度条、统计信息、实时日志显示
- 底部状态栏：当前运行状态

---

## 快速开始

### 环境要求

- Python >= 3.9
- Windows 操作系统

### 直接运行源码

```powershell
pip install -r requirements.txt
python .\app.py
```

---

## 构建

### 打包为单文件 exe

```powershell
.\scripts\build_exe.ps1
```

构建完成后保留产物：

```
dist\3d-batch-copy.exe
```

> 构建脚本会自动创建临时虚拟环境、安装依赖、调用 PyInstaller，并在结束后清理 `.venv`、`build`、spec 文件、缓存等过程文件。

---

## 配置说明 (`config.ini`)

```ini
[Paths]
target_dir_name = copystep
original_list_file = Original file list.txt
log_file = Copy step list.csv

[Settings]
max_workers = 12
retry_attempts = 3
rename_files = false
include_xt_format = false

[SourceDirectories]
source_1 = \\192.168.160.2\生产管理部3d\3D 资料\设计一课3D资料\03-SV GUN STEP
source_2 = \\192.168.160.2\生产管理部3d\3D 资料\吉利标准化\07吉利库STEP
```

| 配置项 | 说明 | 示例 |
| --- | --- | --- |
| `target_dir_name` | 本地目标目录名称 | `copystep` |
| `original_list_file` | 待处理文件清单 | `Original file list.txt` |
| `log_file` | 复制操作日志文件名 | `Copy step list.csv` |
| `max_workers` | 最大并发线程数 | `12`（建议 4~32） |
| `retry_attempts` | 复制失败重试次数 | `3` |
| `rename_files` | 是否按清单重命名文件 | `true` / `false` |
| `include_xt_format` | 是否包含 XT 格式文件 | `true` / `false` |
| `source_1`, `source_2`... | 源目录路径（完整 UNC 或本地路径） | `\\192.168.160.2\...` |

---

## 使用步骤

1. **准备清单文件** — 创建 `.csv` 或 `.txt` 文件，每行一个文件名（无需后缀）：
   ```txt
   SDEX-C0681L
   SDEX-C1036L(500-340)
   SDZX-C1195L
   SRTX-2C14693L
   ```
   > 支持自动清理后缀如 `-L`, `L(` 等。

2. **配置 `config.ini`** — 确保路径正确，特别是源目录路径

3. **启动程序** — 双击 `3d-batch-copy.exe`，程序自动加载配置和默认清单

4. **开始复制** — 点击"开始批量复制"，等待完成

5. **查看结果** — 文件出现在目标目录中，CSV 日志记录每个文件的来源与状态

---

## 日志与统计

### 日志格式（CSV，GBK 编码）

| 原始文件名 | 实际复制文件名 | 来源路径 | 重命名状态 |
| --- | --- | --- | --- |
| SRTX-2C14700L | SRTX-2C14700L.STEP | \\192.168.160.2\生产管理部3d\... | 是 |
| SRTX-2C14700AL | SRTX-2C14700A-L.STEP | \\192.168.160.2\生产管理部3d\... | 否 |

### 处理统计报告

```
============================================================
📊 处理统计报告
============================================================
  总文件数: 100
  ✅ 成功复制: 95 (95.0%)
  ❌ 未找到: 3 (3.0%)
  ⚠️ 复制错误: 2
⏱️ 总耗时: 8.2秒 | 平均速度: 12.2 文件/秒
  重命名模式: 启用
  包含 XT: 是
============================================================
```

> 若失败率 > 50%，会弹出警告提示。

---

## 常见问题 (FAQ)

### 配置文件不存在？
请确保 `config.ini` 与 `exe` 文件在同一目录。

### 无法访问网络路径？
检查网络路径是否正确，或手动映射网络驱动器。

### 复制速度慢？
尝试增加 `max_workers`（如 16 或 24），但不要超过 CPU 核心数。

### 文件名匹配失败？
程序会自动清理 `-L`, `L(` 等后缀，但仍需保证基础名称一致。

### 目标目录不清空？
程序在每次运行前会自动删除目标目录下所有旧的 3D 文件。

---

## 项目结构

```
3d-batch-copy/
├── app.py                  # GUI 主程序，包含全部核心逻辑
├── assets/
│   └── app.ico             # 应用图标
├── scripts/
│   └── build_exe.ps1       # Windows 构建脚本
├── .github/
│   └── workflows/
│       └── release.yml     # GitHub Actions 自动构建发布
├── .gitignore
├── LICENSE
├── pyproject.toml
├── README.md
├── requirements.txt
├── config.ini              # 用户配置文件
└── Original file list.txt  # 待处理文件清单
```

| 条目 | 说明 |
| --- | --- |
| `app.py` | GUI 主程序，包含全部核心逻辑 |
| `assets/` | 图标资源 |
| `scripts/` | 构建脚本 |
| `pyproject.toml` | 项目元数据与依赖配置 |
| `requirements.txt` | pip 依赖清单 |

---

## License

MIT © Tobin
