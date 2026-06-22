# 📄 3D文件批量复制工具 - 使用说明手册

> **软件名称**：3D文件批量复制工具  
> **作者**：Tobin © 2026  
> **项目地址**：[GitHub 项目主页](https://github.com/caifugao110/3d-batch-copy)  
> **使用文档地址**：[在线帮助文档](https://caifugao110.github.io/3d-batch-copy/)  

---

## 🎯 软件功能简介

本工具专为OBARA有调图需求的人员打造，用于从多个网络或本地目录中**快速查找并批量复制3D模型文件**到指定目标目录。

✨ **核心功能**：
- ✅ 多源目录递归扫描，支持海量文件快速索引
- ✅ 支持按清单文件自动匹配并复制
- ✅ 支持多种3D文件格式：STEP (.step, .stp) 和 XT (.xt, .x_t)
- 🔁 可选择是否将复制的文件**统一重命名为清单中的名称**
- 🚀 多线程并行复制，大幅提升处理速度
- 📝 自动生成复制日志，记录来源与结果
- ⚙️ 图形化配置管理，无需手动编辑配置
- 📝 内置清单管理功能，方便编辑待处理文件列表
- ⏹️ 支持任务终止功能，可随时停止正在进行的复制操作
- 🌓 深色/浅色主题切换，保护眼睛
- 🔄 自动更新检查，保持软件最新版本

---

## 📦 软件运行环境

| 项目       | 要求                                                         |
| ---------- | ------------------------------------------------------------ |
| 🖥️ 操作系统 | Windows 7 / 10 / 11（64位推荐）                              |
| 📂 依赖     | 无需安装Python，已打包为独立 `exe` 程序                     |
| 💾 磁盘空间 | 至少 100MB 可用空间（用于日志和缓存）                        |
| 🔗 网络     | 若使用网络驱动器（如 `\\192.168.160.2\...`），请确保已正确映射盘符 |

---

## 🚀 快速开始

### 1️⃣ 启动程序
双击 `3D文件批量复制工具.exe` 即可启动，程序会**自动尝试加载根目录下的 `config.ini` 配置文件**。

> 📌 提示：首次使用请先配置 `config.ini` 文件。

---

## ⚙️ 配置文件详解 (`config.ini`)

配置文件是程序运行的核心，必须存在且格式正确。

### 📁 文件结构示例
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
source_3 = \\192.168.160.2\生产管理部3d\3D 资料\上海3D图库拷贝文件\03-SV GUN STEP
source_4 = \\192.168.160.2\生产管理部3d\3D 资料\上海3D图库拷贝文件\吉利标准化\07吉利库STEP
source_5 = \\192.168.160.2\生产管理部3d\3D 资料\设计一课3D资料\01-SV GUN ASSY\13-PSA\00-STP
source_6 = \\192.168.160.2\生产管理部3d\3D 资料\设计一课3D资料\01-SV GUN ASSY\16-恒大-X2CV2-TOL\STP
source_7 = \\192.168.160.2\生产管理部3d\3D 资料\设计一课3D资料\01-SV GUN ASSY\17-铝焊钳\03-STP
source_8 = \\192.168.160.2\生产管理部3d\3D 资料\设计一课3D资料\01-SV GUN ASSY\18-蔚来-X2CV2-TOL\STEP
source_9 = \\192.168.160.2\生产管理部3d\3D 资料\设计一课3D资料\01-SV GUN ASSY\21-福建奔驰\STP
source_10 = \\192.168.160.2\生产管理部3d\3D 资料\设计一课3D资料\01-SV GUN ASSY\22-印度\STP
source_11 = \\192.168.160.2\生产管理部3d\3D 资料\设计一课3D资料\01-SV GUN ASSY\25-理想-X2CV2-TOL\STP
source_12 = \\192.168.160.2\生产管理部3d\3D 资料\设计一课3D资料\01-SV GUN ASSY\26-海斯坦普\01-STP
source_13 = \\192.168.160.2\生产管理部3d\3D 资料\设计一课3D资料\01-SV GUN ASSY\30-比亚迪\01-STP
source_14 = \\192.168.160.2\生产管理部3d\3D 资料\设计一课3D资料\X2C V2标准化数据\000-STEP
source_15 = \\192.168.160.2\生产管理部3d\3D 资料\设计二课3D资料\0000 3D\STEP
```

### 🔍 配置项说明

| 配置项                    | 说明                                | 示例                                                         |
| ------------------------- | ----------------------------------- | ------------------------------------------------------------ |
| `target_dir_name`         | 本地目标目录名称                    | `copystep` → 在程序同目录创建 `copystep` 文件夹              |
| `original_list_file`      | 待处理文件清单                      | `Original file list.txt`                                     |
| `log_file`                | 复制操作日志文件名                  | `Copy step list.csv`                                         |
| `max_workers`             | 最大并发线程数（影响速度）          | `12`（建议4~32）                                             |
| `retry_attempts`          | 复制失败重试次数                    | `3`                                                          |
| `rename_files`            | 是否按清单重命名文件                | `true` / `false`                                             |
| `include_xt_format`       | 是否包含XT格式文件                  | `true` / `false`                                             |
| `source_1`, `source_2`... | 源目录路径（完整路径）              | `\\192.168.160.2\生产管理部3d\3D 资料\设计一课3D资料\03-SV GUN STEP` |

> 📂 **注意**：源目录路径现在支持完整路径，不再需要单独设置 `drive_letter`。

---

## 🖼️ 软件界面说明

### 🧩 主界面布局

界面主要分为：
- **顶部标题栏**：软件名称、版本信息、主题切换、GitHub链接、使用说明、检查更新
- **左侧面板**：文件设置（配置文件、清单文件选择）、操作按钮（开始复制、停止处理、配置管理、清单管理、打开目标目录、查看复制日志、清空日志框）
- **右侧面板**：处理进度条、统计信息、实时日志显示

---

## 🛠️ 使用步骤

### ✅ 步骤 1：准备清单文件

创建一个 `.csv` 或 `.txt` 文件，每行一个文件名（无需后缀）：

```txt
# Original file list.txt
SDEX-C0681L
SDEX-C1036L(500-340)
SDZX-C1195L
SRTC-2C4809L
SRTX-2C14693L
SRTX-2C14694L
SRTX-2C14695L
SRTX-2C14697L
SRTX-2C14698L
SRTX-2C14699L
SRTX-2C14700AL
SRTX-2C14700L
```

> 📝 支持自动清理后缀如 `-L`, `L(1)` 等。

### ✅ 步骤 2：配置 `config.ini`

确保 `config.ini` 中的路径正确，特别是源目录路径。

### ✅ 步骤 3：启动程序

1. 双击 `3D文件批量复制工具.exe` 文件
2. 程序自动加载 `config.ini` 和默认清单
3. 选择是否启用"按照清单重命名3D文件"和"包含 XT 格式3D文件"
4. 点击 **🚀 开始批量复制**

### ✅ 步骤 4：查看结果

- 成功复制的文件会出现在目标目录中
- 日志文件记录每个文件的：
  - 原始文件名
  - 实际复制文件名
  - 来源路径
  - 是否重命名

---

## ⚙️ 高级功能

### 🔧 配置管理窗口（⚙️ 按钮）

点击 **⚙️ 配置管理** 可图形化修改所有设置：

- ✏️ 修改目标目录名
- ✏️ 修改原始清单文件名
- ✏️ 修改复制日志文件名
- 📂 添加/删除源目录
- 🔄 调整线程数与重试次数
- 🔁 启用/禁用"按清单重命名"
- 🔁 启用/禁用"包含 XT 格式3D文件"
- 💾 自动保存配置

> ✅ 所有修改即时生效，无需重启。

### 📝 清单管理功能（📝 按钮）

点击 **📝 清单管理** 可直接编辑待处理文件列表：

- ✏️ 直接在界面中添加、删除、修改文件列表
- 💾 保存后自动重新加载清单

### 🔄 重命名模式说明

| 模式                              | 说明                                                      |
| --------------------------------- | --------------------------------------------------------- |
| **启用** (`rename_files = true`)  | 复制后文件统一命名为清单中的名称，如 `SRTX-2C14700L.STEP` |
| **禁用** (`rename_files = false`) | 保留源文件名，如 `SRTX-2C14700-L.STEP`                    |

> 💡 建议在需要统一命名规范时启用。

### 🔄 自动更新功能

程序支持自动更新检查：
- **启动时自动检查**：每次启动时在后台检查新版本
- **手动检查更新**：点击"检查更新"按钮立即检查
- **一键更新**：发现新版本时可自动下载并安装

### 🧪 高级用法（命令行模式）

> 如果你使用的是源码版本（非exe），可通过命令行运行：

```
python app.py
```

> `3D文件批量复制工具.py` 仍保留为兼容入口，也可以继续使用原命令启动。

> 如果你想要将源码打包为exe文件，可通过命令行运行：

```
powershell -ExecutionPolicy Bypass -File scripts\build_exe.ps1
```

---

## 📊 日志与调试

### 📄 日志文件格式

| 原始文件名     | 实际复制文件名       | 来源路径                         | 重命名状态 |
| -------------- | -------------------- | -------------------------------- | ---------- |
| SRTX-2C14700L  | SRTX-2C14700L.STEP   | \\192.168.160.2\生产管理部3d\3D 资料\设计一课3D资料\03-SV GUN STEP | 是         |
| SRTX-2C14700AL | SRTX-2C14700A-L.STEP | \\192.168.160.2\生产管理部3d\3D 资料\设计一课3D资料\03-SV GUN STEP | 否         |

> 📂 编码为 `GBK`，兼容 Excel 打开不乱码。

### 📋 处理统计报告

复制完成后，日志中会输出：

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

> ⚠️ 若失败率 > 50%，会弹出警告提示。

---

## ❓ 常见问题 (FAQ)

### ❌ 配置文件不存在？
> 请确保 `config.ini` 与 `exe` 文件在同一目录。

### 🔒 无法访问网络路径？
> 检查网络路径是否正确，或手动映射网络驱动器。

### 📉 复制速度慢？
> 尝试增加 `max_workers`（如 16 或 24），但不要超过CPU核心数。

### 🧹 目标目录不清空？
> 程序会在每次运行前**自动删除目标目录下的所有旧3D文件**。

### 📎 文件名匹配失败？
> 程序会自动清理 `-L`, `L(` 等后缀，但仍需保证基础名称一致。

### 🔄 更新失败？
> 确保网络连接正常，或从项目主页手动下载最新版本。

---

## 📎 技术支持

- 🐞 **报告Bug**：[GitHub Issues](https://github.com/caifugao110/3d-batch-copy/issues)
- 💬 **功能建议**：欢迎提交 PR 或 Issue
- 📧 **联系作者**：caifugao110@gmail.com

---

## 📄 许可证

本项目采用 [MIT 许可证](LICENSE) 开源协议。

```
MIT License

Copyright (c) 2026 Tobin

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

这意味着您可以自由地：
- ✅ 使用、复制、修改、合并、出版发行、散布、再授权及贩售软件副本
- ✅ 用于商业用途
- ✅ 修改源代码

唯一的要求是在所有副本或实质性部分中包含版权声明和许可声明。
