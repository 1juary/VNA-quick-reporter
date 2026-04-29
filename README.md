# VNA Data Automator 🚀

> VNA 测试报告自动生成工具 — 基于 RapidOCR + PySide6 + python-pptx

一键从 VNA 测试截图中提取频率-dB 数据，自动生成格式精美的 PPT 报告。

---

## 📦 项目结构

```
VNA-quick-reporter/
├── main.py                      # 应用程序入口
├── requirements.txt             # Python 依赖清单
├── vna_config.json              # 项目名-规格映射配置（自动生成）
├── vna_core/                    # 核心功能模块包
│   ├── __init__.py              # 包入口，统一导出所有模块
│   ├── ocr_extractor.py         # 模块1: OCR 图像处理与提取
│   ├── ppt_generator.py         # 模块2: PPT 报告生成
│   ├── file_utils.py            # 模块3: 配置文件读写工具
│   ├── ui_components.py         # 模块4: 自定义 UI 组件
│   ├── worker.py                # 模块5: OCR 后台工作线程
│   └── ui_main_window.py        # 模块6: 主窗口 UI 与业务流调度
```

---

## 🧩 模块说明

### 模块 1: OCR 图像处理与提取 (`ocr_extractor.py`)

**类**: `VNAOCRExtractor`

负责从 VNA 测试截图中提取频率和 dB 值。

- **多区域自适应裁剪**：优先裁剪右上角数据区，若识别不完整则自动尝试偏移区域和全图上半部分
- **RapidOCR 引擎**：基于 ONNX Runtime 的轻量 OCR，无需 GPU，首次运行自动下载模型
- **正则宽容解析**：容忍 GHz/dB 之间的任意干扰字符，就近匹配 1.5/3.0/4.5 GHz 三个目标频率
- **中文路径支持**：使用 `cv2.imdecode` 替代 `cv2.imread`，完美支持中文文件名

### 模块 2: PPT 报告生成 (`ppt_generator.py`)

**类**: `PPTGenerator`

基于 python-pptx 生成 16:9 宽屏 PPT 报告。

- **商务主题配色**：深蓝表头 (#4472C4) + 斑马纹交替行 (#E9EDF4 / #D9E1F2)
- **智能列宽分配**：文字列 1.3in，图片列 3.1in，完美容纳宽屏 VNA 截图
- **自动分页**：每页最多 7 行数据，超出自动分页并标注"续"
- **图片嵌入算法**：几何累加计算单元格绝对坐标，等比例缩放 + 居中留白
- **中英双语**：支持 English / 中文 两种报告语言
- **信息头**：支持项目名和规格参数显示

### 模块 3: 配置文件读写 (`file_utils.py`)

**函数**: `load_settings()`, `save_settings()`

- 以 JSON 格式持久化存储项目名与规格的映射关系
- 自动加载历史配置，支持快速选择

### 模块 4: 自定义 UI 组件 (`ui_components.py`)

**类**: `ImageCell` | **函数**: `auto_pair_files()`

- **拖拽互换**：按住图片可拖拽到其他单元格互换位置
- **一键删除**：红色 × 按钮快速清除图片
- **文件拖入**：支持从文件管理器拖入单张图片或整个文件夹
- **智能配对**：自动识别文件名中的 IL/RL 后缀进行配对

### 模块 5: OCR 后台工作线程 (`worker.py`)

**类**: `OCRWorker`

- 继承 `QThread`，在独立线程中执行 OCR 识别
- 通过 Signal-Slot 机制实时更新进度条和状态文本
- 避免大批量图片处理时 UI 假死

### 模块 6: 主窗口 UI (`ui_main_window.py`)

**类**: `SampleTab`, `MainWindow`

- **样品标签页**：每个样品独立标签，支持新增/删除点位行
- **自动填充**：加载图片时自动从文件名提取点位名称
- **批量拖入**：支持拖入整个文件夹，自动扫描并配对 IL/RL 图片
- **顶部信息栏**：报告语言选择、项目名输入、规格选择（带自动补全）
- **QSS 样式**：现代简约 UI 风格，圆角卡片、柔和阴影

---

## 🚀 快速开始

### 环境要求

- Python 3.8+
- Windows / macOS / Linux

### 安装依赖

```bash
pip install rapidocr-onnxruntime opencv-python pandas python-pptx PySide6 Pillow
```

### 运行

```bash
python main.py
```

首次运行会自动下载约 5MB 的 ONNX 模型文件。

---

## 📖 使用指南

### 基本流程

1. **选择语言**：在顶部下拉框选择 English 或 中文
2. **输入项目信息**：项目名和规格（可选，规格会自动记忆）
3. **添加图片**：
   - 点击空白单元格选择单张图片
   - 或直接拖入图片/文件夹到标签页
4. **预览数据**：点击"预览数据"查看 OCR 提取结果
5. **导出报告**：点击"导出完整报告"生成 PPT

### 图片命名规范

推荐命名格式（自动配对 IL/RL）：

```
点位1_IL.jpg    点位1_RL.jpg
Point2_IL.png   Point2_RL.png
```

也支持拖入整个文件夹，工具会自动扫描配对。

### 高级操作

- **互换图片**：按住图片拖拽到另一个单元格即可互换
- **删除图片**：点击图片右上角的红色 × 按钮
- **新增点位**：点击"+ 添加点位"按钮
- **新增样品**：点击标签栏右上角的"+ 新增样品"按钮

---

## 🛠️ 技术栈

| 组件 | 用途 |
|------|------|
| [RapidOCR](https://github.com/RapidAI/RapidOCR) | 基于 ONNX 的轻量 OCR 引擎 |
| [PySide6](https://wiki.qt.io/Qt_for_Python) | 跨平台桌面 GUI 框架 |
| [python-pptx](https://python-pptx.readthedocs.io/) | PowerPoint 文件生成 |
| [OpenCV](https://opencv.org/) | 图像裁剪与预处理 |
| [Pandas](https://pandas.pydata.org/) | 数据结构化与 DataFrame 操作 |

---

## 📄 许可证

MIT License
