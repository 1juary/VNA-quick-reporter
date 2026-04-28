# VNA Data Automator - 自动化报告生成工具

## 项目简介

这是一套生产级的 Python 桌面自动化解决方案，使用 **PaddleOCR** 从 VNA（矢量网络分析仪）测试图像中自动提取 IL（插入损耗）和 RL（回波损耗）数据，并生成专业的 PPT 测试报告。

## 核心特性

- **PaddleOCR 引擎**：对工业设备界面的屏显字体具有极强的识别能力，无需配置系统环境变量
- **多标签页交互**：支持多样品管理，可动态添加样品标签页
- **可点击单元格**：直接点击单元格即可放入图片，无需文件浏览器
- **预览核对功能**：导出前可预览 OCR 提取的数据，确保准确性
- **精准坐标计算**：几何累加算法实现表格单元格绝对坐标计算

## 环境准备

### 1. 安装 Python 依赖

```bash
pip install paddlepaddle paddleocr opencv-python pandas python-pptx PySide6
```

> **注意**：PaddleOCR 首次运行时会自动下载模型文件，请确保网络连接正常。

### 2. GPU 加速（可选）

如需 GPU 加速，请先安装 CUDA 和 cuDNN，然后执行：

```bash
pip install paddlepaddle-gpu
```

## 快速开始

### 运行 GUI 界面

```bash
python main.py
```

### 使用步骤

1. 在标签页中选择或新增样品
2. 点击表格单元格，在弹窗中选择对应的 IL/RL 测试图像
3. 点击"**预览PPT**"按钮，核对 OCR 提取的数据
4. 确认无误后，点击"**导出PPT**"生成报告

## UI 界面说明

- **浅蓝底色** (`#DDE6F5`)：主背景区域
- **绿色表头** (`#70A754`)：表格列标题
- **蓝色按钮** (`#5B8CD9`)：操作按钮
- **浅绿底色** (`#E2EADF`)：图片单元格

## OCR 识别配置

- **裁剪区域**：图像上半部分 + 右半部分（VNA 数据显示区域）
- **目标频率**：1.5 GHz, 3.0 GHz, 4.5 GHz
- **容差范围**：±0.2 GHz

## 项目结构

```
VNA-quick-reporter/
├── main.py              # 主程序入口
├── requirements.txt     # Python 依赖
├── README.md            # 项目说明
└── .gitignore           # Git忽略配置
```

## 技术栈

- **PaddleOCR**：光学字符识别（替代 Tesseract）
- **OpenCV**：图像预处理
- **Pandas**：数据结构化
- **python-pptx**：PPT 报告生成
- **PySide6**：跨平台 GUI 界面
- **Pillow**：图像尺寸计算

## 系统要求

- Windows 10/11 或 macOS/Linux
- Python 3.8+
- 4GB+ RAM（OCR 处理需要）

## 许可证

MIT License