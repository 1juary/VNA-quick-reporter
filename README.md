# VNA Data Automator - RapidOCR & PPTX Edition

## 项目简介

这是一套生产级的 Python 桌面自动化解决方案，使用 **RapidOCR** 从 VNA（矢量网络分析仪）测试图像中自动提取 IL（插入损耗）和 RL（回波损耗）数据，并生成专业的 PPT 测试报告。支持多样品管理、智能图片配对和拖拽操作。

## 核心特性

- **RapidOCR 引擎**：基于 ONNX Runtime 的高速 OCR 引擎，支持中文路径，识别准确率高
- **智能拖拽配对**：支持图片拖拽导入，自动识别 IL/RL 图像并智能配对
- **多标签页交互**：支持多样品管理，可动态添加样品标签页
- **单元格互换**：按住图片可与其他单元格交换位置
- **预览核对功能**：导出前可预览 OCR 提取的数据，确保准确性
- **专业 PPT 生成**：自动生成包含数据表格和图片的 PowerPoint 报告

## 环境准备

### 1. 安装 Python 依赖

```bash
pip install rapidocr-onnxruntime opencv-python pandas python-pptx PySide6 pillow numpy
```

> **注意**：首次运行时会自动下载 OCR 模型文件，请确保网络连接正常。

### 2. 创建虚拟环境（推荐）

```bash
python -m venv .venv
# Windows
.venv\Scripts\activate
# Linux/Mac
source .venv/bin/activate

pip install -r requirements.txt
```

## 快速开始

### 运行 GUI 界面

```bash
python main.py
```

### 使用步骤

1. 输入项目名和规格（可选，会自动保存配置）
2. 在标签页中选择或新增样品
3. 通过以下方式添加图片：
   - 点击单元格选择文件
   - 拖拽单张图片到单元格
   - 拖拽多张图片或文件夹到界面（自动智能配对）
4. 点击"**预览数据**"按钮，核对 OCR 提取的数据
5. 确认无误后，点击"**🚀 导出完整报告**"生成 PPT

## UI 界面说明

- **项目信息卡片**：输入项目名和规格，支持历史记录
- **标签页管理**：每个标签页代表一个样品，支持动态添加
- **图片单元格**：支持点击选择、拖拽导入和互换操作
- **智能配对**：拖拽文件夹时自动识别文件名并配对 IL/RL 图像

## OCR 识别配置

- **裁剪区域**：图像左上部分数据区域（10%-28%高度，65%-88%宽度）
- **目标频率**：1.5 GHz, 3.0 GHz, 4.5 GHz
- **容差范围**：±0.2 GHz
- **支持格式**：PNG, JPG, JPEG

## 项目结构

```
VNA-quick-reporter/
├── main.py              # 主程序入口和 GUI 界面
├── convert_icon.py      # 图标转换脚本
├── main.spec            # PyInstaller 打包配置
├── requirements.txt     # Python 依赖列表
├── vna_config.json      # 用户配置存储
├── README.md            # 项目说明文档
└── build/               # PyInstaller 构建输出
    ├── main/            # 构建中间文件
    └── ...              # 其他构建文件
```

## 技术栈

- **RapidOCR**：基于 ONNX Runtime 的高速光学字符识别
- **OpenCV**：图像预处理和裁剪
- **Pandas**：数据结构化和处理
- **python-pptx**：PowerPoint 报告生成
- **PySide6**：跨平台 GUI 框架
- **Pillow**：图像尺寸计算和处理
- **NumPy**：数值计算和图像处理

## 系统要求

- Windows 10/11, macOS, 或 Linux
- Python 3.8+
- 4GB+ RAM（OCR 处理需要）
- 支持拖拽操作的桌面环境

## 许可证

MIT License