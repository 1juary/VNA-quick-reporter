# VNA Quick Reporter - 自动化报告生成工具

## 项目简介

这是一套生产级的 Python 桌面自动化解决方案，用于从 VNA（矢量网络分析仪）测试图像中自动提取 IL（插入损耗）和 RL（回波损耗）数据，并生成专业的 PPT 测试报告。

## 核心特性

- **四大核心模块**：文件扫描配对、图像预处理 OCR 提取、PPT 报告生成、GUI 界面
- **鲁棒性设计**：图像放大、自适应二值化、智能背景检测
- **精准坐标计算**：几何累加算法实现表格单元格绝对坐标计算
- **多巴胺美学风格 UI**：PySide6 + QSS + QGraphicsDropShadowEffect 3D 悬浮效果

## 环境准备

### 1. 安装 Python 依赖

```bash
pip install -r requirements.txt
```

### 2. 安装 Tesseract-OCR 引擎（Windows）

1. 下载 Tesseract-OCR 安装包：https://github.com/UB-Mannheim/tesseract/wiki
2. 安装到默认路径 `C:\Program Files\Tesseract-OCR\tesseract.exe`
3. 或在代码中修改路径：
   ```python
   pytesseract.pytesseract.tesseract_cmd = r'你的安装路径\tesseract.exe'
   ```

## 快速开始

### 运行 GUI 界面

```bash
python main.py
```

### 使用步骤

1. 点击"📂 选择根目录"按钮，选择包含 VNA 测试图像的文件夹
2. 确保图像命名规范：`点位名_IL.jpg` 和 `点位名_RL.png` 格式
3. 点击"🚀 一键提取并生成 PPT"按钮开始处理
4. 完成后查看生成的 `VNA_Report_Auto.pptx` 文件

## 文件命名规范

```
点位1_IL.jpg    # IL 测试图像
点位1_RL.png    # RL 测试图像
点位2_IL.jpg    # 第二个点位
点位2_RL.png
...
```

支持的正则匹配：`^(.*?)_([iI][lL]|[rR][lL])\.(jpg|png|jpeg)$`

## OCR 识别配置

默认裁剪区域（右上角）：
- Y轴：10% - 30%
- X轴：75% - 100%

目标频率：1.5 GHz, 3.0 GHz, 4.5 GHz

## 项目结构

```
VNA-quick-reporter/
├── main.py              # 主程序入口
├── requirements.txt     # Python 依赖
├── README.md            # 项目说明
└── .gitignore           # Git忽略配置
```

## 技术栈

- **OpenCV**：图像预处理与形态学操作
- **Tesseract-OCR**：光学字符识别
- **Pandas**：数据结构化
- **python-pptx**：PPT 报告生成
- **PySide6**：跨平台 GUI 界面
- **Pillow**：图像尺寸计算

## 系统要求

- Windows 10/11（主要测试平台）
- Python 3.8+
- Tesseract-OCR 引擎

## 许可证

MIT License