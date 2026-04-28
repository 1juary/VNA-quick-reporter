import sys
import os
import re
import cv2
import pytesseract
import pandas as pd
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                               QPushButton, QLabel, QFileDialog, QTextEdit, QHBoxLayout,
                               QGraphicsDropShadowEffect, QFrame, QProgressBar)
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QFont, QColor

# ==========================================
# 模块 1: 文件扫描与配对模块
# ==========================================
class FileScanner:
    def __init__(self, root_dir):
        self.root_dir = Path(root_dir)
        
    def scan_and_pair(self, logger_callback=None):
        """扫描并根据名称特征提取点位，组装 IL 和 RL 的键值对"""
        files =[f for f in self.root_dir.iterdir() if f.suffix.lower() in ['.jpg', '.png', '.jpeg']]
        pairs = {}
        orphans =[]
        
        # 正则匹配示例: 点位1_IL.jpg, Point2_RL.png
        # 兼容性设计: 提取基础名称和测试类型(IL/RL)
        pattern = re.compile(r'^(.*?)_([iI][lL]|[rR][lL])\.(jpg|png|jpeg)$', re.IGNORECASE)
        
        for f in files:
            match = pattern.match(f.name)
            if match:
                base_name = match.group(1)
                test_type = match.group(2).upper()
                if base_name not in pairs:
                    pairs[base_name] = {}
                pairs[base_name][test_type] = str(f.absolute())
            else:
                orphans.append(f.name)
                
        # 防呆设计：检查落单文件
        valid_pairs = {}
        for base_name, paths in pairs.items():
            if 'IL' in paths and 'RL' in paths:
                valid_pairs[base_name] = paths
            else:
                orphans.append(base_name + " (配对缺失)")
                
        if logger_callback and orphans:
            logger_callback(f"[警告] 发现未配对或命名不规范的文件/点位: {', '.join(orphans)}")
            
        return valid_pairs

# ==========================================
# 模块 2: 图像预处理与 OCR 提取模块
# ==========================================
class OCRExtractor:
    def __init__(self):
        # 动态裁剪比例 (类属性，支持后续微调)
        self.crop_y_start = 0.10
        self.crop_y_end = 0.30
        self.crop_x_start = 0.75
        self.crop_x_end = 1.00
        
        # 目标频率，用于就近匹配
        self.target_freqs =[1.5, 3.0, 4.5]

    def process_image(self, img_path):
        """图像预处理与 OCR 识别"""
        img = cv2.imread(img_path)
        if img is None:
            return {}
            
        h, w = img.shape[:2]
        # 动态裁剪右上角
        crop = img[int(h * self.crop_y_start):int(h * self.crop_y_end), 
                   int(w * self.crop_x_start):int(w * self.crop_x_end)]
        
        # 预处理：放大图像提升 OCR 对小字符的敏感度
        crop = cv2.resize(crop, None, fx=2.5, fy=2.5, interpolation=cv2.INTER_CUBIC)
        gray = cv2.cvtColor(crop, cv2.COLOR_BGR2GRAY)
        
        # 自适应二值化 (消除背景网格线和波形干扰)
        # 使用 Otsu 算法找到最优阈值
        _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
        
        # 智能判定背景：如果是黑底白字，则进行反色处理 (Tesseract 对白底黑字识别率更高)
        if cv2.countNonZero(thresh) > (thresh.size / 2):
            thresh = cv2.bitwise_not(thresh)

        # 形态学操作：可选，去噪
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (2, 2))
        thresh = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, kernel, iterations=1)

        # 此处如未配置环境变量，请取消注释并修改为你本地的 Tesseract 路径
        # pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
        
        text = pytesseract.image_to_string(thresh, config='--psm 6')
        return self._parse_text(text)

    def _parse_text(self, text):
        """使用正则表达式清洗并结构化数据"""
        results = {1.5: "N/A", 3.0: "N/A", 4.5: "N/A"}
        
        # 匹配形如: 1.500 GHz   -0.47 dB 的文本行
        # 兼容性极强的正则：允许中间包含任意非数字干扰符
        pattern = re.compile(r'(\d+\.\d+)\s*[Gg][Hh][Zz][^\d-]*(-?\d+\.\d+)\s*[dD][Bb]')
        matches = pattern.findall(text)
        
        for match in matches:
            try:
                freq = float(match[0])
                db_val = f"{float(match[1]):.2f} dB"
                
                # 就近匹配到 1.5, 3.0, 4.5
                closest_freq = min(self.target_freqs, key=lambda x: abs(x - freq))
                # 误差小于0.1GHz才认为匹配成功
                if abs(closest_freq - freq) < 0.1:
                    results[closest_freq] = db_val
            except ValueError:
                continue
                
        return results

# ==========================================
# 模块 3: PPT 报告生成模块 (解决图层坐标难点)
# ==========================================
class PPTGenerator:
    def __init__(self, output_path):
        self.output_path = output_path
        self.prs = Presentation()
        # 设置 16:9 比例
        self.prs.slide_width = Inches(16)
        self.prs.slide_height = Inches(9)

    def generate(self, df):
        """基于 Pandas DataFrame 生成 PPT 报告"""
        blank_layout = self.prs.slide_layouts[6] 
        slide = self.prs.slides.add_slide(blank_layout)
        
        rows = len(df) + 1
        cols = 9
        
        # 创建表格对象
        left, top, width, height = Inches(0.5), Inches(0.5), Inches(15), Inches(8)
        table_shape = slide.shapes.add_table(rows, cols, left, top, width, height)
        table = table_shape.table
        
        # 设置表头
        headers =["Items", "1.500GHz (IL)", "3.000GHz (IL)", "4.500GHz (IL)", "Image (IL)", 
                   "1.500GHz (RL)", "3.000GHz (RL)", "4.500GHz (RL)", "Image (RL)"]
        for col_idx, header in enumerate(headers):
            cell = table.cell(0, col_idx)
            cell.text = header
            self._format_cell(cell, bold=True)
            
        # 填充数据与插入图片
        for row_idx, (index, row_data) in enumerate(df.iterrows(), start=1):
            table.cell(row_idx, 0).text = str(row_data['Item'])
            
            # IL 数据
            table.cell(row_idx, 1).text = row_data['1.5G_IL']
            table.cell(row_idx, 2).text = row_data['3.0G_IL']
            table.cell(row_idx, 3).text = row_data['4.5G_IL']
            self._insert_image_to_cell(slide, table_shape, row_idx, 4, row_data['Img_IL'])
            
            # RL 数据
            table.cell(row_idx, 5).text = row_data['1.5G_RL']
            table.cell(row_idx, 6).text = row_data['3.0G_RL']
            table.cell(row_idx, 7).text = row_data['4.5G_RL']
            self._insert_image_to_cell(slide, table_shape, row_idx, 8, row_data['Img_RL'])

            # 居中对齐所有文本
            for c in range(cols):
                self._format_cell(table.cell(row_idx, c))

        self.prs.save(self.output_path)

    def _format_cell(self, cell, bold=False):
        for paragraph in cell.text_frame.paragraphs:
            paragraph.font.size = Pt(12)
            paragraph.font.bold = bold

    def _insert_image_to_cell(self, slide, table_shape, row_idx, col_idx, img_path):
        """【核心算法】计算单元格绝对坐标并等比例插入图片"""
        if not img_path or not os.path.exists(img_path):
            return

        table = table_shape.table
        
        # 计算绝对 X 坐标
        cell_x = table_shape.left
        for i in range(col_idx):
            cell_x += table.columns[i].width
            
        # 计算绝对 Y 坐标
        cell_y = table_shape.top
        for j in range(row_idx):
            cell_y += table.rows[j].height
            
        cell_w = table.columns[col_idx].width
        cell_h = table.rows[row_idx].height

        # 读取原图尺寸以计算等比例缩放
        try:
            from PIL import Image
            with Image.open(img_path) as img:
                img_w, img_h = img.size
        except Exception:
            return

        # 留白策略：90%填充率，让图片在单元格内呈现出 Margin 效果
        ratio = min(cell_w / img_w, cell_h / img_h) * 0.85
        fit_w = int(img_w * ratio)
        fit_h = int(img_h * ratio)
        
        # 计算居中偏移量
        offset_x = cell_x + (cell_w - fit_w) / 2
        offset_y = cell_y + (cell_h - fit_h) / 2
        
        # 作为独立 Shape 添加到指定坐标
        slide.shapes.add_picture(img_path, offset_x, offset_y, fit_w, fit_h)

# ==========================================
# 模块 4: GUI 业务流线程调度
# ==========================================
class WorkerThread(QThread):
    log_signal = Signal(str)
    progress_signal = Signal(int)
    finished_signal = Signal(str)

    def __init__(self, target_dir):
        super().__init__()
        self.target_dir = target_dir

    def run(self):
        try:
            self.log_signal.emit("[1/4] 开始扫描文件夹并配对图像...")
            scanner = FileScanner(self.target_dir)
            pairs = scanner.scan_and_pair(logger_callback=self.log_signal.emit)
            
            if not pairs:
                self.log_signal.emit("未找到有效配对的测试数据，终止任务。")
                self.finished_signal.emit("Error")
                return

            self.log_signal.emit(f"成功配对 {len(pairs)} 个测试点位，开始 OCR 识别...")
            extractor = OCRExtractor()
            data_list =[]
            
            total = len(pairs)
            for idx, (item_name, paths) in enumerate(pairs.items(), 1):
                self.log_signal.emit(f"正在处理: {item_name} ...")
                
                il_data = extractor.process_image(paths['IL'])
                rl_data = extractor.process_image(paths['RL'])
                
                data_list.append({
                    'Item': item_name,
                    '1.5G_IL': il_data.get(1.5, "N/A"),
                    '3.0G_IL': il_data.get(3.0, "N/A"),
                    '4.5G_IL': il_data.get(4.5, "N/A"),
                    'Img_IL': paths['IL'],
                    '1.5G_RL': rl_data.get(1.5, "N/A"),
                    '3.0G_RL': rl_data.get(3.0, "N/A"),
                    '4.5G_RL': rl_data.get(4.5, "N/A"),
                    'Img_RL': paths['RL']
                })
                self.progress_signal.emit(int((idx / total) * 100))

            self.log_signal.emit("[3/4] 数据提取完成，开始基于 Pandas 结构化并生成 PPT...")
            df = pd.DataFrame(data_list)
            
            output_ppt = os.path.join(self.target_dir, "VNA_Report_Auto.pptx")
            ppt_gen = PPTGenerator(output_ppt)
            ppt_gen.generate(df)
            
            self.log_signal.emit(f"[4/4] 报告生成完毕！\n存储路径: {output_ppt}")
            self.finished_signal.emit("Success")
            
        except Exception as e:
            self.log_signal.emit(f"[发生异常]: {str(e)}")
            self.finished_signal.emit("Error")

# ==========================================
# 模块 5: Dopamine 多巴胺美学风格 UI (PySide6)
# ==========================================
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("VNA Automator Pro")
        self.resize(850, 600)
        self.target_dir = ""
        self.init_ui()
        self.apply_dopamine_style()

    def init_ui(self):
        # 主面板（卡片容器）
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)
        main_layout.setContentsMargins(30, 30, 30, 30)
        main_layout.setSpacing(20)

        # 顶部卡片 (控制区)
        control_card = QFrame()
        control_card.setObjectName("ControlCard")
        self.add_shadow(control_card)
        
        control_layout = QVBoxLayout(control_card)
        control_layout.setContentsMargins(20, 20, 20, 20)
        
        # 路径选择器
        path_layout = QHBoxLayout()
        self.lbl_path = QLabel("尚未选择目标文件夹")
        self.lbl_path.setObjectName("PathLabel")
        self.btn_select = QPushButton("📂 选择根目录")
        self.btn_select.setObjectName("BtnSelect")
        self.btn_select.setCursor(Qt.PointingHandCursor)
        self.btn_select.clicked.connect(self.select_folder)
        
        path_layout.addWidget(self.lbl_path, stretch=1)
        path_layout.addWidget(self.btn_select)
        
        # 一键生成按钮
        self.btn_generate = QPushButton("🚀 一键提取并生成 PPT")
        self.btn_generate.setObjectName("BtnGenerate")
        self.btn_generate.setCursor(Qt.PointingHandCursor)
        self.btn_generate.setEnabled(False)  # 防误触锁定
        self.btn_generate.clicked.connect(self.start_processing)

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setFixedHeight(8)
        self.progress_bar.setObjectName("ProgressBar")

        control_layout.addLayout(path_layout)
        control_layout.addSpacing(15)
        control_layout.addWidget(self.btn_generate)
        control_layout.addSpacing(10)
        control_layout.addWidget(self.progress_bar)

        # 底部卡片 (日志滚动区)
        log_card = QFrame()
        log_card.setObjectName("LogCard")
        self.add_shadow(log_card)
        
        log_layout = QVBoxLayout(log_card)
        log_layout.setContentsMargins(0, 0, 0, 0)
        
        self.text_log = QTextEdit()
        self.text_log.setObjectName("LogText")
        self.text_log.setReadOnly(True)
        log_layout.addWidget(self.text_log)

        main_layout.addWidget(control_card, stretch=2)
        main_layout.addWidget(log_card, stretch=5)

    def add_shadow(self, widget):
        """给 UI 元素添加 3D 悬浮投影效果"""
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(20)
        shadow.setColor(QColor(0, 0, 0, 30))
        shadow.setOffset(0, 6)
        widget.setGraphicsEffect(shadow)

    def apply_dopamine_style(self):
        """注入多巴胺美学 QSS (明亮、高饱和、现代微投影、圆角过渡)"""
        self.setStyleSheet("""
            QMainWindow {
                background-color: #F4F7FE; /* 极简明亮的灰蓝底色 */
            }
            #ControlCard, #LogCard {
                background-color: #FFFFFF;
                border-radius: 16px;
            }
            #PathLabel {
                font-size: 14px;
                color: #A3AED0;
                font-weight: bold;
                background-color: #F4F7FE;
                padding: 10px;
                border-radius: 8px;
            }
            #BtnSelect {
                background-color: #4318FF; /* 活力亮蓝 */
                color: white;
                font-size: 14px;
                font-weight: bold;
                padding: 10px 20px;
                border-radius: 10px;
                border: none;
            }
            #BtnSelect:hover { background-color: #3311DB; }
            
            #BtnGenerate {
                background-color: #FF5E3A; /* 活力多巴胺橙 */
                color: white;
                font-size: 16px;
                font-weight: bold;
                padding: 15px;
                border-radius: 12px;
                border: none;
            }
            #BtnGenerate:hover { background-color: #E04D2C; }
            #BtnGenerate:disabled { background-color: #FFE2DB; color: #FFAA99; }
            
            #ProgressBar {
                background-color: #E9EDF7;
                border-radius: 4px;
            }
            #ProgressBar::chunk {
                background-color: #00D563; /* 活力亮绿 */
                border-radius: 4px;
            }
            
            #LogText {
                background-color: transparent;
                border: none;
                padding: 15px;
                font-family: Consolas, "Microsoft YaHei";
                font-size: 13px;
                color: #2B3674;
            }
            QScrollBar:vertical {
                width: 8px;
                background: transparent;
            }
            QScrollBar::handle:vertical {
                background: #E9EDF7;
                border-radius: 4px;
            }
        """)

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "选择 VNA 图像根目录")
        if folder:
            self.target_dir = folder
            self.lbl_path.setText(f"当前路径: {folder}")
            self.btn_generate.setEnabled(True)
            self.text_log.clear()
            self.append_log("📂 已加载目录，等待执行任务...")

    def append_log(self, text):
        self.text_log.append(text)
        # 自动滚动到最底部
        scrollbar = self.text_log.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def start_processing(self):
        self.btn_generate.setEnabled(False)
        self.btn_select.setEnabled(False)
        self.progress_bar.setValue(0)
        self.text_log.clear()
        
        # 启动后台处理线程，避免主 UI 假死
        self.thread = WorkerThread(self.target_dir)
        self.thread.log_signal.connect(self.append_log)
        self.thread.progress_signal.connect(self.progress_bar.setValue)
        self.thread.finished_signal.connect(self.task_finished)
        self.thread.start()

    def task_finished(self, status):
        self.btn_select.setEnabled(True)
        self.btn_generate.setEnabled(True)
        if status == "Success":
            self.append_log("✨ 所有任务圆满结束！")
        else:
            self.append_log("❌ 任务以异常状态终止。")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion") # 跨平台 UI 一致性保障
    window = MainWindow()
    window.show()
    sys.exit(app.exec())