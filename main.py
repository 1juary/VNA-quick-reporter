import sys
import os
import re
import cv2
import pandas as pd
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN

# 【引入 RapidOCR 替代 PaddleOCR】
from rapidocr_onnxruntime import RapidOCR

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QPushButton, QLabel, QFileDialog, QTabWidget, QTableWidget, 
    QHeaderView, QProgressDialog, QMessageBox, QDialog, QTextEdit
)
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QPixmap

# ==========================================
# 模块 1: RapidOCR 图像处理与提取模块 (轻量、极速、无BUG)
# ==========================================
class VNAOCRExtractor:
    def __init__(self):
        # 初始化 RapidOCR
        self.ocr = RapidOCR()
        self.target_freqs =[1.5, 3.0, 4.5]

    def process_image(self, img_path):
        if not img_path or not os.path.exists(img_path):
            return {1.5: "N/A", 3.0: "N/A", 4.5: "N/A"}

        img = cv2.imread(img_path)
        if img is None:
            return {1.5: "N/A", 3.0: "N/A", 4.5: "N/A"}

        h, w = img.shape[:2]
        
        # 精准裁剪右上角数据区
        crop = img[int(h * 0.10):int(h * 0.28), int(w * 0.65):int(w * 0.88)]

        # 调用 RapidOCR 提取文本
        result, _ = self.ocr(crop)
        
        text_content = ""
        # RapidOCR 的结果格式为: [[boxes, 'text1', score], [boxes, 'text2', score]]
        if result:
            for line in result:
                text_content += str(line[1]) + " "

        return self._parse_text(text_content)

    def _parse_text(self, text):
        results = {1.5: "N/A", 3.0: "N/A", 4.5: "N/A"}
        
        pattern = re.compile(r'(\d+\.\d+)\s*[Gg][Hh][Zz][^\d-]{0,10}(-?\d+\.\d+)\s*[dD][Bb]')
        matches = pattern.findall(text)
        
        for match in matches:
            try:
                freq = float(match[0])
                db_val = f"{float(match[1]):.2f}dB" 
                
                closest_freq = min(self.target_freqs, key=lambda x: abs(x - freq))
                if abs(closest_freq - freq) < 0.2:
                    if results[closest_freq] == "N/A":
                        results[closest_freq] = db_val
            except ValueError:
                continue
                
        return results

# ==========================================
# 模块 2: PPT 报告生成模块
# ==========================================
class PPTGenerator:
    def __init__(self, output_path):
        self.output_path = output_path
        self.prs = Presentation()
        self.prs.slide_width = Inches(16)
        self.prs.slide_height = Inches(9)

    def generate(self, dataset):
        blank_layout = self.prs.slide_layouts[6] 

        for sample_name, df in dataset.items():
            if df.empty:
                continue
            slide = self.prs.slides.add_slide(blank_layout)
            
            title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(15), Inches(0.8))
            p = title_box.text_frame.paragraphs[0]
            p.text = sample_name
            p.font.size = Pt(24)
            p.font.bold = True

            rows = len(df) + 2 
            cols = 9
            
            left, top, width, height = Inches(0.5), Inches(1.0), Inches(15), Inches(7.5)
            table_shape = slide.shapes.add_table(rows, cols, left, top, width, height)
            table = table_shape.table
            
            table.cell(0, 0).merge(table.cell(1, 0))
            table.cell(0, 1).merge(table.cell(0, 4))
            table.cell(0, 5).merge(table.cell(0, 8))

            table.cell(0, 1).text = "IL"
            table.cell(0, 5).text = "RL"
            
            headers =["Items", "1.500GHz", "3.000GHz", "4.500GHz", "Image", 
                       "1.500GHz", "3.000GHz", "4.500GHz", "Image"]
            for col_idx, header in enumerate(headers):
                target_cell = table.cell(0, 0) if col_idx == 0 else table.cell(1, col_idx)
                target_cell.text = header

            for r in range(2):
                for c in range(cols):
                    for paragraph in table.cell(r, c).text_frame.paragraphs:
                        paragraph.alignment = PP_ALIGN.CENTER
                        paragraph.font.bold = True
                
            for idx, (index, row_data) in enumerate(df.iterrows()):
                row_idx = idx + 2
                table.cell(row_idx, 0).text = f"点位 {idx + 1}"
                
                table.cell(row_idx, 1).text = row_data['1.5G_IL']
                table.cell(row_idx, 2).text = row_data['3.0G_IL']
                table.cell(row_idx, 3).text = row_data['4.5G_IL']
                self._insert_image_to_cell(slide, table_shape, row_idx, 4, row_data['Img_IL'])
                
                table.cell(row_idx, 5).text = row_data['1.5G_RL']
                table.cell(row_idx, 6).text = row_data['3.0G_RL']
                table.cell(row_idx, 7).text = row_data['4.5G_RL']
                self._insert_image_to_cell(slide, table_shape, row_idx, 8, row_data['Img_RL'])

                for c in range(cols):
                    for paragraph in table.cell(row_idx, c).text_frame.paragraphs:
                        paragraph.alignment = PP_ALIGN.CENTER

        self.prs.save(self.output_path)

    def _insert_image_to_cell(self, slide, table_shape, row_idx, col_idx, img_path):
        if not img_path or not os.path.exists(img_path):
            return

        table = table_shape.table
        cell_x = table_shape.left + sum(table.columns[i].width for i in range(col_idx))
        cell_y = table_shape.top + sum(table.rows[j].height for j in range(row_idx))
        cell_w = table.columns[col_idx].width
        cell_h = table.rows[row_idx].height

        try:
            from PIL import Image
            with Image.open(img_path) as img:
                img_w, img_h = img.size
        except Exception:
            return

        ratio = min(cell_w / img_w * 0.95, cell_h / img_h * 0.85)
        fit_w = int(img_w * ratio)
        fit_h = int(img_h * ratio)
        
        offset_x = cell_x + (cell_w - fit_w) / 2
        offset_y = cell_y + (cell_h - fit_h) / 2
        
        slide.shapes.add_picture(img_path, offset_x, offset_y, fit_w, fit_h)

# ==========================================
# 模块 3: 自定义 UI 组件
# ==========================================
class ImageCell(QLabel):
    def __init__(self):
        super().__init__()
        self.image_path = ""
        self.setText("点击放入图片")
        self.setAlignment(Qt.AlignCenter)
        self.setCursor(Qt.PointingHandCursor) 
        self.setStyleSheet("""
            QLabel {
                background-color: #E2EADF;
                color: #555555;
                font-size: 14px;
                border: 1px dashed #A5A5A5;
            }
            QLabel:hover {
                background-color: #D3DDCF;
            }
        """)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            path, _ = QFileDialog.getOpenFileName(self, "选择图片", "", "Images (*.png *.jpg *.jpeg)")
            if path:
                self.image_path = path
                pixmap = QPixmap(path)
                self.setPixmap(pixmap.scaled(self.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation))
                self.setStyleSheet("border: none; background-color: white;")

# ==========================================
# 模块 4: GUI 业务流线程与界面
# ==========================================
class OCRWorker(QThread):
    progress_update = Signal(int, str)
    finished = Signal(dict)

    def __init__(self, ui_data):
        super().__init__()
        self.ui_data = ui_data 
        
    def run(self):
        extractor = VNAOCRExtractor()
        result_dataset = {}
        total_tasks = sum(len(pairs) for pairs in self.ui_data.values())
        current_task = 0

        for sample_name, pairs in self.ui_data.items():
            sample_data =[]
            for il_path, rl_path in pairs:
                current_task += 1
                self.progress_update.emit(int(current_task / total_tasks * 100), f"正在处理: {sample_name}...")
                
                il_data = extractor.process_image(il_path)
                rl_data = extractor.process_image(rl_path)
                
                sample_data.append({
                    '1.5G_IL': il_data.get(1.5),
                    '3.0G_IL': il_data.get(3.0),
                    '4.5G_IL': il_data.get(4.5),
                    'Img_IL': il_path,
                    '1.5G_RL': rl_data.get(1.5),
                    '3.0G_RL': rl_data.get(3.0),
                    '4.5G_RL': rl_data.get(4.5),
                    'Img_RL': rl_path
                })
            result_dataset[sample_name] = pd.DataFrame(sample_data)

        self.finished.emit(result_dataset)

class SampleTab(QWidget):
    def __init__(self):
        super().__init__()
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(20, 20, 20, 20)
        
        self.table = QTableWidget(3, 2)
        self.table.setHorizontalHeaderLabels(["IL IMAGE", "RL image"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.verticalHeader().setDefaultSectionSize(120)
        self.table.setSelectionMode(QTableWidget.NoSelection)
        self.table.setFocusPolicy(Qt.NoFocus)

        for row in range(3):
            self.table.setCellWidget(row, 0, ImageCell())
            self.table.setCellWidget(row, 1, ImageCell())
            
        self.layout.addWidget(self.table)
        
        self.btn_add_row = QPushButton("+")
        self.btn_add_row.setObjectName("BtnAddRow")
        self.btn_add_row.clicked.connect(self.add_row)
        self.layout.addWidget(self.btn_add_row)

    def add_row(self):
        row = self.table.rowCount()
        self.table.insertRow(row)
        self.table.setCellWidget(row, 0, ImageCell())
        self.table.setCellWidget(row, 1, ImageCell())

    def get_image_pairs(self):
        pairs =[]
        for row in range(self.table.rowCount()):
            il_cell = self.table.cellWidget(row, 0)
            rl_cell = self.table.cellWidget(row, 1)
            if il_cell.image_path and rl_cell.image_path:
                pairs.append((il_cell.image_path, rl_cell.image_path))
        return pairs

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("VNA Data Automator (RapidOCR 稳健版)")
        self.resize(950, 750)
        self.init_ui()
        self.apply_styles()

    def init_ui(self):
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)
        
        self.tabs = QTabWidget()
        self.tabs.addTab(SampleTab(), "样品1")
        self.tabs.addTab(SampleTab(), "样品2")
        self.tabs.addTab(SampleTab(), "样品3")
        
        self.btn_add_tab = QPushButton("+ 新增样品")
        self.btn_add_tab.clicked.connect(lambda: self.tabs.addTab(SampleTab(), f"样品{self.tabs.count()+1}"))
        self.tabs.setCornerWidget(self.btn_add_tab, Qt.TopRightCorner)

        main_layout.addWidget(self.tabs)

        bottom_layout = QHBoxLayout()
        bottom_layout.addStretch()
        
        self.btn_preview = QPushButton("预览数据")
        self.btn_preview.setObjectName("BtnAction")
        self.btn_preview.clicked.connect(self.preview_data)
        
        self.btn_export = QPushButton("导出PPT")
        self.btn_export.setObjectName("BtnAction")
        self.btn_export.clicked.connect(self.export_ppt)
        
        bottom_layout.addWidget(self.btn_preview)
        bottom_layout.addWidget(self.btn_export)
        bottom_layout.addStretch()
        
        main_layout.addLayout(bottom_layout)

    def apply_styles(self):
        self.setStyleSheet("""
            QMainWindow { background-color: #FFFFFF; }
            QTabWidget::pane {
                background-color: #DDE6F5; 
                border-radius: 10px; padding: 10px;
            }
            QTabBar::tab {
                background-color: #5B8CD9; color: white;
                padding: 10px 20px; margin-right: 5px;
                border-top-left-radius: 8px; border-top-right-radius: 8px;
                font-weight: bold; font-size: 14px;
            }
            QTabBar::tab:!selected { background-color: #9DBBEE; }
            QTableWidget {
                background-color: #E2EADF; border: 1px solid #C0C0C0; gridline-color: white;
            }
            QHeaderView::section {
                background-color: #70A754; color: white; font-weight: bold; font-size: 16px;
                border: 1px solid white; padding: 8px;
            }
            #BtnAddRow {
                background-color: transparent; font-size: 24px; color: #555555; border: none;
            }
            #BtnAddRow:hover { color: #000000; }
            #BtnAction {
                background-color: #5B8CD9; color: white; border-radius: 6px;
                padding: 12px 30px; font-size: 16px; font-weight: bold; border: 1px solid #4A75B8;
            }
            #BtnAction:hover { background-color: #4A75B8; }
        """)

    def gather_ui_data(self):
        ui_data = {}
        for i in range(self.tabs.count()):
            tab_name = self.tabs.tabText(i)
            tab_widget = self.tabs.widget(i)
            pairs = tab_widget.get_image_pairs()
            if pairs:
                ui_data[tab_name] = pairs
        return ui_data

    def preview_data(self):
        ui_data = self.gather_ui_data()
        if not ui_data:
            QMessageBox.warning(self, "提示", "请至少放入一组完整的 IL / RL 图片！")
            return
        self.start_ocr_task(ui_data, mode="preview")

    def export_ppt(self):
        ui_data = self.gather_ui_data()
        if not ui_data:
            QMessageBox.warning(self, "提示", "请至少放入一组完整的 IL / RL 图片！")
            return
            
        save_path, _ = QFileDialog.getSaveFileName(self, "保存 PPT", "VNA_Report.pptx", "PowerPoint (*.pptx)")
        if not save_path:
            return

        self.save_path = save_path
        self.start_ocr_task(ui_data, mode="export")

    def start_ocr_task(self, ui_data, mode):
        self.mode = mode
        self.progress_dialog = QProgressDialog("正在通过 RapidOCR 极速提取数据...", "取消", 0, 100, self)
        self.progress_dialog.setWindowTitle("处理中")
        self.progress_dialog.setWindowModality(Qt.WindowModal)
        self.progress_dialog.show()

        self.worker = OCRWorker(ui_data)
        self.worker.progress_update.connect(self.update_progress)
        self.worker.finished.connect(self.on_ocr_finished)
        self.worker.start()

    def update_progress(self, val, text):
        self.progress_dialog.setValue(val)
        self.progress_dialog.setLabelText(text)

    def on_ocr_finished(self, result_dataset):
        self.progress_dialog.setValue(100)
        
        if self.mode == "preview":
            self.show_preview_dialog(result_dataset)
        elif self.mode == "export":
            try:
                ppt_gen = PPTGenerator(self.save_path)
                ppt_gen.generate(result_dataset)
                QMessageBox.information(self, "成功", f"PPT已成功导出至:\n{self.save_path}")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"PPT生成失败: {str(e)}")

    def show_preview_dialog(self, result_dataset):
        dialog = QDialog(self)
        dialog.setWindowTitle("数据提取预览")
        dialog.resize(800, 400)
        layout = QVBoxLayout(dialog)
        
        preview_text = ""
        for sample, df in result_dataset.items():
            preview_text += f"=== {sample} ===\n"
            cols_to_show =[c for c in df.columns if 'Img' not in c]
            preview_text += df[cols_to_show].to_string() + "\n\n"
            
        label = QLabel("提取的数据如下（核对无误后请关闭此窗口并点击导出）：")
        text_edit = QTextEdit()
        text_edit.setPlainText(preview_text)
        text_edit.setReadOnly(True)
        text_edit.setStyleSheet("font-family: Consolas; font-size: 14px;")
        
        layout.addWidget(label)
        layout.addWidget(text_edit)
        dialog.exec()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())