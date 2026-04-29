"""
模块 5: 主窗口 UI
- 包含 SampleTab（样品标签页）和 MainWindow（主窗口）
- 负责所有 UI 布局、信号连接、业务流调度
"""

import re
import json
from pathlib import Path
from collections import defaultdict

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QTabWidget, QTableWidget,
    QHeaderView, QProgressDialog, QMessageBox, QDialog, QTextEdit,
    QLineEdit, QComboBox, QFrame,
)
from PySide6.QtCore import Qt, QThread, Signal

from .ui_components import ImageCell, auto_pair_files
from .worker import OCRWorker
from .ppt_generator import PPTGenerator
from .file_utils import load_settings, save_settings


class SampleTab(QWidget):
    """单个样品的标签页，包含点位名称编辑和 IL/RL 图片选择"""

    def __init__(self):
        super().__init__()
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(15, 15, 15, 15)

        self.table = QTableWidget(3, 3)
        self.table.setHorizontalHeaderLabels(["测试点位", "IL IMAGE", "RL IMAGE"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Fixed)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.table.setColumnWidth(0, 110)
        self.table.verticalHeader().setDefaultSectionSize(120)
        self.table.setSelectionMode(QTableWidget.NoSelection)
        self.table.setFocusPolicy(Qt.NoFocus)

        for row in range(3):
            self.init_row_widgets(row)
        self.layout.addWidget(self.table)

        self.btn_add_row = QPushButton("+ 添加点位")
        self.btn_add_row.setObjectName("BtnAddRow")
        self.btn_add_row.clicked.connect(self.add_row)
        self.layout.addWidget(self.btn_add_row, alignment=Qt.AlignCenter)

    def init_row_widgets(self, row):
        """初始化一行的三个控件：点位名输入框 + IL/RL 图片单元格"""
        point_edit = QLineEdit(f"点位{row+1}")
        point_edit.setAlignment(Qt.AlignCenter)
        point_edit.setStyleSheet(
            "border: none; background: transparent; font-weight: bold; "
            "color: #2C3E50; font-size: 14px;"
        )
        self.table.setCellWidget(row, 0, point_edit)

        il_cell = ImageCell()
        il_cell.imageLoaded.connect(lambda path, r=row: self.auto_fill_point_name(r, path))
        il_cell.filesDroppedToTab.connect(self.handle_dropped_files)
        self.table.setCellWidget(row, 1, il_cell)

        rl_cell = ImageCell()
        rl_cell.imageLoaded.connect(lambda path, r=row: self.auto_fill_point_name(r, path))
        rl_cell.filesDroppedToTab.connect(self.handle_dropped_files)
        self.table.setCellWidget(row, 2, rl_cell)

    def add_row(self):
        """新增一行"""
        row = self.table.rowCount()
        self.table.insertRow(row)
        self.init_row_widgets(row)

    def auto_fill_point_name(self, row, path):
        """根据图片文件名自动填充点位名称"""
        stem = Path(path).stem
        clean_name = re.sub(r'[-_]([iI][lL]|[rR][lL]|\d+)$', '', stem)
        edit_widget = self.table.cellWidget(row, 0)
        current_text = edit_widget.text().strip()
        if current_text.startswith("点位") or not current_text:
            edit_widget.setText(clean_name)

    def handle_dropped_files(self, urls):
        """处理拖入的文件/文件夹，自动配对 IL/RL"""
        files = []
        for url in urls:
            path = Path(url.toLocalFile())
            if path.is_dir():
                files.extend(
                    [str(p) for p in path.rglob("*") if p.is_file()]
                )
            elif path.is_file():
                files.append(str(path))

        img_files = sorted(
            [f for f in files if f.lower().endswith(('.jpg', '.jpeg', '.png'))]
        )
        if not img_files:
            return

        final_pairs, orphans = auto_pair_files(img_files)

        for base, il_path, rl_path in final_pairs:
            row = self.find_empty_row_or_add()
            self.table.cellWidget(row, 0).setText(base)
            self.table.cellWidget(row, 1).load_image(il_path)
            self.table.cellWidget(row, 2).load_image(rl_path)

        if orphans:
            QMessageBox.information(
                self, "警告",
                f"由于文件数为奇数，有 {len(orphans)} 个文件未能配对！"
            )

    def find_empty_row_or_add(self):
        """查找空行，如果没有则新增"""
        for row in range(self.table.rowCount()):
            if (not self.table.cellWidget(row, 1).image_path
                    and not self.table.cellWidget(row, 2).image_path):
                return row
        row = self.table.rowCount()
        self.add_row()
        return row

    def get_image_pairs(self):
        """获取当前标签页中所有已配对的图片数据"""
        pairs = []
        for row in range(self.table.rowCount()):
            point_name = self.table.cellWidget(row, 0).text().strip()
            il_cell = self.table.cellWidget(row, 1)
            rl_cell = self.table.cellWidget(row, 2)
            if il_cell.image_path and rl_cell.image_path:
                pairs.append({
                    'PointName': point_name if point_name else f"点位{row+1}",
                    'IL': il_cell.image_path,
                    'RL': rl_cell.image_path,
                })
        return pairs


class MainWindow(QMainWindow):
    """VNA Data Automator 主窗口"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("VNA Data Automator")
        self.resize(1000, 780)
        self.setAcceptDrops(True)
        self.config_map = load_settings()
        self.init_ui()
        self.apply_styles()

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        current_tab = self.tabs.currentWidget()
        if hasattr(current_tab, 'handle_dropped_files'):
            current_tab.handle_dropped_files(urls)

    def init_ui(self):
        """初始化 UI 布局"""
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)
        main_layout.setSpacing(15)

        # --- 顶部信息栏 ---
        top_card = QFrame()
        top_card.setObjectName("InfoCard")
        top_layout = QHBoxLayout(top_card)
        top_layout.setContentsMargins(20, 15, 20, 15)

        lbl_lang = QLabel("报告语言 :")
        lbl_lang.setObjectName("InfoTitle")
        self.combo_lang = QComboBox()
        self.combo_lang.addItems(["English", "中文"])
        self.combo_lang.setFocusPolicy(Qt.NoFocus)

        lbl_proj = QLabel("项目名 :")
        lbl_proj.setObjectName("InfoTitle")
        self.edit_proj = QLineEdit()
        self.edit_proj.setPlaceholderText("例如: Project X")
        self.edit_proj.textChanged.connect(self.on_project_name_changed)

        lbl_spec = QLabel("规 格 :")
        lbl_spec.setObjectName("InfoTitle")
        self.combo_spec = QComboBox()
        self.combo_spec.setEditable(True)
        self.combo_spec.setPlaceholderText("请选择或手动输入规格")
        self.combo_spec.addItems(list(set(self.config_map.values())))

        top_layout.addWidget(lbl_lang)
        top_layout.addWidget(self.combo_lang, stretch=1)
        top_layout.addSpacing(20)
        top_layout.addWidget(lbl_proj)
        top_layout.addWidget(self.edit_proj, stretch=2)
        top_layout.addSpacing(20)
        top_layout.addWidget(lbl_spec)
        top_layout.addWidget(self.combo_spec, stretch=2)
        main_layout.addWidget(top_card)

        # --- 样品标签页 ---
        self.tabs = QTabWidget()
        self.tabs.addTab(SampleTab(), "样品1")
        self.tabs.addTab(SampleTab(), "样品2")
        self.tabs.addTab(SampleTab(), "样品3")

        self.btn_add_tab = QPushButton("+ 新增样品")
        self.btn_add_tab.setObjectName("ChartButton")
        self.btn_add_tab.setCursor(Qt.PointingHandCursor)
        self.btn_add_tab.clicked.connect(
            lambda: self.tabs.addTab(SampleTab(), f"样品{self.tabs.count()+1}")
        )
        self.tabs.setCornerWidget(self.btn_add_tab, Qt.TopRightCorner)
        main_layout.addWidget(self.tabs)

        # --- 底部操作按钮 ---
        bottom_layout = QHBoxLayout()
        bottom_layout.addStretch()

        self.btn_preview = QPushButton("预览数据")
        self.btn_preview.setObjectName("ActionButton")
        self.btn_preview.setCursor(Qt.PointingHandCursor)
        self.btn_preview.clicked.connect(self.preview_data)

        self.btn_export = QPushButton("🚀 导出完整报告")
        self.btn_export.setObjectName("ActionButton")
        self.btn_export.setCursor(Qt.PointingHandCursor)
        self.btn_export.clicked.connect(self.export_ppt)

        bottom_layout.addWidget(self.btn_preview)
        bottom_layout.addSpacing(15)
        bottom_layout.addWidget(self.btn_export)
        bottom_layout.addStretch()
        main_layout.addLayout(bottom_layout)

    def on_project_name_changed(self, text):
        """项目名变更时自动填充对应的规格"""
        text = text.strip()
        if text in self.config_map:
            self.combo_spec.setCurrentText(self.config_map[text])

    def apply_styles(self):
        """应用 QSS 样式"""
        self.setStyleSheet("""
            QMainWindow { background-color: #F4F6F8; }
            QFrame#InfoCard {
                background-color: #FFFFFF; border-radius: 12px;
                border: 1px solid #D1D8E0;
            }
            QLabel#InfoTitle {
                font-size: 14px; font-weight: 800; color: #2C3E50;
            }
            QLineEdit, QComboBox {
                border: 1px solid #D1D8E0; border-radius: 6px;
                padding: 6px 10px; font-size: 13px; color: #2C3E50;
                background-color: #F8F9FA;
            }
            QLineEdit:focus, QComboBox:focus {
                border-color: #3DC2EC; background-color: #FFFFFF;
            }
            QTabWidget::pane {
                border: 1px solid #D1D8E0; border-radius: 8px;
                background-color: #FFFFFF;
            }
            QTabBar::tab {
                background-color: #E2E8F0; color: #7F8C8D;
                padding: 10px 25px; margin-right: 4px;
                border-top-left-radius: 8px; border-top-right-radius: 8px;
                font-weight: bold; font-size: 14px;
            }
            QTabBar::tab:selected {
                background-color: #FFFFFF; color: #2C3E50;
                border: 1px solid #D1D8E0; border-bottom: none;
            }
            QTableWidget {
                background-color: #FFFFFF; border: none;
                gridline-color: #E9ECEF;
            }
            QHeaderView::section {
                background-color: #F8F9FA; color: #34495E;
                font-weight: bold; font-size: 14px; border: none;
                border-bottom: 2px solid #D1D8E0; padding: 12px;
            }
            QPushButton#BtnAddRow {
                background-color: transparent; font-size: 14px;
                font-weight: bold; color: #95A5A6; padding: 10px;
            }
            QPushButton#BtnAddRow:hover { color: #3DC2EC; }
            QPushButton#ChartButton {
                background-color: transparent; border-radius: 8px;
                font-size: 13px; font-weight: bold; color: #34495E;
                padding: 5px 15px; margin-top: 5px;
            }
            QPushButton#ChartButton:hover { background-color: #E2E8F0; }
            QPushButton#ActionButton {
                background-color: #3DC2EC; color: #FFFFFF;
                font-size: 15px; font-weight: bold; border: none;
                border-radius: 20px; padding: 12px 35px;
            }
            QPushButton#ActionButton:hover { background-color: #5ED1F4; }
            QPushButton#ActionButton:pressed { background-color: #2BAAD4; }
        """)

    def gather_ui_data(self):
        """从所有标签页收集图片配对数据"""
        ui_data = {}
        for i in range(self.tabs.count()):
            tab_name = self.tabs.tabText(i)
            tab_widget = self.tabs.widget(i)
            pairs = tab_widget.get_image_pairs()
            if pairs:
                ui_data[tab_name] = pairs
        return ui_data

    def preview_data(self):
        """预览 OCR 提取结果"""
        ui_data = self.gather_ui_data()
        if not ui_data:
            QMessageBox.warning(self, "提示", "请至少放入一组完整的 IL / RL 图片！")
            return
        self.start_ocr_task(ui_data, mode="preview")

    def export_ppt(self):
        """导出 PPT 报告"""
        ui_data = self.gather_ui_data()
        if not ui_data:
            QMessageBox.warning(self, "提示", "请至少放入一组完整的 IL / RL 图片！")
            return

        save_path, _ = QFileDialog.getSaveFileName(
            self, "保存 PPT", "VNA_Report.pptx", "PowerPoint (*.pptx)"
        )
        if not save_path:
            return

        proj = self.edit_proj.text().strip()
        spec = self.combo_spec.currentText().strip()
        save_settings(self.config_map, proj, spec)

        self.save_path = save_path
        self.start_ocr_task(ui_data, mode="export")

    def start_ocr_task(self, ui_data, mode):
        """启动后台 OCR 线程"""
        self.mode = mode
        self.current_lang = (
            "en" if self.combo_lang.currentText() == "English" else "zh"
        )

        self.progress_dialog = QProgressDialog(
            "正在通过 RapidOCR 极速提取数据...", "取消", 0, 100, self
        )
        self.progress_dialog.setWindowTitle("处理中")
        self.progress_dialog.setWindowModality(Qt.WindowModal)
        self.progress_dialog.show()

        self.worker = OCRWorker(ui_data)
        self.worker.progress_update.connect(self.progress_dialog.setValue)
        self.worker.progress_update.connect(
            lambda v, t: self.progress_dialog.setLabelText(t)
        )
        self.worker.finished.connect(self.on_ocr_finished)
        self.worker.start()

    def on_ocr_finished(self, result_dataset):
        """OCR 完成后的回调"""
        self.progress_dialog.setValue(100)

        if self.mode == "preview":
            self.show_preview_dialog(result_dataset)
        elif self.mode == "export":
            try:
                ppt_gen = PPTGenerator(self.save_path)
                ppt_gen.generate(
                    result_dataset,
                    proj_name=self.edit_proj.text().strip(),
                    spec=self.combo_spec.currentText().strip(),
                    lang=self.current_lang,
                )
                QMessageBox.information(
                    self, "成功", f"PPT已成功导出至:\n{self.save_path}"
                )
            except Exception as e:
                QMessageBox.critical(self, "错误", f"PPT生成失败: {str(e)}")

    def show_preview_dialog(self, result_dataset):
        """显示数据预览对话框"""
        dialog = QDialog(self)
        dialog.setWindowTitle("数据提取预览")
        dialog.resize(800, 400)
        layout = QVBoxLayout(dialog)

        preview_text = ""
        for sample, df in result_dataset.items():
            preview_text += f"=== {sample} ===\n"
            cols_to_show = [c for c in df.columns if 'Img' not in c]
            preview_text += df[cols_to_show].to_string() + "\n\n"

        label = QLabel("提取的数据如下（核对无误后请关闭此窗口并点击导出）：")
        text_edit = QTextEdit()
        text_edit.setPlainText(preview_text)
        text_edit.setReadOnly(True)
        text_edit.setStyleSheet("font-family: Consolas; font-size: 14px;")

        layout.addWidget(label)
        layout.addWidget(text_edit)
        dialog.exec()
