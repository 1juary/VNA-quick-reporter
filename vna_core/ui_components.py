"""
模块 3: 自定义 UI 组件
- ImageCell: 支持拖拽互换、一键删除、文件拖入的图片单元格
"""

import re
from pathlib import Path
from collections import defaultdict

from PySide6.QtWidgets import (
    QLabel, QPushButton, QVBoxLayout, QFileDialog, QMessageBox
)
from PySide6.QtCore import Qt, Signal, QMimeData
from PySide6.QtGui import QPixmap, QDrag


class ImageCell(QLabel):
    """支持拖拽互换、一键删除、文件拖入的图片单元格"""

    imageLoaded = Signal(str)
    filesDroppedToTab = Signal(list)

    def __init__(self):
        super().__init__()
        self.image_path = ""
        self.setAcceptDrops(True)
        self.drag_start_pos = None

        self.btn_layout = QVBoxLayout(self)
        self.btn_layout.setContentsMargins(5, 5, 5, 5)
        self.btn_layout.setAlignment(Qt.AlignTop | Qt.AlignRight)

        self.btn_delete = QPushButton("×")
        self.btn_delete.setFixedSize(22, 22)
        self.btn_delete.setCursor(Qt.PointingHandCursor)
        self.btn_delete.setStyleSheet("""
            QPushButton {
                background-color: #E74C3C; color: white;
                border-radius: 11px; font-weight: bold;
                font-family: Arial; font-size: 14px; padding-bottom: 2px;
            }
            QPushButton:hover { background-color: #C0392B; }
        """)
        self.btn_delete.clicked.connect(self.clear_image)
        self.btn_delete.hide()

        self.btn_layout.addWidget(self.btn_delete)
        self.reset_ui()

    def reset_ui(self):
        """重置为空白占位状态"""
        self.setText("点击 或 拖拽\n(按住图片可互相交换)")
        self.setAlignment(Qt.AlignCenter)
        self.setCursor(Qt.PointingHandCursor)
        self.setStyleSheet("""
            QLabel {
                background-color: #F8F9FA; color: #95A5A6;
                font-size: 13px; font-weight: bold;
                border: 2px dashed #D1D8E0; border-radius: 6px; margin: 4px;
            }
            QLabel:hover {
                background-color: #E2E8F0; border-color: #3DC2EC; color: #2C3E50;
            }
        """)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.drag_start_pos = event.position().toPoint()

    def mouseMoveEvent(self, event):
        if not self.drag_start_pos or not self.image_path:
            return
        if (event.position().toPoint() - self.drag_start_pos).manhattanLength() > QApplication.startDragDistance():
            drag = QDrag(self)
            mime = QMimeData()
            mime.setText(self.image_path)
            drag.setMimeData(mime)
            drag.setPixmap(self.pixmap().scaled(100, 100, Qt.KeepAspectRatio))
            drag.exec(Qt.MoveAction)
            self.drag_start_pos = None

    def mouseReleaseEvent(self, event):
        if self.drag_start_pos is not None:
            path, _ = QFileDialog.getOpenFileName(
                self, "选择图片", "", "Images (*.png *.jpg *.jpeg)"
            )
            if path:
                self.load_image(path)
            self.drag_start_pos = None

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls() or event.mimeData().hasText():
            event.acceptProposedAction()

    def dropEvent(self, event):
        source = event.source()
        if isinstance(source, ImageCell) and source != self:
            source_path = source.image_path
            target_path = self.image_path

            if source_path:
                self.load_image(source_path)
            else:
                self.clear_image()

            if target_path:
                source.load_image(target_path)
            else:
                source.clear_image()

            event.acceptProposedAction()

        elif event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if len(urls) == 1 and Path(urls[0].toLocalFile()).is_file():
                self.load_image(urls[0].toLocalFile())
                event.acceptProposedAction()
            else:
                self.filesDroppedToTab.emit(urls)
                event.acceptProposedAction()

    def load_image(self, path):
        """加载图片并显示"""
        self.image_path = str(path)
        pixmap = QPixmap(self.image_path)
        self.setPixmap(
            pixmap.scaled(self.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation)
        )
        self.setStyleSheet("border: none; background-color: transparent; margin: 4px;")
        self.btn_delete.show()
        self.imageLoaded.emit(self.image_path)

    def clear_image(self):
        """清除当前图片"""
        self.image_path = ""
        self.setPixmap(QPixmap())
        self.btn_delete.hide()
        self.reset_ui()


def auto_pair_files(img_files):
    """智能配对 IL/RL 图片文件"""
    groups = defaultdict(list)
    for f in img_files:
        stem = Path(f).stem
        base_name = re.sub(r'[-_]([iI][lL]|[rR][lL]|\d+)$', '', stem)
        groups[base_name].append(f)

    final_pairs = []
    orphans = []

    for base, f_list in groups.items():
        f_list = sorted(f_list)

        def intelligent_sort(x):
            xl = x.lower()
            if 'il' in xl and 'rl' not in xl:
                return 0
            if 'rl' in xl and 'il' not in xl:
                return 2
            return 1

        f_list.sort(key=intelligent_sort)

        while len(f_list) >= 2:
            final_pairs.append((base, f_list[0], f_list[1]))
            f_list = f_list[2:]

        if f_list:
            orphans.extend(f_list)

    orphans = sorted(orphans)
    while len(orphans) >= 2:
        o1, o2 = orphans[0], orphans[1]
        stem = Path(o1).stem
        base = re.sub(r'[-_]([iI][lL]|[rR][lL]|\d+)$', '', stem)
        final_pairs.append((base, o1, o2))
        orphans = orphans[2:]

    return final_pairs, orphans
