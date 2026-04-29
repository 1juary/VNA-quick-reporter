"""
VNA Data Automator - 入口文件
VNA 测试报告自动生成工具
基于 RapidOCR + PySide6 + python-pptx
"""

import sys
from PySide6.QtWidgets import QApplication
from vna_core import MainWindow


def main():
    """应用程序入口"""
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
