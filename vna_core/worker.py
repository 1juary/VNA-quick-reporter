"""
模块 4: OCR 后台工作线程
- 在独立线程中执行 OCR 识别，避免阻塞 UI
"""

import pandas as pd
from PySide6.QtCore import QThread, Signal

from .ocr_extractor import VNAOCRExtractor


class OCRWorker(QThread):
    """OCR 后台工作线程"""

    progress_update = Signal(int, str)
    finished = Signal(dict)

    def __init__(self, ui_data):
        super().__init__()
        self.ui_data = ui_data

    def run(self):
        """执行 OCR 识别任务"""
        extractor = VNAOCRExtractor()
        result_dataset = {}
        total_tasks = sum(len(pairs) for pairs in self.ui_data.values())
        current_task = 0

        for sample_name, pairs in self.ui_data.items():
            sample_data = []
            for pair in pairs:
                current_task += 1
                self.progress_update.emit(
                    int(current_task / total_tasks * 100),
                    f"正在处理: {sample_name}...",
                )

                il_data = extractor.process_image(pair["IL"])
                rl_data = extractor.process_image(pair["RL"])

                sample_data.append(
                    {
                        "PointName": pair["PointName"],
                        "1.5G_IL": il_data.get(1.5),
                        "3.0G_IL": il_data.get(3.0),
                        "4.5G_IL": il_data.get(4.5),
                        "Img_IL": pair["IL"],
                        "1.5G_RL": rl_data.get(1.5),
                        "3.0G_RL": rl_data.get(3.0),
                        "4.5G_RL": rl_data.get(4.5),
                        "Img_RL": pair["RL"],
                    }
                )
            result_dataset[sample_name] = pd.DataFrame(sample_data)

        self.finished.emit(result_dataset)
