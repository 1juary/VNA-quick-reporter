"""
模块 1: RapidOCR 图像处理与提取模块
- 负责图像裁剪、OCR 识别、频率-dB 值解析
- 支持多区域自适应裁剪策略
"""

import re
import os
import cv2
import numpy as np
from rapidocr_onnxruntime import RapidOCR


class VNAOCRExtractor:
    """VNA 测试截图 OCR 提取器"""

    def __init__(self):
        self.ocr = RapidOCR()
        self.target_freqs = [1.5, 3.0, 4.5]
        self.pattern = re.compile(
            r'(\d+\.\d+)\s*[Gg][Hh][Zz][^\d-]{0,25}(-?\d+\.\d+)\s*[dD][Bb]'
        )

    def process_image(self, img_path):
        """处理单张图片并提取所需频率的 dB 值"""
        if not img_path or not os.path.exists(img_path):
            return {1.5: "N/A", 3.0: "N/A", 4.5: "N/A"}

        try:
            img = cv2.imdecode(
                np.fromfile(img_path, dtype=np.uint8), cv2.IMREAD_COLOR
            )
        except Exception:
            img = None

        if img is None:
            return {1.5: "N/A", 3.0: "N/A", 4.5: "N/A"}

        h, w = img.shape[:2]

        # 策略 1: 默认右上角裁剪
        crop_default = img[
            int(h * 0.10) : int(h * 0.28), int(w * 0.65) : int(w * 0.88)
        ]
        res_default = self._extract(crop_default)
        if self._is_complete(res_default):
            return res_default

        # 策略 2: 偏移区域裁剪
        crop_shifted = img[
            int(h * 0.05) : int(h * 0.35), int(w * 0.35) : int(w * 0.80)
        ]
        res_shifted = self._extract(crop_shifted)
        if self._has_any_data(res_shifted):
            return res_shifted

        # 策略 3: 全图上半部分
        crop_extreme = img[0 : int(h * 0.50), 0:w]
        return self._extract(crop_extreme)

    def _extract(self, crop):
        """对裁剪区域执行 OCR 并解析结果"""
        result, _ = self.ocr(crop)
        text_content = ""

        if result:
            for line in result:
                text_content += str(line[1]) + " "

        results = {1.5: "N/A", 3.0: "N/A", 4.5: "N/A"}
        matches = self.pattern.findall(text_content)

        for match in matches:
            try:
                freq = float(match[0])
                db_val = f"{float(match[1]):.2f}dB"
                closest_freq = min(
                    self.target_freqs, key=lambda x: abs(x - freq)
                )
                if abs(closest_freq - freq) < 0.2:
                    if results[closest_freq] == "N/A":
                        results[closest_freq] = db_val
            except ValueError:
                continue
        return results

    @staticmethod
    def _is_complete(res):
        """检查是否提取到了全部三个目标频率"""
        return sum(1 for v in res.values() if v != "N/A") == 3

    @staticmethod
    def _has_any_data(res):
        """检查是否提取到了至少一个有效数据"""
        return sum(1 for v in res.values() if v != "N/A") > 0
