"""
文件工具模块
- 配置文件的读写（项目名-规格映射）
"""

import os
import json

CONFIG_FILE = "vna_config.json"


def load_settings():
    """从 JSON 配置文件加载项目名-规格映射"""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_settings(config_map, proj, spec):
    """保存项目名-规格映射到 JSON 配置文件"""
    if proj:
        config_map[proj] = spec
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(config_map, f, ensure_ascii=False, indent=2)
        except Exception:
            pass
