# -*- coding: utf-8 -*-
"""
配置管理模块
统一管理presets.json中的所有配置，包括路径和预设值
"""

import os
import json


class ConfigManager:
    """配置管理器"""

    _instance = None
    _config = None

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._load()
        return cls._instance

    def _load(self):
        """加载配置文件"""
        # __file__ is core/config_manager.py, so go up 2 levels to get to shipping_helper root
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        presets_path = os.path.join(base_dir, 'config', 'presets.json')

        if os.path.exists(presets_path):
            with open(presets_path, 'r', encoding='utf-8') as f:
                self._config = json.load(f)
        else:
            self._config = {}

    def get_base_dir(self):
        """获取项目根目录"""
        return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    def get_path(self, key: str) -> str:
        """
        获取配置路径（相对路径转换为绝对路径）
        :param key: paths中的键名
        :return: 绝对路径
        """
        relative_path = self._config.get('paths', {}).get(key, '')
        if not relative_path:
            return ''

        base_dir = self.get_base_dir()
        abs_path = os.path.join(base_dir, relative_path)

        # 规范化路径
        return os.path.normpath(abs_path)

    def get(self, key: str, default=None):
        """获取配置值"""
        return self._config.get(key, default)

    def set(self, key: str, value):
        """设置配置值（仅内存中）"""
        self._config[key] = value

    def get_shipper(self) -> str:
        """获取发货人"""
        return self._config.get('发货人', '')

    def set_shipper(self, value: str):
        """设置发货人"""
        self._config['发货人'] = value

    def get_shipper_address(self) -> str:
        """获取公司地址"""
        return self._config.get('公司地址', '')

    def get_phone(self) -> str:
        """获取电话"""
        return self._config.get('电话', '')

    def get_fax(self) -> str:
        """获取传真"""
        return self._config.get('传真', '')

    def get_all_presets(self) -> dict:
        """获取所有预设值"""
        return {
            '发货人': self.get_shipper(),
            '公司地址': self.get_shipper_address(),
            '电话': self.get_phone(),
            '传真': self.get_fax(),
        }

    def reload(self):
        """重新加载配置"""
        self._load()


# 全局单例
_config_manager = None


def get_config() -> ConfigManager:
    """获取配置管理器单例"""
    global _config_manager
    if _config_manager is None:
        _config_manager = ConfigManager()
    return _config_manager