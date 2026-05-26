# -*- coding: utf-8 -*-
"""
商品编码表匹配模块
从内置的商品编码表JSON中根据内部编号匹配数据
"""

import json
import os


class CodeMatcher:
    """商品编码表匹配器"""

    def __init__(self, codes_file: str = None):
        """
        初始化匹配器
        :param codes_file: 商品编码表JSON文件路径
        """
        self.codes_file = codes_file
        self.codes = []
        self.version = ""

    def load(self, codes_file: str = None) -> bool:
        """
        加载商品编码表
        :param codes_file: 商品编码表JSON文件路径
        :return: 是否加载成功
        """
        if codes_file:
            self.codes_file = codes_file

        if not self.codes_file or not os.path.exists(self.codes_file):
            return False

        try:
            with open(self.codes_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                # 支持两种格式：codes数组 或 products对象
                if 'codes' in data:
                    self.codes = data.get('codes', [])
                elif 'products' in data:
                    # 转换为字典数组格式
                    self.codes = []
                    for code, info in data.get('products', {}).items():
                        item = dict(info)
                        item['产品内编'] = code
                        self.codes.append(item)
                self.version = data.get('version', '')
            return True
        except Exception as e:
            print(f"加载商品编码表失败: {e}")
            return False

    def match(self, internal_code: str) -> dict:
        """
        根据内部编号匹配商品信息
        :param internal_code: 内部编号（如 LB-001, K70-EB）
        :return: 匹配到的商品信息字典，未找到返回空字典
        """
        if not internal_code:
            return {}

        internal_code = internal_code.strip().upper()

        for item in self.codes:
            item_code = item.get('产品内编', '').strip().upper()
            if item_code == internal_code:
                return item.copy()

        # 精确匹配失败，尝试模糊匹配
        return self._fuzzy_match(internal_code)

    def _fuzzy_match(self, internal_code: str) -> dict:
        """模糊匹配"""
        for item in self.codes:
            item_code = item.get('产品内编', '').strip().upper()
            # 检查是否包含
            if internal_code in item_code or item_code in internal_code:
                return item.copy()
        return {}

    def get_product_color(self, internal_code: str) -> str:
        """获取产品外观（颜色）"""
        item = self.match(internal_code)
        return item.get('产品外观', '')

    def get_customs_code(self, internal_code: str) -> str:
        """获取新商品编码（海关编码）"""
        item = self.match(internal_code)
        code = item.get('新商品编码', '')
        return str(code) if code else ''

    def get_composition(self, internal_code: str) -> str:
        """获取报关成分"""
        item = self.match(internal_code)
        return item.get('成分', '')

    def get_customs_name(self, internal_code: str) -> str:
        """获取报关名称"""
        item = self.match(internal_code)
        return item.get('报关名称', '')

    def get_all_info(self, internal_code: str) -> dict:
        """
        获取商品的完整信息
        :param internal_code: 内部编号
        :return: 包含产品颜色、海关编码、报关成分等字段的字典
        """
        item = self.match(internal_code)
        if not item:
            return {}

        return {
            '产品颜色': item.get('产品外观', ''),
            '海关编码': str(item.get('新商品编码', '')),
            '报关成分': item.get('成分', ''),
            '报关名称': item.get('报关名称', ''),
            '用途': item.get('用途', ''),
        }

    def get_version(self) -> str:
        """获取编码表版本"""
        return self.version