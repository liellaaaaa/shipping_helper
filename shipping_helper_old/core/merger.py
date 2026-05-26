# -*- coding: utf-8 -*-
"""
数据汇聚模块
将三份源文件的数据汇聚成统一的结构化数据
"""

import json
import os
from .order_parser import OrderParser
from .pi_extractor import PIExtractor
from .code_matcher import CodeMatcher


class Merger:
    """数据汇聚器"""

    def __init__(self, shipper: str = None):
        self.order_data = {}
        self.pi_data = {}
        self.code_data = {}
        # 如果没有传入发货人，从配置读取
        if shipper is None:
            try:
                from core.config_manager import get_config
                self.shipper = get_config().get_shipper()
            except:
                self.shipper = "广东宏昊化工有限公司"
        else:
            self.shipper = shipper

    def merge(self, order_text: str, pi_filepath: str, code_matcher: CodeMatcher) -> dict:
        """
        汇聚三份源文件的数据
        :param order_text: 外贸销售订单表粘贴的文本
        :param pi_filepath: Proforma Invoice文件路径
        :param code_matcher: 商品编码表匹配器
        :return: 汇聚后的数据字典
        """
        # 1. 解析外贸销售订单表
        parser = OrderParser()
        self.order_data = parser.parse(order_text)

        # 2. 解析Proforma Invoice
        if pi_filepath and os.path.exists(pi_filepath):
            extractor = PIExtractor()
            self.pi_data = extractor.parse_file(pi_filepath)
        else:
            self.pi_data = {}

        # 3. 从商品编码表匹配
        internal_code = parser.get_internal_code()
        if internal_code:
            self.code_data = code_matcher.get_all_info(internal_code)
        else:
            self.code_data = {}

        # 4. 构建汇聚结果
        result = self._build_result()

        return result

    def _build_result(self) -> dict:
        """构建汇聚结果"""
        result = {}

        # 从外贸销售订单表获取
        result['订单号'] = self.order_data.get('订单号', '')
        result['客户编码'] = self.order_data.get('客户编号', '')
        result['业务员'] = self.order_data.get('业务员', '')
        result['内部编号'] = self.order_data.get('内部编号', '')
        result['产品中文名'] = self.order_data.get('产品中文名', '')
        result['报关名称'] = self.order_data.get('报关名称', '') or self.code_data.get('报关名称', '')
        result['规格kg'] = self.order_data.get('规格kg', '')
        result['订单量kg'] = self.order_data.get('订单量kg', '')
        result['订单要求'] = self.order_data.get('订单要求', '')
        result['交货日期'] = self.order_data.get('交货日期', '')

        # 从Proforma Invoice获取
        result['收货人'] = self.pi_data.get('收货人', '')
        result['地址'] = self.pi_data.get('地址', '')
        result['品名英文'] = self.pi_data.get('品名英文', '')
        result['数量'] = self.pi_data.get('数量', '')
        result['单价'] = self.pi_data.get('单价', '')
        result['金额'] = self.pi_data.get('金额', '')
        result['H.S.Code'] = self.pi_data.get('H.S.Code', '')
        result['卸货港'] = self.pi_data.get('卸货港', '')
        result['包装说明'] = self.pi_data.get('包装说明', '')
        result['日期'] = self.pi_data.get('日期', '')

        # 从商品编码表获取
        result['产品颜色'] = self.code_data.get('产品颜色', '')
        result['海关编码'] = self.code_data.get('海关编码', '')
        result['报关成分'] = self.code_data.get('报关成分', '')

        # 固定值
        result['发货人'] = self.shipper

        return result

    def to_json(self) -> str:
        """转换为JSON字符串"""
        return json.dumps(self.order_data, ensure_ascii=False, indent=2)
