# -*- coding: utf-8 -*-
"""
Proforma Invoice 解析模块
解析自由文本格式的PI文件（.xls格式）
"""

import xlrd
import re
from datetime import datetime


class PIExtractor:
    """解析 Proforma Invoice 文件"""

    def __init__(self):
        self.data = {}

    def parse_file(self, filepath: str) -> dict:
        """
        解析Proforma Invoice文件
        :param filepath: .xls文件路径
        :return: 解析后的字典
        """
        wb = xlrd.open_workbook(filepath, encoding_override='utf-8')
        ws = wb.sheet_by_index(0)

        # 使用行索引直接提取关键字段（基于样本结构）
        self._extract_consignee(ws)
        self._extract_date(ws)
        self._extract_pi_no(ws)
        self._extract_goods(ws)
        self._extract_quantity(ws)
        self._extract_hs_code(ws)
        self._extract_destination(ws)
        self._extract_packing(ws)

        return self.data

    def _get_cell(self, ws, row, col):
        """安全获取单元格值"""
        try:
            val = ws.cell_value(row, col)
            return str(val).strip() if val else ''
        except:
            return ''

    def _extract_consignee(self, ws):
        """提取收货人信息"""
        # row[8][0]: Name : TOA-DOVECHEM INDUSTRIES CO., LTD
        # row[10][0]: Add.: 31/5 Moo 3...
        name_cell = self._get_cell(ws, 8, 0)
        if 'Name :' in name_cell:
            self.data['收货人'] = name_cell.replace('Name :', '').strip()

        addr_cell = self._get_cell(ws, 10, 0)
        if 'Add.:' in addr_cell:
            self.data['收货人地址'] = addr_cell.replace('Add.:', '').strip()

    def _extract_date(self, ws):
        """提取日期 - 转换Excel序列号"""
        # row[8][5] 是 Excel 日期序列号
        try:
            val = ws.cell_value(8, 5)
            if isinstance(val, float) and 45000 < val < 60000:
                date = datetime(*xlrd.xldate_as_tuple(val, 0)[:3])
                self.data['日期'] = date.strftime('%Y-%m-%d')
        except:
            pass

    def _extract_pi_no(self, ws):
        """提取PI号 - row[13][5]"""
        pi_no = self._get_cell(ws, 13, 5)
        if 'HT' in pi_no:
            self.data['PI号'] = pi_no

    def _extract_goods(self, ws):
        """提取品名英文 - row[15][0]"""
        goods = self._get_cell(ws, 15, 0)
        if '(' in goods and ')' in goods:
            self.data['品名英文'] = goods

    def _extract_quantity(self, ws):
        """提取数量、单价、金额 - row[15][2], [3], [5]"""
        try:
            qty = ws.cell_value(15, 2)
            cif = ws.cell_value(15, 3)
            amount = ws.cell_value(15, 5)
            if isinstance(qty, float) and isinstance(amount, float):
                self.data['数量'] = qty
                self.data['单价'] = cif
                self.data['金额'] = amount
        except:
            pass

    def _extract_hs_code(self, ws):
        """提取H.S. Code - row[26][0]"""
        hs_cell = self._get_cell(ws, 26, 0)
        if 'H.S. Code' in hs_cell:
            match = re.search(r'(\d{10})', hs_cell)
            if match:
                self.data['H.S.Code'] = match.group(1)

    def _extract_destination(self, ws):
        """提取卸货港 - row[28][0]"""
        dest_cell = self._get_cell(ws, 28, 0)
        if 'Destination' in dest_cell and ':' in dest_cell:
            dest = dest_cell.split(':', 1)[1].strip()
            if dest:
                self.data['卸货港'] = dest

    def _extract_packing(self, ws):
        """提取包装说明 - REMARKS部分"""
        packing_cell = self._get_cell(ws, 22, 2)
        if 'Packing:' in packing_cell:
            self.data['包装说明'] = packing_cell.replace('Packing:', '').strip()

    def get_all_data(self) -> dict:
        """获取所有解析数据"""
        return self.data.copy()