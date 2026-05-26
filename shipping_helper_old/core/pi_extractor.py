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

        # 获取所有文本内容
        full_text = []
        for i in range(ws.nrows):
            row_text = []
            for j in range(ws.ncols):
                val = ws.cell_value(i, j)
                if val:
                    row_text.append(str(val).strip())
            full_text.append(row_text)

        # 解析各个字段
        self._extract_consignee(full_text)
        self._extract_date(full_text)
        self._extract_pi_no(full_text)
        self._extract_goods(full_text)
        self._extract_quantity(full_text)
        self._extract_hs_code(full_text)
        self._extract_destination(full_text)
        self._extract_packing(full_text)

        return self.data

    def _find_text(self, full_text, keywords):
        """在文本中查找包含关键词的行"""
        for row in full_text:
            for cell in row:
                if any(kw.lower() in cell.lower() for kw in keywords):
                    return row
        return None

    def _extract_consignee(self, full_text):
        """提取收货人信息"""
        # 查找 Name : 后面的收货人名称（通常在 Consignee 行之后）
        for i, row in enumerate(full_text):
            for cell in row:
                if 'Name :' in cell:
                    self.data['收货人'] = cell.replace('Name :', '').strip()
                    break
            if '收货人' in self.data:
                break

        # 如果上面没找到，在 Name : 同一行的下一格查找
        if '收货人' not in self.data:
            for row in full_text:
                for i, cell in enumerate(row):
                    if 'Name :' in str(cell):
                        if i + 1 < len(row) and row[i + 1]:
                            self.data['收货人'] = str(row[i + 1]).strip()
                            break
                if '收货人' in self.data:
                    break

        # 查找地址 - Add.: 通常在收货人信息附近
        for row in full_text:
            row_str = ' '.join([str(c) for c in row])
            if 'Add.:' in row_str or 'Add:' in row_str:
                for cell in row:
                    if str(cell).strip().lower().startswith('add.:'):
                        addr = str(cell).replace('Add.:', '').replace('Add:', '').strip()
                        # 收集同一行后面的内容作为地址
                        for j in range(row.index(cell) + 1, len(row)):
                            if row[j]:
                                addr += ' ' + str(row[j]).strip()
                        self.data['地址'] = addr.strip()
                        break
            if '地址' in self.data:
                break

    def _extract_date(self, full_text):
        """提取日期"""
        # 查找 Date : 标签后面跟着实际日期值（如 2026-3-4）
        for row in full_text:
            for cell in row:
                cell_str = str(cell).strip()
                if 'Date :' in cell_str or cell_str.startswith('Date:'):
                    date_str = cell_str.replace('Date :', '').replace('Date:', '').strip()
                    # 如果提取到的看起来像数字（Excel序列号如46085），跳过
                    if date_str and not date_str.replace('.', '').isdigit():
                        self.data['日期'] = date_str
                        break
                    # 如果同一行有后续单元格且看起来像日期
                    idx = row.index(cell)
                    for j in range(idx + 1, len(row)):
                        next_val = str(row[j]).strip()
                        if next_val and '20' in next_val:
                            self.data['日期'] = next_val
                            break
            if '日期' in self.data:
                break

        # 备选方案：在整个文本中搜索 2026 开头的日期格式
        if '日期' not in self.data:
            for row in full_text:
                for cell in row:
                    cell_str = str(cell).strip()
                    if '20' in cell_str and ('-' in cell_str or '/' in cell_str):
                        # 可能是日期
                        match = re.search(r'20\d[/-]\d[/-]\d', cell_str)
                        if match:
                            self.data['日期'] = match.group()
                            break
                if '日期' in self.data:
                    break

    def _extract_pi_no(self, full_text):
        """提取PI号"""
        for row in full_text:
            row_str = ' '.join(row)
            if 'PI No.:' in row_str or 'PI No.:' in row_str:
                # 查找PI号
                for cell in row:
                    if 'HT' in cell and ('PI' in cell or len(cell) > 5):
                        # 可能是PI号
                        match = re.search(r'HT\d+[A-Z0-9]+', cell)
                        if match:
                            self.data['PI号'] = match.group()
                            break
                    # 直接是PI号的情况
                    match = re.search(r'HT\d+[A-Z0-9]+', cell)
                    if match:
                        self.data['PI号'] = match.group()
                        break
            if 'PI号' in self.data:
                break

        # 直接在文本中搜索HT开头的PI号
        if 'PI号' not in self.data:
            all_text = ' '.join([' '.join(row) for row in full_text])
            match = re.search(r'HT\d+[A-Z0-9]+', all_text)
            if match:
                self.data['PI号'] = match.group()

    def _extract_goods(self, full_text):
        """提取商品信息"""
        # 查找 DESCRIPTION OF GOODS 行之后的商品数据
        in_goods_section = False
        for row in full_text:
            row_str = ' '.join(row)
            if 'DESCRIPTION OF GOODS' in row_str:
                in_goods_section = True
                continue
            if in_goods_section and row:
                # 商品行通常包含数量和价格
                if len(row) >= 2:
                    # 检查是否包含数字（数量或金额）
                    has_number = any(
                        isinstance(x, (int, float)) or (isinstance(x, str) and x.replace('.', '').isdigit())
                        for x in row
                    )
                    if has_number and len(row[0]) > 3:
                        self.data['品名英文'] = row[0].strip()
                        break

    def _extract_quantity(self, full_text):
        """提取数量、单价、金额"""
        for row in full_text:
            if len(row) >= 4:
                # 检查是否是商品行（有4列：品名、数量、单价、金额）
                try:
                    qty = row[1]
                    cif = row[2]
                    amount = row[3]
                    if isinstance(qty, (int, float)) and isinstance(amount, (int, float)):
                        self.data['数量'] = qty
                        self.data['单价'] = cif
                        self.data['金额'] = amount
                        break
                except (IndexError, ValueError):
                    continue

        # 检查是否有 TOTAL 行
        for row in full_text:
            if 'TOTAL' in ' '.join(row) and len(row) >= 3:
                try:
                    qty = row[1]
                    amount = row[2]
                    if isinstance(qty, (int, float)) and isinstance(amount, (int, float)):
                        if '数量' not in self.data:
                            self.data['数量'] = qty
                        if '金额' not in self.data:
                            self.data['金额'] = amount
                        break
                except:
                    continue

    def _extract_hs_code(self, full_text):
        """提取H.S. Code"""
        for row in full_text:
            for cell in row:
                if 'H.S. Code' in cell or 'H.S. Code' in cell:
                    match = re.search(r'(\d{10})', cell)
                    if match:
                        self.data['H.S.Code'] = match.group(1)
                        break
            if 'H.S.Code' in self.data:
                break

    def _extract_destination(self, full_text):
        """提取卸货港"""
        for row in full_text:
            for cell in row:
                if 'Destination' in cell and ':' in cell:
                    dest = cell.split(':', 1)[1].strip()
                    if dest:
                        self.data['卸货港'] = dest
                        break
            if '卸货港' in self.data:
                break

    def _extract_packing(self, full_text):
        """提取包装说明"""
        for row in full_text:
            for cell in row:
                if 'Packing' in cell and ':' in cell:
                    packing = cell.split(':', 1)[1].strip()
                    if packing:
                        self.data['包装说明'] = packing
                        break
            if '包装说明' in self.data:
                break

        # 检查 REMARKS 部分的包装信息
        if '包装说明' not in self.data:
            for row in full_text:
                row_str = ' '.join(row)
                if 'Packing:' in row_str or 'export standard' in row_str.lower():
                    for cell in row:
                        if 'export standard' in cell.lower() or 'kg' in cell.lower():
                            self.data['包装说明'] = cell.strip()
                            break
                    if '包装说明' not in self.data:
                        self.data['包装说明'] = row_str.strip()
                    break

    def get_all_data(self) -> dict:
        """获取所有解析数据"""
        return self.data.copy()
