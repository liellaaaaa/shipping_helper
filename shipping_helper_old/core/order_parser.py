# -*- coding: utf-8 -*-
"""
外贸销售订单表粘贴解析模块
从在线文档复制的一行数据（Tab分隔）进行解析
"""

import re


class OrderParser:
    """解析外贸销售订单表粘贴的数据"""

    COLUMNS = [
        '业务员', '客户编号', '内部编号', '产品中文名', '报关名称', '规格kg',
        '订单量kg', '是否调价', '有无样品', '订单要求', '交货日期', '审核',
        '销售区域', '订单号', '出货抬头', '单据类型', '跟单员', '下单日期',
        '确认下单', '生产交期', '出货渠道', '出货方式', '规格异常'
    ]

    def __init__(self):
        self.data = {}

    def parse(self, text: str) -> dict:
        """
        解析粘贴的文本
        :param text: Tab分隔或按行分隔的文本
        :return: 解析后的字典
        """
        # 预处理：去除前后空白
        text = text.strip()

        # 如果是空文本，直接返回空字典
        if not text:
            return {}

        # 判断分隔方式
        has_tabs = '\t' in text
        newline_count = text.count('\n')

        # 核心逻辑：优先判断是Tab分隔还是换行分隔
        if has_tabs:
            # 有Tab，说明是企业微信复制的数据（Tab分隔）
            # 如果第一行是空的，说明开头有空行，取第二行
            lines = text.split('\n')
            if lines and not lines[0].strip():
                lines = lines[1:]
            if lines:
                first_line = lines[0]
                parts = first_line.split('\t')
            else:
                parts = []
        elif newline_count >= 10:
            # 换行分隔（每行一个字段）
            lines = text.split('\n')
            parts = [line for line in lines if line.strip()]
        else:
            # 空格分隔（Tab被转换了）
            parts = text.split()

        # 清理每个字段
        parts = [p.strip() for p in parts]

        # 【重要】跳过行号列：如果第一个字段是纯数字，则跳过
        # （Excel列A可能是行号列）
        if parts and parts[0].isdigit():
            parts = parts[1:]

        # 【重要】智能修复列偏移：
        # 如果报关名称位置是纯数字，说明报关名称实际为空，需要插入空值
        parts = self._fix_column_offset(parts)

        # 【重要】检测并跳过表头行：如果第一个字段是"业务员"，则跳过
        if parts and parts[0] == '业务员':
            parts = parts[1:]

        # 如果字段数量超过23，说明"订单要求"字段内有嵌入换行，需要智能合并
        if len(parts) > len(self.COLUMNS):
            parts = self._collapse_extra_fields(parts)
        elif len(parts) == len(self.COLUMNS):
            # 刚好23个字段，检查是否有字段内容包含换行符（嵌入换行）
            # 这种情况不需要合并，各字段已经对齐
            pass

        # 确保有足够的列
        while len(parts) < len(self.COLUMNS):
            parts.append('')

        # 限制列数
        parts = parts[:len(self.COLUMNS)]

        # 构建字典
        for i, col in enumerate(self.COLUMNS):
            self.data[col] = parts[i] if i < len(parts) else ''

        return self.data

    def _fix_column_offset(self, parts: list) -> list:
        """
        智能修复列偏移问题。

        报关名称（index 4）必定为中文或空。如果是纯数字，说明报关名称实际为空。

        策略：
        - 如果报关名称位置是纯数字，检查规格kg位置(index 5)是否有内容
        - 如果规格kg位置也是纯数字，说明数据正常（这两个就是规格kg和订单量kg）
        - 如果规格kg位置为空或不存在，需要在报关名称位置插入空值
        """
        if len(parts) <= 4:
            return parts

        col_e = parts[4] if len(parts) > 4 else ''

        # 如果报关名称是纯数字（这是规格kg或订单量kg的值，报关名称实际为空）
        if col_e.isdigit():
            # 检查规格kg位置（index 5）是否有内容
            col_f = parts[5] if len(parts) > 5 else ''
            if col_f.isdigit():
                # 规格kg位置也是纯数字，说明数据对齐，不需要修复
                return parts
            # 规格kg位置为空，需要在报关名称位置插入空值
            return parts[:4] + [''] + parts[4:]

        return parts

    def _collapse_extra_fields(self, parts: list) -> list:
        """
        当Split后字段数量超过23时，说明某些字段内容包含了分隔符（Tab或换行）。
        策略：
        - 索引0-8（业务员~有无样品）：逐行对应
        - 索引9开始是"订单要求"，需要合并直到遇到"真正的新字段"
        - 从索引10开始（交货日期~规格异常）：逐行对应
        """
        if len(parts) <= len(self.COLUMNS):
            return parts

        result = []
        i = 0
        while i < len(parts):
            if i < 9:
                # 索引0-8：正常对应列
                result.append(parts[i])
                i += 1
            elif i == 9:
                # 索引9：订单要求，需要合并后续内容直到真正的新字段开始
                merged = parts[i]
                j = i + 1
                while j < len(parts):
                    next_part = parts[j].strip()
                    # 遇到真正的新字段（从索引10开始判断）
                    if j >= 10 and self._looks_like_real_field_start(next_part):
                        break
                    # 遇到空行且下一行是真正的新字段，说明当前合并结束
                    if not next_part and j + 1 < len(parts) and self._looks_like_real_field_start(parts[j + 1].strip()):
                        break
                    merged += '\n' + next_part
                    j += 1
                result.append(merged.strip())
                i = j
            else:
                # 索引10之后：逐行对应
                result.append(parts[i])
                i += 1

        return result

    def _reassemble_embedded_newlines(self, lines: list, parts: list) -> list:
        """
        当Tab分割后字段数量不足23时，说明"订单要求"字段内发生了嵌入换行。
        策略：从lines中找出哪些行属于"订单要求"，合并后确保parts总数达到23。

        lines: 按换行符分割后的所有行
        parts: 按Tab分割后的初始字段（部分字段缺失）
        """
        if len(parts) >= len(self.COLUMNS):
            return parts[:len(self.COLUMNS)]

        result = []

        for i, col in enumerate(self.COLUMNS):
            if i < len(parts) and parts[i]:
                # 已经有值的字段，直接使用
                result.append(parts[i])
            elif col == '订单要求':
                # 订单要求字段可能分散在多行，需要合并
                # 从lines中找到对应位置的内容进行合并
                # 策略：收集从当前i位置开始，到下一个"真正的新字段"为止的所有内容
                merged = parts[i] if i < len(parts) else ''
                if i < len(lines):
                    for line in lines[i:]:
                        stripped = line.strip()
                        if not stripped:
                            continue
                        # 如果这行是真正的新字段开始，停止合并
                        if self._looks_like_real_field_start(stripped):
                            break
                        # 否则这行属于订单要求
                        if stripped and stripped not in result:
                            if merged:
                                merged += '\n' + stripped
                            else:
                                merged = stripped
                result.append(merged)
            elif col == '生产交期':
                # 找到生产交期的实际位置（可能在lines中更后面）
                found = ''
                for line in lines:
                    stripped = line.strip()
                    if stripped and not self._looks_like_real_field_start(stripped):
                        if re.match(r'^\d+月\d+日', stripped) or re.match(r'^\d{4}年', stripped):
                            found = stripped
                            break
                result.append(found)
            else:
                result.append('')

        return result

    def _looks_like_real_field_start(self, text: str) -> bool:
        """
        判断一段文本是否是"真正的新字段的开始"。
        注意：以下关键词虽然在正常字段列表中出现，但在此数据中它们出现在
        "订单要求"的多行内容里，所以不应该被视为"真正的新字段开始"。
        - 出货日期：2026年6月12日 -> 属于订单要求
        - 出货方式：40尺整柜出货 -> 属于订单要求
        - 注意事项： -> 属于订单要求
        """
        import re
        text = text.strip()
        if not text:
            return False

        real_markers = [
            '已审核', '待审核',
            '正常单',
            '外贸部',
            'HT20',
            '宏昊抬头',
            '刘洁婷',
            '下单',
            '出货渠道',
        ]

        for marker in real_markers:
            # 使用精确匹配（字段完整等于标记）而不是包含检查
            if text == marker:
                return True

        # 匹配完整日期格式：2026年5月12日 20:38
        if re.match(r'^\d{4}\s*年\s*\d+月', text) and '：' not in text:
            return True

        # 匹配短日期格式：如 6月12日、6月12日 20:38
        if re.match(r'^\d{1,2}\s*月\s*\d{1,2}\s*日', text):
            return True

        # 匹配纯数字（数量、编号）- 但要排除业务员的客户编号格式
        if re.match(r'^\d+$', text) and len(text) <= 6:
            return True

        return False

    def get_pi号(self) -> str:
        """获取PI号（订单号）"""
        return self.data.get('订单号', '')

    def get_internal_code(self) -> str:
        """获取内部编号"""
        return self.data.get('内部编号', '')

    def get_all_data(self) -> dict:
        """获取所有解析数据"""
        return self.data.copy()

    def validate(self) -> tuple:
        """
        验证必填字段
        :return: (is_valid, error_message)
        """
        if not self.get_pi号():
            return False, "订单号（PI号）不能为空"

        if not self.get_internal_code():
            return False, "内部编号不能为空"

        return True, ""
