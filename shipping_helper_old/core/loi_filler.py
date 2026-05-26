# -*- coding: utf-8 -*-
"""
LOI模板填写器 - 填写docx格式的LOI保函模板
支持非危险品保函和液体保函两种模板
"""

import os
import re
from typing import Dict


class LOIFiller:
    """LOI模板填写器"""

    # LOI模板中的占位符格式
    PLACEHOLDER_PATTERN = re.compile(r'\{\{(\w+)\}\}')

    # 字段映射（用于将提取的字段映射到模板占位符）
    FIELD_MAPPING = {
        '样品名称中文': ['样品名称中文', '品名中文', '中文品名', 'cargo_name_cn'],
        '样品名称英文': ['样品名称英文', '品名英文', '英文品名', 'cargo_name_en'],
        '报告编号': ['报告编号', '鉴定书编号', 'report_no', 'certificate_no'],
        '外观性状': ['外观性状', '产品外观与性状', 'appearance', '性状'],
        '主要成分': ['主要成分', '成分', 'components'],
        '委托单位': ['委托单位', 'client', 'company'],
        '运输类型': ['运输类型', 'transport_type'],
    }

    def __init__(self):
        self.template_path = None
        self.template_type = None  # 'non_hazardous' or 'liquid'

    def load_template(self, template_path: str) -> bool:
        """
        加载LOI模板文件

        Args:
            template_path: 模板文件路径

        Returns:
            是否加载成功
        """
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"模板文件不存在: {template_path}")

        self.template_path = template_path

        # 判断模板类型
        filename = os.path.basename(template_path).lower()
        if '液体' in template_path or 'liquid' in filename:
            self.template_type = 'liquid'
        elif '非危险品' in template_path or 'non' in filename:
            self.template_type = 'non_hazardous'
        else:
            self.template_type = 'unknown'

        return True

    def fill(self, data: Dict[str, str], output_path: str) -> bool:
        """
        填写模板

        Args:
            data: 字段数据字典
            output_path: 输出文件路径

        Returns:
            是否成功
        """
        from docx import Document

        if not self.template_path:
            raise RuntimeError("未加载模板文件")

        # 打开模板
        doc = Document(self.template_path)

        # 替换段落中的占位符
        for para in doc.paragraphs:
            self._replace_in_text(para, data)

        # 替换表格中的占位符
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        self._replace_in_text(para, data)

        # 保存
        doc.save(output_path)
        return True

    def _replace_in_text(self, para, data: Dict[str, str]):
        """替换段落或表格中的占位符"""
        if not para.text:
            return

        text = para.text

        # 查找所有占位符
        matches = list(self.PLACEHOLDER_PATTERN.finditer(text))

        if not matches:
            # 如果没有占位符，尝试在表格中匹配标签和值的位置
            # 例如："中文品名：                                  英文品名："
            return

        # 替换每个占位符
        for match in matches:
            placeholder = match.group(1)
            value = self._get_field_value(placeholder, data)
            if value:
                text = text.replace(match.group(0), value)

        # 设置新的文本（保留格式）
        if matches:
            for run in para.runs:
                if run.text:
                    for m in matches:
                        if m.group(0) in run.text:
                            run.text = run.text.replace(m.group(0), self._get_field_value(m.group(1), data) or '')

    def _get_field_value(self, key: str, data: Dict[str, str]) -> str:
        """获取字段值，支持映射"""
        # 直接查找
        if key in data:
            return data[key]

        # 尝试映射
        for target_key, aliases in self.FIELD_MAPPING.items():
            if key in aliases or key == target_key:
                if target_key in data:
                    return data[target_key]
                for alias in aliases:
                    if alias in data:
                        return data[alias]

        return ''

    def fill_manual(self, output_path: str) -> bool:
        """
        手动填写模式 - 返回模板内容供用户编辑

        Args:
            output_path: 输出文件路径

        Returns:
            是否成功
        """
        if not self.template_path:
            raise RuntimeError("未加载模板文件")

        # 直接复制模板到输出路径，不做自动替换
        import shutil
        shutil.copy2(self.template_path, output_path)
        return True

    def get_template_info(self) -> Dict[str, str]:
        """获取模板信息"""
        return {
            'path': self.template_path,
            'type': self.template_type,
        }


class LOITemplateEditor:
    """LOI模板编辑器 - 用于交互式填写LOI模板"""

    def __init__(self):
        self.filler = LOIFiller()
        self.current_output_path = None

    def open_template(self, template_path: str):
        """打开模板"""
        self.filler.load_template(template_path)
        return self.filler.get_template_info()

    def prepare_for_edit(self, output_path: str) -> bool:
        """准备编辑 - 复制模板到输出路径"""
        self.current_output_path = output_path
        return self.filler.fill_manual(output_path)

    def get_editable_fields(self) -> list:
        """获取可编辑的字段列表"""
        from docx import Document

        if not self.filler.template_path:
            return []

        doc = Document(self.filler.template_path)
        fields = []

        # 扫描段落和表格中的占位符
        for para in doc.paragraphs:
            matches = self.PLACEHOLDER_PATTERN.findall(para.text)
            fields.extend(matches)

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        matches = self.PLACEHOLDER_PATTERN.findall(para.text)
                        fields.extend(matches)

        return list(set(fields))