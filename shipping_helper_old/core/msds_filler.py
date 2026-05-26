# -*- coding: utf-8 -*-
"""
MSDS模板填写器 - 填写doc格式的MSDS模板
使用pywin32 Word COM自动化保持原有格式
"""

import os
import re
import shutil
from typing import Dict, Optional


class MSDSFiller:
    """MSDS模板填写器 - 使用Word COM自动化"""

    # MSDS模板中的占位符格式
    PLACEHOLDER_PATTERN = re.compile(r'\{\{\{(\w+)\}\}\}')

    def __init__(self):
        self.template_path = None
        self.word_app = None
        self.document = None

    def load_template(self, template_path: str) -> bool:
        """
        加载MSDS模板文件

        Args:
            template_path: 模板文件路径

        Returns:
            是否加载成功
        """
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"模板文件不存在: {template_path}")

        self.template_path = template_path
        return True

    def fill(self, data: Dict[str, str], output_path: str) -> bool:
        """
        填写MSDS模板（自动模式）

        Args:
            data: 字段数据字典
            output_path: 输出文件路径

        Returns:
            是否成功
        """
        try:
            import win32com.client
        except ImportError:
            raise ImportError("需要安装pywin32库: pip install pywin32")

        # 启动Word应用
        self.word_app = win32com.client.Dispatch("Word.Application")
        self.word_app.Visible = False

        try:
            # 打开模板
            self.document = self.word_app.Documents.Open(self.template_path)

            # 替换占位符
            for key, value in data.items():
                placeholder = f"{{{{{key}}}}}"

                # 查找替换
                find = self.document.Find
                find.Text = placeholder
                find.Replacement.Text = value
                find.Execute(Replace=2)  # wdReplaceAll = 2

            # 保存到新文件
            self.document.SaveAs(output_path)
            return True

        finally:
            # 关闭文档和Word
            if self.document:
                self.document.Close(False)
            if self.word_app:
                self.word_app.Quit()
            self.word_app = None
            self.document = None

    def fill_manual(self, output_path: str) -> bool:
        """
        手动填写模式 - 复制模板到输出路径供用户编辑

        Args:
            output_path: 输出文件路径

        Returns:
            是否成功
        """
        if not self.template_path:
            raise RuntimeError("未加载模板文件")

        # 直接复制模板到输出路径
        shutil.copy2(self.template_path, output_path)
        return True

    def open_for_edit(self, output_path: str) -> bool:
        """
        打开模板准备编辑（手动模式）

        Args:
            output_path: 输出文件路径

        Returns:
            是否成功
        """
        self.fill_manual(output_path)

        try:
            import win32com.client
        except ImportError:
            raise ImportError("需要安装pywin32库: pip install pywin32")

        # 启动Word应用并打开文件
        self.word_app = win32com.client.Dispatch("Word.Application")
        self.word_app.Visible = True

        self.document = self.word_app.Documents.Open(output_path)
        return True

    def close(self):
        """关闭Word文档"""
        if self.document:
            self.document.Close(False)
            self.document = None
        if self.word_app:
            self.word_app.Quit()
            self.word_app = None

    def get_template_info(self) -> Dict[str, str]:
        """获取模板信息"""
        return {
            'path': self.template_path,
            'type': 'msds',
        }

    def get_editable_fields(self) -> list:
        """获取可编辑的字段列表"""
        if not self.template_path or not os.path.exists(self.template_path):
            return []

        try:
            import win32com.client
        except ImportError:
            return []

        word_app = win32com.client.Dispatch("Word.Application")
        word_app.Visible = False

        try:
            doc = word_app.Documents.Open(self.template_path)
            fields = set()

            # 遍历文档查找占位符
            for para in doc.Content.Paragraphs:
                matches = self.PLACEHOLDER_PATTERN.findall(para.Range.Text)
                fields.update(matches)

            doc.Close(False)
            return list(fields)

        finally:
            word_app.Quit()

    def extract_fields_from_document(self, doc_path: str) -> Dict[str, str]:
        """
        从MSDS文档提取字段值（用于预填充参考）

        Args:
            doc_path: 文档路径

        Returns:
            字段字典
        """
        try:
            import win32com.client
        except ImportError:
            return {}

        word_app = win32com.client.Dispatch("Word.Application")
        word_app.Visible = False

        try:
            doc = word_app.Documents.Open(doc_path)
            fields = {}

            # 提取所有文本内容进行分析
            text = doc.Content.Text

            # 尝试提取一些常见字段
            patterns = {
                '产品名称': r'产品名称[：:]\s*(.+)',
                '英文名称': r'英文名称[：:]\s*(.+)',
                'CAS号': r'CAS[号No\.]*[：:]\s*([0-9-]+)',
                '外观': r'外观[：:]\s*(.+)',
                'pH值': r'pH[值]*[：:]\s*([0-9.]+)',
            }

            for field, pattern in patterns.items():
                match = re.search(pattern, text, re.IGNORECASE)
                if match:
                    fields[field] = match.group(1).strip()

            doc.Close(False)
            return fields

        finally:
            word_app.Quit()


class MSDSTemplateEditor:
    """MSDS模板编辑器 - 用于交互式填写MSDS模板"""

    def __init__(self):
        self.filler = MSDSFiller()
        self.current_output_path = None

    def open_template(self, template_path: str):
        """打开模板"""
        self.filler.load_template(template_path)
        return self.filler.get_template_info()

    def prepare_for_edit(self, output_path: str) -> bool:
        """准备编辑 - 复制模板到输出路径"""
        self.current_output_path = output_path
        return self.filler.fill_manual(output_path)