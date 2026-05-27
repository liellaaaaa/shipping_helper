"""
运输鉴定报告解析器
解析PDF和DOCX格式的运输鉴定报告，提取货物信息
"""

import os
import re
from typing import List, Optional

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

from .data_models import CargoInfo


class ReportParser:
    """
    运输鉴定报告解析器
    支持：PDF（海运/空运）、DOCX（登记表）
    """

    def __init__(self):
        self.fields = {}

    def parse(self, filepath: str) -> Optional[CargoInfo]:
        """
        解析单个报告文件

        Args:
            filepath: 文件路径

        Returns:
            CargoInfo对象，解析失败返回None
        """
        if not os.path.exists(filepath):
            return None

        ext = os.path.splitext(filepath)[1].lower()

        if ext == '.pdf':
            return self._parse_pdf(filepath)
        elif ext in ['.docx']:
            return self._parse_docx(filepath)
        elif ext in ['.xls', '.xlsx']:
            return self._parse_excel(filepath)
        else:
            return None

    def _parse_excel(self, filepath: str) -> Optional[CargoInfo]:
        """解析Excel格式的运输鉴定报告"""
        try:
            import xlrd
            wb = xlrd.open_workbook(filepath, encoding_override='utf-8')
            sh = wb.sheet_by_index(0)

            full_text = ""
            for r in range(sh.nrows):
                for c in range(sh.ncols):
                    val = sh.cell_value(r, c)
                    if val:
                        full_text += str(val) + "\n"

            return self._extract_cargo_info(full_text, filepath)

        except Exception as e:
            print(f"Excel解析失败 {filepath}: {e}")
            return None

    def _parse_pdf(self, filepath: str) -> Optional[CargoInfo]:
        """使用PyMuPDF解析PDF"""
        if fitz is None:
            raise ImportError("PyMuPDF (fitz) 未安装，请运行: pip install pymupdf")

        try:
            doc = fitz.open(filepath)
            full_text = ""

            for page in doc:
                full_text += page.get_text() + "\n"

            doc.close()

            return self._extract_cargo_info(full_text, filepath)

        except Exception as e:
            print(f"PDF解析失败 {filepath}: {e}")
            return None

    def _parse_docx(self, filepath: str) -> Optional[CargoInfo]:
        """解析DOCX格式（委托服务登记表）"""
        try:
            from docx import Document

            doc = Document(filepath)
            full_text = ""

            # 提取段落文本
            for para in doc.paragraphs:
                full_text += para.text + "\n"

            # 提取表格文本
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        full_text += cell.text + " "
                full_text += "\n"

            return self._extract_cargo_info(full_text, filepath)

        except Exception as e:
            print(f"DOCX解析失败 {filepath}: {e}")
            return None

    def _extract_cargo_info(self, text: str, filepath: str = "") -> Optional[CargoInfo]:
        """从文本中提取货物信息"""
        cargo = CargoInfo()

        # 规范化文本：处理换行和多余空格
        lines = text.split('\n')
        normalized_lines = []
        for line in lines:
            line = line.strip()
            if line:
                normalized_lines.append(line)

        n = len(normalized_lines)
        i = 0

        while i < n:
            line = normalized_lines[i]

            # 样品名称（中文）
            if '样品名称' in line and 'Name' not in line:
                for j in range(i + 1, min(i + 4, n)):
                    content = normalized_lines[j]
                    if 'Name of Goods' in content or not content:
                        continue
                    if self._contains_chinese(content):
                        cargo.product_name_cn = content.strip()
                        break

            # 样品名称（英文）
            if 'Name of Goods' in line or (i > 0 and 'Name' in normalized_lines[i-1] and 'Name of Goods' in line):
                for j in range(i + 1, min(i + 4, n)):
                    content = normalized_lines[j]
                    if content and not self._contains_chinese(content):
                        if self._contains_english(content):
                            cargo.product_name_en = content.strip()
                            break

            # 报告编号 - 格式：No.XXXXXXXX
            match = re.search(r'No\.?\s*([A-Z0-9]{10,})', line)
            if match:
                cargo.report_no = match.group(1).strip()

            # 样品编号 - 格式：WXST2513478-036
            if 'Sample No' in line or '样品编号' in line:
                match = re.search(r'([A-Z0-9]{8,}-?\d*)', line)
                if match:
                    cargo.sample_no = match.group(1).strip()
                else:
                    for j in range(i + 1, min(i + 3, n)):
                        content = normalized_lines[j].strip()
                        if content:
                            match = re.search(r'([A-Z0-9]{8,}-?\d*)', content)
                            if match:
                                cargo.sample_no = match.group(1).strip()
                                break

            # 样品描述（外观性状）- 可能跨行
            if '样品描述' in line:
                desc_lines = []
                skip_next = False
                for j in range(i + 1, min(i + 6, n)):
                    content = normalized_lines[j].strip()
                    if 'Appearance' in content:
                        skip_next = True
                        continue
                    if skip_next:
                        skip_next = False
                        continue
                    if '鉴定结论' in content or 'Conclusion' in content:
                        break
                    if content:
                        desc_lines.append(content)
                    if len(desc_lines) >= 3:
                        break

                if desc_lines:
                    chinese = ''
                    english = ''
                    for line_content in desc_lines:
                        if self._contains_chinese(line_content) and '鉴定结论' not in line_content:
                            chinese += line_content + ' '
                        elif self._contains_english(line_content):
                            english += line_content + ' '

                    if chinese:
                        cargo.appearance_cn = chinese.strip()
                    if english:
                        cargo.appearance_en = english.strip()

            # 委托单位
            if '委托单位' in line and 'Applicant' not in line:
                for j in range(i + 1, min(i + 4, n)):
                    content = normalized_lines[j].strip()
                    if content and not re.match(r'^[A-Za-z\s\.,]+$', content):
                        cargo.consigner = content
                        break

            # 生产单位
            if '生产单位' in line and 'Manufacturer' not in line:
                for j in range(i + 1, min(i + 4, n)):
                    content = normalized_lines[j].strip()
                    if content and not re.match(r'^[A-Za-z\s\.,]+$', content):
                        cargo.producer = content
                        break

            # 鉴定结论 - 查找"鉴定结论"或"Conclusion"标签后的内容
            # 精确匹配：只匹配单独的"鉴定结论"行或"Conclusion"行，不匹配包含这两个词的混合文本
            is_conclusion_label = (line == '鉴定结论' or line == 'Conclusion' or
                                   '鉴定结论' in line and 'Conclusion' in line)
            if is_conclusion_label:
                # 检查后续几行找结论内容
                for j in range(i + 1, min(i + 4, n)):
                    content = normalized_lines[j].strip()
                    # 跳过空行和纯标签
                    if not content or content in ['Conclusion', '鉴定结论']:
                        continue
                    # 跳过纯英文标签行或太短的行
                    if re.match(r'^[A-Za-z\s]+$', content) or len(content) < 5:
                        continue
                    # 跳过危险评估项目的标题（以数字开头）
                    if re.match(r'^\d+\.', content):
                        continue
                    # 跳过PDF声明相关的段落（通常很长）
                    if len(content) > 100:
                        continue
                    # 找到结论内容
                    if content:
                        cargo.conclusion = content
                        break

            i += 1

        # 提取运输类型
        text_lower = text.lower()
        if 'imo imdg' in text_lower or 'imdg code' in text_lower:
            cargo.transport_type = '海运'
        elif 'iata dgr' in text_lower or '空运' in text:
            cargo.transport_type = '空运'
        elif 'jt/t' in text_lower or '道路运输' in text:
            cargo.transport_type = '陆运'

        # 判断是否危险品
        if cargo.conclusion:
            dangerous_keywords = ["危险", "dangerous", "hazard", "class"]
            cargo.is_dangerous = any(k in cargo.conclusion.lower() for k in dangerous_keywords)

        return cargo

    def _contains_chinese(self, text: str) -> bool:
        """检查是否包含中文"""
        return bool(re.search(r'[一-鿿]', text))

    def _contains_english(self, text: str) -> bool:
        """检查是否包含英文"""
        return bool(re.search(r'[a-zA-Z]', text))

    def parse_directory(self, directory: str) -> List[CargoInfo]:
        """
        批量解析目录下的所有报告

        Args:
            directory: 目录路径

        Returns:
            解析的CargoInfo列表
        """
        results = []

        if not os.path.exists(directory):
            return results

        for filename in os.listdir(directory):
            filepath = os.path.join(directory, filename)
            if os.path.isfile(filepath):
                ext = os.path.splitext(filename)[1].lower()
                if ext in ['.pdf', '.docx']:
                    cargo = self.parse(filepath)
                    if cargo:
                        results.append(cargo)

        return results


if __name__ == "__main__":
    # 测试代码
    parser = ReportParser()

    # 测试空运目录
    air_dir = r"c:\Users\windows\Desktop\shipping_helper\shipping_helper\02.订舱出货\源文件\2026中观测空运运输鉴定报告"

    if os.path.exists(air_dir):
        results = parser.parse_directory(air_dir)
        print(f"解析 {len(results)} 个空运鉴定报告:")
        for cargo in results:
            print(f"  - {cargo.product_name_cn} / {cargo.product_name_en}")
            print(f"    报告编号: {cargo.report_no}, 运输类型: {cargo.transport_type}")

    # 测试海运目录
    sea_dir = r"c:\Users\windows\Desktop\shipping_helper\shipping_helper\02.订舱出货\源文件\2026年中广测海运鉴定报告"

    if os.path.exists(sea_dir):
        results = parser.parse_directory(sea_dir)
        print(f"\n解析 {len(results)} 个海运鉴定报告:")
        for cargo in results[:3]:  # 只显示前3个
            print(f"  - {cargo.product_name_cn} / {cargo.product_name_en}")