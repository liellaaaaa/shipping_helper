"""
MSDS解析器
解析DOC格式的MSDS文档，提取第1/3/9部分的字段信息
处理第3部分成分表的不规则换行和合并单元格
"""

import os
import re
from typing import List, Optional

from .data_models import MSDSInfo, Component, PhysicalChemical


class MSDSParser:
    """
    MSDS解析器
    挑战：第3部分成分表有不规则换行和合并单元格
    技术：使用 pywin32 Word COM 读取 .doc 格式
    """

    def __init__(self):
        self.word_app = None
        self.document = None

    def parse(self, filepath: str) -> Optional[MSDSInfo]:
        """
        解析MSDS文档

        Args:
            filepath: MSDS文件路径（.doc格式）

        Returns:
            MSDSInfo对象，解析失败返回None
        """
        if not os.path.exists(filepath):
            return None

        if not filepath.lower().endswith('.doc'):
            # python-docx 不支持 .doc，需要用 Word COM
            return None

        try:
            import win32com.client
        except ImportError:
            print("需要安装pywin32库: pip install pywin32")
            return None

        msds = MSDSInfo()

        try:
            # 启动Word应用
            self.word_app = win32com.client.Dispatch("Word.Application")
            self.word_app.Visible = False

            # 打开文档
            self.document = self.word_app.Documents.Open(filepath)

            # 提取文本内容
            full_text = self.document.Content.Text

            # 提取各部分
            msds.section1_identifiers = self._extract_section1(full_text)
            msds.section3_raw = self._extract_section3_raw()
            msds.section9_raw = self._extract_section9_raw(full_text)

            msds.product_name = self._parse_product_name(msds.section1_identifiers)
            msds.components = self._parse_components(msds.section3_raw)
            msds.physical_chemical = self._parse_physical_chemical(msds.section9_raw)

            return msds

        except Exception as e:
            print(f"MSDS解析失败 {filepath}: {e}")
            return None

        finally:
            self._close_document()

    def _close_document(self):
        """关闭Word文档"""
        if self.document:
            try:
                self.document.Close(False)
            except:
                pass
            self.document = None
        if self.word_app:
            try:
                self.word_app.Quit()
            except:
                pass
            self.word_app = None

    def _extract_section1(self, full_text: str) -> str:
        """提取第1部分：标识"""
        # 第1部分通常包含产品名称、供应商等信息
        section1 = ""
        lines = full_text.split('\n')

        in_section1 = False
        section_count = 0

        for line in lines:
            line = line.strip()

            # 检测第1部分开始
            if re.match(r'^第\s*1\s*部分|^\s*1\s*\.?\s*标识', line, re.IGNORECASE):
                in_section1 = True
                section_count = 1
                continue

            # 检测其他部分开始
            if in_section1 and re.match(r'^第\s*[2-9]\s*部分|^\s*[2-9]\s*\.', line):
                break

            if in_section1:
                section1 += line + "\n"

        return section1.strip()

    def _extract_section3_raw(self) -> str:
        """提取第3部分原始内容（包含表格）"""
        section3 = ""

        if not self.document:
            return section3

        try:
            # 遍历文档中的表格
            for table in self.document.Tables:
                table_text = ""

                for row in table.Rows:
                    row_text = ""
                    for cell in row.Cells:
                        cell_text = cell.Range.Text.strip()
                        row_text += cell_text + "\t"
                    table_text += row_text + "\n"

                # 如果表格包含组分/CAS/含量关键词，认为是成分表
                if any(k in table_text for k in ["组分", "CAS", "成分", "含量"]):
                    section3 = table_text
                    break

        except Exception as e:
            print(f"表格提取失败: {e}")

        return section3

    def _extract_section9_raw(self, full_text: str) -> str:
        """提取第9部分：理化特性"""
        section9 = ""
        lines = full_text.split('\n')

        in_section9 = False

        for line in lines:
            line = line.strip()

            # 检测第9部分开始
            if re.match(r'^第\s*9\s*部分|^\s*9\s*\.?\s*理化', line, re.IGNORECASE):
                in_section9 = True
                continue

            # 检测第10部分开始
            if in_section9 and re.match(r'^第\s*1[0-9]\s*部分|^\s*1[0-9]\s*\.', line):
                break

            if in_section9:
                section9 += line + "\n"

        return section9.strip()

    def _parse_product_name(self, section1_text: str) -> str:
        """从第1部分提取产品名称"""
        if not section1_text:
            return ""

        lines = section1_text.split('\n')

        for line in lines:
            # 匹配 "产品名称：" 或 "产品名称：" 后面跟的内容
            match = re.search(r'产品名称[：:]\s*(.+)', line)
            if match:
                return match.group(1).strip()

            # 也可能直接是产品名称
            if len(line) > 2 and len(line) < 100:
                if not any(k in line for k in ["第1部分", "标识", "供应商", "电话", "传真", "应急电话"]):
                    if line.strip() and not line.strip().startswith("1."):
                        return line.strip()

        return ""

    def _parse_components(self, section3_text: str) -> List[Component]:
        """
        解析第3部分成分表
        处理不规则换行和合并单元格的情况
        """
        components = []

        if not section3_text:
            return components

        lines = section3_text.split('\n')
        header_row_idx = -1

        # 查找表头行
        for i, line in enumerate(lines):
            if any(k in line for k in ["组分", "成分", "Component"]):
                header_row_idx = i
                break

        if header_row_idx == -1:
            # 没有找到表头，尝试模糊匹配
            header_row_idx = 0

        # 确定列索引（组分/CAS/含量）
        col_indices = {"name": -1, "cas": -1, "conc": -1}

        if header_row_idx < len(lines):
            header_parts = lines[header_row_idx].split('\t')

            for i, part in enumerate(header_parts):
                part_lower = part.lower().strip()
                if '组分' in part or '成分' in part or 'component' in part_lower:
                    col_indices["name"] = i
                elif 'cas' in part_lower or 'cas号' in part:
                    col_indices["cas"] = i
                elif '含量' in part or '浓度' in part or 'concentration' in part_lower:
                    col_indices["conc"] = i

        # 解析数据行
        for i in range(header_row_idx + 1, len(lines)):
            line = lines[i].strip()
            if not line:
                continue

            parts = line.split('\t')

            # 过滤非数据行（太少的列或表头）
            if len(parts) < 2:
                continue

            # 提取各列数据
            name = ""
            cas = ""
            conc = ""

            if col_indices["name"] >= 0 and col_indices["name"] < len(parts):
                name = parts[col_indices["name"]].strip()

            if col_indices["cas"] >= 0 and col_indices["cas"] < len(parts):
                cas = parts[col_indices["cas"]].strip()

            if col_indices["conc"] >= 0 and col_indices["conc"] < len(parts):
                conc = parts[col_indices["conc"]].strip()

            # 清理数据
            name = self._clean_component_name(name)
            cas = self._clean_cas(cas)
            conc = self._clean_concentration(conc)

            # 验证数据有效性
            if name and len(name) > 1:
                component = Component(name=name, cas_no=cas, concentration=conc)
                components.append(component)

        return components

    def _clean_component_name(self, name: str) -> str:
        """清理成分名称"""
        if not name:
            return ""

        # 移除序号、特殊字符
        name = re.sub(r'^[\d\.\、\、]+', '', name)
        name = name.strip()

        # 过滤无效内容
        if len(name) < 2 or name in ["组分", "成分", "CAS", "含量", "浓度"]:
            return ""

        return name

    def _clean_cas(self, cas: str) -> str:
        """清理CAS号"""
        if not cas:
            return ""

        # CAS号格式通常是数字-数字-数字
        cas = cas.strip()
        match = re.search(r'\d{2,7}-\d{2}-\d', cas)
        if match:
            return match.group(0)

        return cas

    def _clean_concentration(self, conc: str) -> str:
        """清理浓度/含量"""
        if not conc:
            return ""

        conc = conc.strip()

        # 移除"约"、"≈"等
        conc = re.sub(r'[≈约]', '', conc)

        # 确保有%
        if not conc.endswith('%') and not re.search(r'\d+%', conc):
            conc += '%'

        return conc

    def _parse_physical_chemical(self, section9_text: str) -> Optional[PhysicalChemical]:
        """解析第9部分理化特性"""
        if not section9_text:
            return None

        pc = PhysicalChemical()

        # 定义各字段的模式
        patterns = {
            'appearance': [r'外观与性状[：:]\s*(.+)', r'外观[：:]\s*(.+)'],
            'ionic': [r'离子性[：:]\s*(.+)'],
            'ph': [r'pH[值]*[：:]\s*([0-9.]+)', r'PH[值]*[：:]\s*([0-9.]+)'],
            'melting_point': [r'熔点[：:]\s*(.+)'],
            'boiling_point': [r'沸点[：:]\s*(.+)', r'沸点范围[：:]\s*(.+)'],
            'density': [r'相对密度[：:]\s*(.+)', r'密度[：:]\s*(.+)'],
            'vapor_pressure': [r'蒸汽压[：:]\s*(.+)'],
            'flash_point': [r'闪点[：:]\s*(.+)'],
            'autoignition': [r'自燃点[：:]\s*(.+)'],
            'explosion_lower': [r'爆炸下限[：:]\s*(.+)'],
            'explosion_upper': [r'爆炸上限[：:]\s*(.+)'],
            'solubility': [r'溶解性[：:]\s*(.+)', r'溶解度[：:]\s*(.+)'],
        }

        for field, field_patterns in patterns.items():
            for pattern in field_patterns:
                match = re.search(pattern, section9_text, re.IGNORECASE)
                if match:
                    value = match.group(1).strip()
                    setattr(pc, field, value)
                    break

        return pc


if __name__ == "__main__":
    # 测试代码
    parser = MSDSParser()

    # 测试MSDS目录
    msds_dir = r"c:\Users\windows\Desktop\shipping_helper\shipping_helper\02.订舱出货\源文件\MSDS\2026年MSDS（陆运上化院）"

    if os.path.exists(msds_dir):
        files = [f for f in os.listdir(msds_dir) if f.endswith('.doc')]
        if files:
            test_file = os.path.join(msds_dir, files[0])
            print(f"测试解析: {files[0]}")

            msds = parser.parse(test_file)
            if msds:
                print(f"产品名称: {msds.product_name}")
                print(f"成分数量: {len(msds.components)}")
                for comp in msds.components[:3]:
                    print(f"  - {comp.name} | {comp.cas_no} | {comp.concentration}")
            else:
                print("解析失败")