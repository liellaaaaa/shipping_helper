"""
MSDS生成器
生成中文版和英文版MSDS文档
"""

import os
from typing import Optional, Dict, List

from .data_models import ShipmentData, Component, PhysicalChemical


class MSDSGenerator:
    """
    MSDS生成器
    基于提取的MSDS信息和出口商品编码数据生成新版MSDS
    """

    def __init__(self):
        pass

    def generate_chinese(self, shipment: ShipmentData, output_dir: str = None) -> str:
        """
        生成中文版MSDS

        Args:
            shipment: 完整货物信息
            output_dir: 输出目录

        Returns:
            生成的MSDS文件路径
        """
        if output_dir is None:
            output_dir = os.path.join(os.path.expanduser("~"), "Desktop")

        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        # 准备数据
        data = self._prepare_chinese_data(shipment)

        # 生成文件名
        product_name = shipment.product_name_cn or "MSDS"
        filename = f"中文MSDS-{product_name}.docx"
        output_path = os.path.join(output_dir, filename)

        # 创建文档
        self._create_msds_document(data, output_path, is_english=False)

        return output_path

    def generate_english(self, shipment: ShipmentData, output_dir: str = None) -> str:
        """
        生成英文版MSDS

        Args:
            shipment: 完整货物信息
            output_dir: 输出目录

        Returns:
            生成的MSDS文件路径
        """
        if output_dir is None:
            output_dir = os.path.join(os.path.expanduser("~"), "Desktop")

        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        # 准备数据
        data = self._prepare_english_data(shipment)

        # 生成文件名
        product_name = shipment.product_name_en or "MSDS"
        filename = f"MSDS-{product_name}.docx"
        output_path = os.path.join(output_dir, filename)

        # 创建文档
        self._create_msds_document(data, output_path, is_english=True)

        return output_path

    def generate_both(self, shipment: ShipmentData, output_dir: str = None) -> tuple:
        """
        同时生成中英文版MSDS

        Returns:
            (中文版路径, 英文版路径)
        """
        cn_path = self.generate_chinese(shipment, output_dir)
        en_path = self.generate_english(shipment, output_dir)
        return cn_path, en_path

    def _prepare_chinese_data(self, shipment: ShipmentData) -> Dict:
        """准备中文MSDS数据"""
        data = {
            "product_name": "",
            "section1": {},
            "section3": [],
            "section9": {},
        }

        # 产品名称
        if shipment.msds and shipment.msds.product_name:
            data["product_name"] = shipment.msds.product_name
        elif shipment.product_name_cn:
            data["product_name"] = shipment.product_name_cn
        elif shipment.cargo:
            data["product_name"] = shipment.cargo.product_name_cn

        # 第1部分：标识
        data["section1"]["product_name"] = data["product_name"]
        data["section1"]["cas_no"] = ""
        data["section1"][" supplier"] = shipment.shipper or ""
        data["section1"]["emergency_phone"] = ""

        # 第3部分：成分
        if shipment.msds and shipment.msds.components:
            for comp in shipment.msds.components:
                data["section3"].append({
                    "name": comp.name,
                    "cas_no": comp.cas_no,
                    "concentration": comp.concentration,
                })
        elif shipment.export_code and shipment.export_code.composition:
            # 从出口商品编码获取成分
            data["section3"].append({
                "name": shipment.export_code.composition,
                "cas_no": "",
                "concentration": "100%",
            })

        # 第9部分：理化特性
        if shipment.msds and shipment.msds.physical_chemical:
            pc = shipment.msds.physical_chemical
            data["section9"] = {
                "appearance": pc.appearance or (shipment.cargo.appearance_cn if shipment.cargo else ""),
                "ph": pc.ph or "",
                "melting_point": pc.melting_point or "",
                "boiling_point": pc.boiling_point or "",
                "density": pc.density or "",
                "flash_point": pc.flash_point or "",
                "solubility": pc.solubility or "",
            }
        else:
            data["section9"] = {
                "appearance": shipment.cargo.appearance_cn if shipment.cargo else "",
                "ph": "",
                "melting_point": "",
                "boiling_point": "",
                "density": "",
                "flash_point": "",
                "solubility": "",
            }

        return data

    def _prepare_english_data(self, shipment: ShipmentData) -> Dict:
        """准备英文MSDS数据（翻译）"""
        data = {
            "product_name": "",
            "section1": {},
            "section3": [],
            "section9": {},
        }

        # 产品英文名称
        if shipment.msds:
            # 如果MSDS有英文名
            data["product_name"] = shipment.product_name_en
        elif shipment.product_name_en:
            data["product_name"] = shipment.product_name_en
        elif shipment.cargo:
            data["product_name"] = shipment.cargo.product_name_en

        # 第1部分：标识
        data["section1"]["product_name"] = data["product_name"]
        data["section1"]["cas_no"] = ""
        data["section1"][" supplier"] = shipment.shipper or ""
        data["section1"]["emergency_phone"] = ""

        # 第3部分：成分（保持原文，可能需要翻译）
        if shipment.msds and shipment.msds.components:
            for comp in shipment.msds.components:
                data["section3"].append({
                    "name": self._translate_to_english(comp.name),
                    "cas_no": comp.cas_no,
                    "concentration": comp.concentration,
                })

        # 第9部分：理化特性
        if shipment.msds and shipment.msds.physical_chemical:
            pc = shipment.msds.physical_chemical
            data["section9"] = {
                "appearance": self._translate_to_english(pc.appearance) if pc.appearance else "",
                "ph": pc.ph or "",
                "melting_point": pc.melting_point or "",
                "boiling_point": pc.boiling_point or "",
                "density": pc.density or "",
                "flash_point": pc.flash_point or "",
                "solubility": self._translate_to_english(pc.solubility) if pc.solubility else "",
            }
        else:
            data["section9"] = {
                "appearance": shipment.cargo.appearance_en if shipment.cargo else "",
                "ph": "",
                "melting_point": "",
                "boiling_point": "",
                "density": "",
                "flash_point": "",
                "solubility": "",
            }

        return data

    def _translate_to_english(self, chinese_text: str) -> str:
        """
        简单的中文到英文翻译（用于MSDS）
        注意：这是基础翻译，实际生产中可能需要专业翻译
        """
        if not chinese_text:
            return ""

        # 常用翻译映射
        translations = {
            "液体": "liquid",
            "固体": "solid",
            "粉末": "powder",
            "颗粒": "granular",
            "白色": "white",
            "黄色": "yellow",
            "透明": "clear",
            "棕色": "brown",
            "黑色": "black",
            "粘稠": "viscous",
            "无色": "colorless",
            "易溶于水": "freely soluble in water",
            "不溶于水": "insoluble in water",
            "可燃": "flammable",
            "刺激": "irritant",
        }

        result = chinese_text
        for cn, en in translations.items():
            result = result.replace(cn, en)

        return result

    def _create_msds_document(self, data: Dict, output_path: str, is_english: bool):
        """创建MSDS文档"""
        try:
            from docx import Document
            from docx.shared import Pt, Inches
            from docx.enum.text import WD_ALIGN_PARAGRAPH
        except ImportError:
            raise ImportError("需要安装python-docx: pip install python-docx")

        doc = Document()

        # 设置标题
        title = "化学品安全资料说明书" if not is_english else "Material Safety Data Sheet"
        heading = doc.add_heading(title, 0)
        heading.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # 第1部分：标识
        doc.add_heading("第1部分 标识" if not is_english else "Section 1 Identification", level=1)
        self._add_field(doc, "产品名称" if not is_english else "Product Name:", data["product_name"])
        self._add_field(doc, "CAS号" if not is_english else "CAS No:", data["section1"].get("cas_no", ""))
        self._add_field(doc, "供应商" if not is_english else "Supplier:", data["section1"].get("supplier", ""))

        # 第3部分：成分/组成
        section3_title = "第3部分 成分/组成信息" if not is_english else "Section 3 Composition/Ingredients"
        doc.add_heading(section3_title, level=1)

        # 添加成分表
        if data["section3"]:
            table = doc.add_table(rows=1, cols=3)
            table.style = 'Table Grid'

            # 表头
            hdr_cells = table.rows[0].cells
            headers = ["组分", "CAS号", "含量"] if not is_english else ["Component", "CAS No.", "Concentration"]
            for i, header in enumerate(headers):
                hdr_cells[i].text = header

            # 数据行
            for comp in data["section3"]:
                row_cells = table.add_row().cells
                row_cells[0].text = comp.get("name", "")
                row_cells[1].text = comp.get("cas_no", "")
                row_cells[2].text = comp.get("concentration", "")
        else:
            doc.add_paragraph("无相关数据" if not is_english else "No relevant data")

        # 第9部分：理化特性
        section9_title = "第9部分 理化特性" if not is_english else "Section 9 Physical/Chemical Properties"
        doc.add_heading(section9_title, level=1)

        pc = data["section9"]
        self._add_field(doc, "外观与性状" if not is_english else "Appearance:", pc.get("appearance", ""))
        self._add_field(doc, "pH值" if not is_english else "pH:", pc.get("ph", ""))
        self._add_field(doc, "熔点" if not is_english else "Melting Point:", pc.get("melting_point", ""))
        self._add_field(doc, "沸点" if not is_english else "Boiling Point:", pc.get("boiling_point", ""))
        self._add_field(doc, "相对密度" if not is_english else "Density:", pc.get("density", ""))
        self._add_field(doc, "闪点" if not is_english else "Flash Point:", pc.get("flash_point", ""))
        self._add_field(doc, "溶解性" if not is_english else "Solubility:", pc.get("solubility", ""))

        doc.save(output_path)

    def _add_field(self, doc, label: str, value: str):
        """添加字段行"""
        para = doc.add_paragraph()
        run = para.add_run(f"{label} ")
        run.bold = True
        para.add_run(value)


if __name__ == "__main__":
    # 测试代码
    generator = MSDSGenerator()

    from data_models import ShipmentData, CargoInfo, MSDSInfo, Component, PhysicalChemical

    shipment = ShipmentData()
    shipment.product_name_cn = "固色剂"
    shipment.product_name_en = "Fixing Agent"
    shipment.shipper = "宏昊化工"

    # Cargo
    cargo = CargoInfo()
    cargo.product_name_cn = "固色剂（棕黑色液体）"
    cargo.product_name_en = "Fixing Agent (Brown-black liquid)"
    cargo.appearance_cn = "棕黑色液体"
    cargo.appearance_en = "Brown-black liquid"
    shipment.cargo = cargo

    # MSDS
    msds = MSDSInfo()
    msds.product_name = "固色剂"
    comp1 = Component(name="聚氨酯", cas_no="9009-54-5", concentration="30%")
    comp2 = Component(name="水", cas_no="7732-18-5", concentration="70%")
    msds.components = [comp1, comp2]

    pc = PhysicalChemical()
    pc.appearance = "棕黑色液体"
    pc.ph = "7.0"
    pc.boiling_point = "100"
    pc.density = "1.05"
    msds.physical_chemical = pc
    shipment.msds = msds

    # 生成
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    cn_path = generator.generate_chinese(shipment, desktop)
    print(f"生成中文MSDS: {cn_path}")

    en_path = generator.generate_english(shipment, desktop)
    print(f"生成英文MSDS: {en_path}")