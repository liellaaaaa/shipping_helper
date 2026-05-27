"""
LOI保函生成器
生成非危险品保函和液体保函
复用 shipping_helper_old/core/loi_filler.py 的占位符逻辑
"""

import os
import shutil
from typing import Optional, Dict

from .data_models import ShipmentData


class LOIGenerator:
    """
    LOI保函生成器
    根据货物类型选择模板（非危险品/液体）
    模板路径：02.订舱出货/源文件/LOI-op-非危险品保函模板.docx
              02.订舱出货/源文件/LOI-op-液体保函模板.docx
    """

    # 默认模板路径
    DEFAULT_TEMPLATES = {
        "non_hazardous": r"02.订舱出货\源文件\LOI-op-非危险品保函模板.docx",
        "liquid": r"02.订舱出货\源文件\LOI-op-液体保函模板.docx",
    }

    # 占位符格式
    PLACEHOLDER_PATTERN = __import__('re').compile(r'\{\{(\w+)\}\}')

    def __init__(self, templates_dir: str = None):
        """
        初始化LOI生成器

        Args:
            templates_dir: 模板目录路径，默认为 shipping_helper/02.订舱出货/源文件/
        """
        if templates_dir:
            self.templates_dir = templates_dir
        else:
            # 默认使用项目内的模板目录
            base_dir = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
            self.templates_dir = os.path.join(base_dir, "02.订舱出货", "源文件")

    def select_template(self, shipment: ShipmentData) -> str:
        """
        根据货物信息选择模板

        Args:
            shipment: 货物信息

        Returns:
            模板文件路径
        """
        is_liquid = shipment.is_liquid() if shipment else False

        template_key = "liquid" if is_liquid else "non_hazardous"
        template_filename = os.path.basename(self.DEFAULT_TEMPLATES[template_key])

        template_path = os.path.join(self.templates_dir, template_filename)

        if not os.path.exists(template_path):
            # 尝试从templates_dir
            template_path = os.path.join(self.templates_dir, template_filename)
            if not os.path.exists(template_path):
                raise FileNotFoundError(f"模板文件不存在: {template_path}")

        return template_path

    def generate(self, shipment: ShipmentData, output_dir: str = None) -> str:
        """
        生成LOI保函文件

        Args:
            shipment: 完整货物信息
            output_dir: 输出目录，默认为桌面

        Returns:
            生成的保函文件路径
        """
        from docx import Document

        if output_dir is None:
            output_dir = os.path.join(os.path.expanduser("~"), "Desktop")

        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        # 选择模板
        template_path = self.select_template(shipment)

        # 准备数据
        data = self._prepare_data(shipment)

        # 打开模板
        doc = Document(template_path)

        # 替换占位符
        for para in doc.paragraphs:
            self._replace_in_text(para, data)

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        self._replace_in_text(para, data)

        # 生成文件名
        product_name = shipment.product_name_cn or shipment.product_name_en or "LOI"
        filename = f"LOI-{product_name}.docx"
        output_path = os.path.join(output_dir, filename)

        doc.save(output_path)
        return output_path

    def _replace_in_text(self, para, data: Dict[str, str]):
        """替换段落或表格中的占位符"""
        if not para.text:
            return

        text = para.text
        matches = list(self.PLACEHOLDER_PATTERN.finditer(text))

        if not matches:
            return

        for match in matches:
            placeholder = match.group(1)
            value = self._get_field_value(placeholder, data)
            if value:
                text = text.replace(match.group(0), value)

        if matches:
            for run in para.runs:
                if run.text:
                    for m in matches:
                        if m.group(0) in run.text:
                            run.text = run.text.replace(m.group(0), self._get_field_value(m.group(1), data) or '')

    def _get_field_value(self, key: str, data: Dict[str, str]) -> str:
        """获取字段值"""
        return data.get(key, '')

    def _prepare_data(self, shipment: ShipmentData) -> Dict[str, str]:
        """准备填充数据"""
        data = {}

        # 发货人信息
        data['发货人'] = shipment.shipper or ""
        data['发货人地址'] = shipment.shipper_address or ""

        # 收货人信息
        data['收货人'] = shipment.consignee or ""
        data['收货人地址'] = shipment.consignee_address or ""

        # 通知人信息
        data['通知人'] = shipment.notifier or ""
        data['通知人地址'] = shipment.notifier_address or ""

        # PI号
        data['PI号'] = shipment.pi_no or ""

        # 产品信息
        if shipment.cargo:
            data['样品名称中文'] = shipment.cargo.product_name_cn or ""
            data['样品名称英文'] = shipment.cargo.product_name_en or ""
            data['报告编号'] = shipment.cargo.report_no or ""
            data['外观性状'] = shipment.cargo.appearance_cn or shipment.cargo.appearance_en or ""
            data['运输类型'] = shipment.cargo.transport_type or ""

            # 判断是否危险品
            if shipment.cargo.is_dangerous:
                data['货物类型'] = '危险品'
            else:
                data['货物类型'] = '非危险品'
        else:
            data['样品名称中文'] = shipment.product_name_cn or ""
            data['样品名称英文'] = shipment.product_name_en or ""

        # MSDS成分信息
        if shipment.msds and shipment.msds.components:
            component_names = [c.name for c in shipment.msds.components if c.name]
            data['主要成分'] = '、'.join(component_names[:5])  # 最多5个
        elif shipment.export_code and shipment.export_code.composition:
            data['主要成分'] = shipment.export_code.composition
        else:
            data['主要成分'] = ""

        # 包装信息
        if shipment.package_info:
            data['桶数'] = str(shipment.package_info.drums or "")
            data['卡板数'] = str(shipment.package_info.pallets or "")
            data['数量kg'] = str(shipment.package_info.quantity_kg or "")
            data['体积CBM'] = str(shipment.package_info.volume_cbm or "")
            data['毛重kg'] = str(shipment.package_info.gross_weight_kg or "")
        else:
            data['桶数'] = ""
            data['卡板数'] = ""
            data['数量kg'] = ""
            data['体积CBM'] = ""
            data['毛重kg'] = ""

        # 内部编号
        data['内部编号'] = shipment.internal_code or ""

        # 海关编码
        data['H.S.Code'] = shipment.hs_code or ""

        return data

    def prepare_for_edit(self, shipment: ShipmentData, output_path: str) -> bool:
        """
        准备编辑模式 - 复制模板到输出路径供用户在Word中编辑

        Args:
            shipment: 货物信息
            output_path: 输出文件路径

        Returns:
            是否成功
        """
        template_path = self.select_template(shipment)

        # 直接复制模板到输出路径
        shutil.copy2(template_path, output_path)
        return True

    def get_editable_fields(self, template_path: str = None) -> list:
        """获取模板中的可编辑字段列表"""
        from docx import Document

        if template_path is None:
            template_path = self.select_template(ShipmentData())

        doc = Document(template_path)
        fields = set()

        for para in doc.paragraphs:
            matches = self.PLACEHOLDER_PATTERN.findall(para.text)
            fields.update(matches)

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        matches = self.PLACEHOLDER_PATTERN.findall(para.text)
                        fields.update(matches)

        return list(fields)


if __name__ == "__main__":
    # 测试代码
    generator = LOIGenerator()

    # 测试模板选择
    from data_models import ShipmentData, CargoInfo

    shipment = ShipmentData()
    shipment.product_name_cn = "固色剂"
    shipment.shipper = "宏昊"
    shipment.consignee = "TOA-DOVECHEM"
    shipment.pi_no = "HT2024001"

    cargo = CargoInfo()
    cargo.product_name_cn = "固色剂（棕黑色液体）"
    cargo.report_no = "NACCWX25033581"
    cargo.appearance_cn = "棕黑色液体"
    cargo.transport_type = "空运"
    shipment.cargo = cargo

    print(f"选择的模板: {generator.select_template(shipment)}")
    print(f"是否为液体: {shipment.is_liquid()}")