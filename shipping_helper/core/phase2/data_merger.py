"""
数据整合器
将Phase 1数据和Phase 2解析结果整合到ShipmentData
"""

from typing import Optional, Dict, List
import os

from .data_models import ShipmentData, PackageInfo
from .export_codes_loader import ExportCodesLoader
from .report_parser import ReportParser
from .msds_parser import MSDSParser


class DataMerger:
    """
    数据整合器
    整合Phase 1 PI合同数据、Phase 2解析结果
    """

    def __init__(self):
        self.export_codes_loader = ExportCodesLoader()
        self.report_parser = ReportParser()
        self.msds_parser = MSDSParser()

    def load_export_codes(self, filepath: str) -> bool:
        """加载出口商品编码"""
        return self.export_codes_loader.load(filepath)

    def create_shipment_from_phase1(self, pi_data: dict, order_data: dict = None) -> ShipmentData:
        """
        从Phase 1数据创建ShipmentData

        Args:
            pi_data: PI合同数据字典
            order_data: 订单数据字典

        Returns:
            ShipmentData对象
        """
        return ShipmentData.from_phase1(pi_data, order_data)

    def assign_cargo(self, shipment: ShipmentData, cargo_data: dict):
        """为ShipmentData分配货物信息"""
        from .data_models import CargoInfo

        cargo = CargoInfo()
        cargo.product_name_cn = cargo_data.get("样品名称中文", "")
        cargo.product_name_en = cargo_data.get("样品名称英文", "")
        cargo.report_no = cargo_data.get("报告编号", "")
        cargo.sample_no = cargo_data.get("样品编号", "")
        cargo.appearance_cn = cargo_data.get("外观性状中文", "")
        cargo.appearance_en = cargo_data.get("外观性状英文", "")
        cargo.transport_type = cargo_data.get("运输类型", "")
        cargo.conclusion = cargo_data.get("鉴定结论", "")
        cargo.consigner = cargo_data.get("委托单位", "")
        cargo.producer = cargo_data.get("生产单位", "")

        # 判断是否危险品
        if cargo.conclusion:
            dangerous_keywords = ["危险", "dangerous", "hazard", "class"]
            cargo.is_dangerous = any(k in cargo.conclusion.lower() for k in dangerous_keywords)

        shipment.cargo = cargo

    def assign_msds(self, shipment: ShipmentData, msds_data: dict):
        """为ShipmentData分配MSDS信息"""
        from .data_models import MSDSInfo, Component, PhysicalChemical

        msds = MSDSInfo()
        msds.product_name = msds_data.get("产品名称", "")

        # 成分表
        components_data = msds_data.get("components", [])
        for comp_data in components_data:
            component = Component(
                name=comp_data.get("name", ""),
                cas_no=comp_data.get("cas_no", ""),
                concentration=comp_data.get("concentration", "")
            )
            msds.components.append(component)

        # 理化特性
        pc_data = msds_data.get("physical_chemical", {})
        if pc_data:
            pc = PhysicalChemical(
                appearance=pc_data.get("appearance", ""),
                ph=pc_data.get("ph", ""),
                boiling_point=pc_data.get("boiling_point", ""),
                melting_point=pc_data.get("melting_point", ""),
                flash_point=pc_data.get("flash_point", ""),
                density=pc_data.get("density", ""),
                solubility=pc_data.get("solubility", "")
            )
            msds.physical_chemical = pc

        shipment.msds = msds

    def assign_export_code(self, shipment: ShipmentData, internal_code: str = None):
        """根据内部编号匹配出口商品编码"""
        if not internal_code and shipment.internal_code:
            internal_code = shipment.internal_code

        if not internal_code:
            return

        # 从已加载的编码库查找
        if self.export_codes_loader.is_loaded():
            export_code = self.export_codes_loader.find_by_internal_code(internal_code)
            if export_code:
                shipment.export_code = export_code

                # 同步海关编码到主数据
                if export_code.hs_code and not shipment.hs_code:
                    shipment.hs_code = export_code.hs_code

    def assign_package_info(self, shipment: ShipmentData, package_data: dict):
        """为ShipmentData分配包装信息"""
        package = PackageInfo()
        package.package_type = package_data.get("package_type", "")
        package.quantity_kg = float(package_data.get("quantity_kg", 0))
        package.drums = int(package_data.get("drums", 0))
        package.pallets = int(package_data.get("pallets", 0))
        package.volume_cbm = float(package_data.get("volume_cbm", 0))
        package.gross_weight_kg = float(package_data.get("gross_weight_kg", 0))
        package.pallet_type = package_data.get("pallet_type", "")

        shipment.package_info = package

    def merge_all(
        self,
        pi_data: dict,
        order_data: dict = None,
        cargo_data: dict = None,
        msds_data: dict = None,
        package_data: dict = None
    ) -> ShipmentData:
        """
        合并所有数据源

        Args:
            pi_data: PI合同数据
            order_data: 订单数据
            cargo_data: 货物信息（来自运输鉴定报告）
            msds_data: MSDS信息
            package_data: 包装信息

        Returns:
            完整的ShipmentData
        """
        shipment = self.create_shipment_from_phase1(pi_data, order_data)

        if cargo_data:
            self.assign_cargo(shipment, cargo_data)

        if msds_data:
            self.assign_msds(shipment, msds_data)

        # 自动匹配出口商品编码
        if shipment.internal_code:
            self.assign_export_code(shipment)

        if package_data:
            self.assign_package_info(shipment, package_data)

        return shipment

    def auto_match_report(self, shipment: ShipmentData, reports_dir: str) -> Optional[dict]:
        """
        自动匹配运输鉴定报告

        根据产品名称或内部编号匹配报告
        """
        if not reports_dir or not shipment.product_name_cn:
            return None

        # 解析目录下的所有报告
        cargos = self.report_parser.parse_directory(reports_dir)

        for cargo in cargos:
            # 按产品名称匹配
            if shipment.product_name_cn and shipment.product_name_cn in cargo.product_name_cn:
                return {
                    "样品名称中文": cargo.product_name_cn,
                    "样品名称英文": cargo.product_name_en,
                    "报告编号": cargo.report_no,
                    "样品编号": cargo.sample_no,
                    "外观性状中文": cargo.appearance_cn,
                    "外观性状英文": cargo.appearance_en,
                    "运输类型": cargo.transport_type,
                    "鉴定结论": cargo.conclusion,
                    "委托单位": cargo.consigner,
                    "生产单位": cargo.producer,
                }

        return None

    def auto_match_msds(self, shipment: ShipmentData, msds_dir: str) -> Optional[dict]:
        """
        自动匹配MSDS

        根据产品名称匹配MSDS文档
        """
        if not msds_dir or not shipment.product_name_cn:
            return None

        if not os.path.exists(msds_dir):
            return None

        # 查找匹配的MSDS文件
        for filename in os.listdir(msds_dir):
            if not filename.endswith('.doc'):
                continue

            filepath = os.path.join(msds_dir, filename)

            # 简单文件名匹配
            if shipment.product_name_cn and shipment.product_name_cn in filename:
                msds = self.msds_parser.parse(filepath)
                if msds:
                    return self._msds_to_dict(msds)

        return None

    def _msds_to_dict(self, msds) -> dict:
        """将MSDSInfo转换为字典"""
        result = {
            "产品名称": msds.product_name,
            "components": [],
            "physical_chemical": {}
        }

        for comp in msds.components:
            result["components"].append({
                "name": comp.name,
                "cas_no": comp.cas_no,
                "concentration": comp.concentration
            })

        if msds.physical_chemical:
            pc = msds.physical_chemical
            result["physical_chemical"] = {
                "appearance": pc.appearance,
                "ph": pc.ph,
                "boiling_point": pc.boiling_point,
                "melting_point": pc.melting_point,
                "flash_point": pc.flash_point,
                "density": pc.density,
                "solubility": pc.solubility
            }

        return result


if __name__ == "__main__":
    import os

    # 测试数据整合
    merger = DataMerger()

    # 加载出口商品编码
    codes_path = r"c:\Users\windows\Desktop\shipping_helper\shipping_helper\02.订舱出货\2024.12.5 最新出口商品编码及报关成分.xlsx"
    if merger.load_export_codes(codes_path):
        print(f"加载 {merger.export_codes_loader.get_count()} 条出口商品编码")

    # 测试自动匹配
    test_dir = r"c:\Users\windows\Desktop\shipping_helper\shipping_helper\02.订舱出货\源文件\2026中观测空运运输鉴定报告"
    if os.path.exists(test_dir):
        # 创建测试用的shipment
        from data_models import ShipmentData
        shipment = ShipmentData()
        shipment.product_name_cn = "固色剂"
        shipment.internal_code = "1011"

        result = merger.auto_match_report(shipment, test_dir)
        if result:
            print(f"匹配到报告: {result.get('报告编号')}")
        else:
            print("未匹配到报告")