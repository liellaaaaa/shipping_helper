"""
Phase 2 Data Models
Defines all core data classes for the booking/shipping module.
"""

from dataclasses import dataclass, field
from typing import List, Optional


@dataclass
class Component:
    """成分项 - MSDS第3部分"""
    name: str = ""           # 组分名称
    cas_no: str = ""         # CAS号
    concentration: str = ""  # 含量（如 "30%"）


@dataclass
class PhysicalChemical:
    """理化特性 - MSDS第9部分"""
    appearance: str = ""      # 外观与性状
    ionic: str = ""          # 离子性
    ph: str = ""             # pH值
    melting_point: str = ""   # 熔点
    boiling_point: str = ""  # 沸点/沸点范围（℃）
    density: str = ""         # 相对密度
    vapor_pressure: str = "" # 蒸汽压
    flash_point: str = ""    # 闪点
    autoignition: str = ""   # 自燃点（℃）
    explosion_lower: str = "" # 爆炸下限
    explosion_upper: str = "" # 爆炸上限
    solubility: str = ""      # 溶解性


@dataclass
class CargoInfo:
    """货物信息 - 从运输鉴定报告提取"""
    product_name_cn: str = ""     # 样品名称（中文）
    product_name_en: str = ""     # 样品名称（英文）
    report_no: str = ""          # 鉴定书编号
    sample_no: str = ""          # 样品编号
    appearance_cn: str = ""      # 外观描述（中文）
    appearance_en: str = ""      # 外观描述（英文）
    transport_type: str = ""      # 运输类型：海运/空运/陆运
    conclusion: str = ""         # 鉴定结论
    is_dangerous: bool = False   # 是否危险品
    consigner: str = ""          # 委托单位
    producer: str = ""           # 生产单位


@dataclass
class MSDSInfo:
    """MSDS信息 - 从MSDS文档提取"""
    product_name: str = ""                       # 第1部分：产品名称
    components: List[Component] = field(default_factory=list)  # 第3部分：成分表
    physical_chemical: Optional[PhysicalChemical] = None       # 第9部分：理化特性

    # 额外信息
    section1_identifiers: str = ""  # 标识信息（第1部分原始文本）
    section3_raw: str = ""         # 第3部分原始文本
    section9_raw: str = ""         # 第9部分原始文本


@dataclass
class ExportCode:
    """出口商品编码 - 从Excel提取"""
    internal_code: str = ""    # 产品内编
    appearance: str = ""       # 产品外观
    hs_code: str = ""         # 新商品编码/海关编码
    customs_name: str = ""    # 报关名称
    composition: str = ""     # 报关成分


@dataclass
class PackageInfo:
    """包装信息 - 来自Phase 1包装计算"""
    package_type: str = ""     # 包装类型（如 "30kg蓝桶"）
    quantity_kg: float = 0.0   # 数量(kg)
    drums: int = 0             # 桶数
    pallets: int = 0           # 卡板数
    volume_cbm: float = 0.0    # 体积(CBM)
    gross_weight_kg: float = 0.0  # 毛重(kg)
    pallet_type: str = ""      # 卡板类型（1.0*1.0m / 1.1*1.1m）


@dataclass
class ShipmentData:
    """完整订舱数据 - 整合Phase 1和Phase 2"""

    # === 继承自Phase 1 PI合同 ===
    shipper: str = ""          # 发货人
    shipper_address: str = ""  # 发货人地址
    consignee: str = ""        # 收货人
    consignee_address: str = "" # 收货人地址
    notifier: str = ""         # 通知人
    notifier_address: str = "" # 通知人地址
    pi_no: str = ""            # PI号
    internal_code: str = ""    # 内部编号
    product_name_cn: str = ""  # 产品中文名
    product_name_en: str = ""  # 产品英文名
    hs_code: str = ""          # H.S.Code
    destination: str = ""      # 卸货港/目的地
    packing_note: str = ""     # 包装说明

    # === 从运输鉴定报告提取 ===
    cargo: Optional[CargoInfo] = None

    # === 从MSDS提取 ===
    msds: Optional[MSDSInfo] = None

    # === 从出口商品编码表提取 ===
    export_code: Optional[ExportCode] = None

    # === 包装信息（来自Phase 1）===
    package_info: Optional[PackageInfo] = None

    @classmethod
    def from_phase1(cls, pi_data: dict, order_data: dict = None) -> "ShipmentData":
        """从Phase 1数据初始化ShipmentData"""
        shipment = cls()

        # 从PI数据继承
        shipment.consignee = pi_data.get("收货人", "")
        shipment.consignee_address = pi_data.get("收货人地址", "")
        shipment.notifier = pi_data.get("收货人", "")  # 默认 = 收货人
        shipment.notifier_address = pi_data.get("收货人地址", "")
        shipment.pi_no = pi_data.get("PI号", "")
        shipment.product_name_en = pi_data.get("品名英文", "")
        shipment.hs_code = pi_data.get("H.S.Code", "")
        shipment.destination = pi_data.get("卸货港", "")
        shipment.packing_note = pi_data.get("包装说明", "")

        # 从订单数据继承
        if order_data:
            shipment.internal_code = order_data.get("内部编号", "")
            shipment.product_name_cn = order_data.get("产品中文名", "")

        return shipment

    def set_shipper(self, name: str, address: str = ""):
        """设置发货人"""
        self.shipper = name
        self.shipper_address = address
        # 默认通知人同收货人
        if not self.notifier:
            self.notifier = self.consignee
            self.notifier_address = self.consignee_address

    def is_liquid(self) -> bool:
        """判断是否为液体货物"""
        if self.cargo and self.cargo.appearance_cn:
            liquid_keywords = ["液", "水剂", "浆"]
            return any(k in self.cargo.appearance_cn for k in liquid_keywords)
        if self.msds and self.msds.physical_chemical:
            return "液" in self.msds.physical_chemical.appearance
        return False