"""
订舱单生成器
读取Excel模板并填充数据
"""

import os
from typing import Optional, Dict

from .data_models import ShipmentData


class BookingGenerator:
    """
    订舱单生成器
    基于Excel模板生成订舱单
    """

    def __init__(self):
        self.template_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            "长晟出口海运BOOKING模板.xls"
        )

    def generate(self, shipment: ShipmentData, output_dir: str = None) -> str:
        """
        生成订舱单

        Args:
            shipment: 完整货物信息
            output_dir: 输出目录

        Returns:
            生成的订舱单文件路径
        """
        if output_dir is None:
            output_dir = os.path.join(os.path.expanduser("~"), "Desktop")

        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        # 生成文件名
        pi_no = shipment.pi_no or "Booking"
        product_name_file = shipment.product_name_cn or shipment.product_name_en or ""
        filename = f"订舱单-{pi_no}-{product_name_file}.xls"
        filename = filename.replace("/", "-").replace("\\", "-")
        output_path = os.path.join(output_dir, filename)

        # 如果模板存在，读取模板结构
        if os.path.exists(self.template_path):
            self._generate_from_template(shipment, output_path)
        else:
            # 模板不存在，生成简化版本
            self._generate_simple(shipment, output_path)

        return output_path

    def _generate_from_template(self, shipment: ShipmentData, output_path: str):
        """从Excel模板生成订舱单"""
        try:
            import xlrd
            import xlwt
            from xlwt import Workbook, XFStyle

            # 读取模板
            wb_read = xlrd.open_workbook(self.template_path)
            sh_read = wb_read.sheet_by_index(0)

            # 创建新工作簿
            wb_write = Workbook()
            ws_write = wb_write.add_sheet('订舱单')

            # 复制模板内容
            for r in range(sh_read.nrows):
                for c in range(sh_read.ncols):
                    cell = sh_read.cell_value(r, c)
                    ws_write.write(r, c, cell)

            # 填充数据到指定位置
            # 根据模板结构确定数据位置
            self._fill_shipper_data(ws_write, shipment)
            self._fill_consignee_data(ws_write, shipment)
            self._fill_notify_data(ws_write, shipment)
            self._fill_cargo_data(ws_write, shipment)

            wb_write.save(output_path)

        except ImportError as e:
            print(f"缺少库: {e}, 使用简化版本")
            self._generate_simple(shipment, output_path)
        except Exception as e:
            print(f"模板生成失败: {e}, 使用简化版本")
            self._generate_simple(shipment, output_path)

    def _fill_shipper_data(self, ws, shipment: ShipmentData):
        """填充发货人数据"""
        # 模板中Row 1是发货人区域，在A列找到标签后填充右侧内容
        # 根据模板结构，发货人名称应该在Row 1的某列
        for r in range(ws.rows):
            cell_val = str(ws.cell_value(r, 0))
            if 'Shipper' in cell_val or '发货人' in cell_val:
                # 填充发货人信息到右侧列
                if shipment.shipper:
                    ws.write(r, 1, shipment.shipper)
                if hasattr(shipment, 'shipper_address') and shipment.shipper_address:
                    ws.write(r + 1, 1, shipment.shipper_address)
                break

    def _fill_consignee_data(self, ws, shipment: ShipmentData):
        """填充收货人数据"""
        for r in range(ws.rows):
            cell_val = str(ws.cell_value(r, 0))
            if 'Consignee' in cell_val or '收货人' in cell_val:
                if shipment.consignee:
                    ws.write(r, 1, shipment.consignee)
                if hasattr(shipment, 'consignee_address') and shipment.consignee_address:
                    ws.write(r + 1, 1, shipment.consignee_address)
                break

    def _fill_notify_data(self, ws, shipment: ShipmentData):
        """填充通知人数据"""
        for r in range(ws.rows):
            cell_val = str(ws.cell_value(r, 0))
            if 'Notify' in cell_val or '通知人' in cell_val:
                if shipment.notifier:
                    ws.write(r, 1, shipment.notifier)
                break

    def _fill_cargo_data(self, ws, shipment: ShipmentData):
        """填充货物数据"""
        # 查找Port of Discharge行，填充卸货港
        for r in range(ws.rows):
            cell_val = str(ws.cell_value(r, 0))
            if 'Port of' in cell_val and 'Discharge' in cell_val:
                if shipment.destination:
                    ws.write(r, 1, shipment.destination)
                break

        # 在Row 31区域填充货物信息
        for r in range(ws.rows):
            cell_val = str(ws.cell_value(r, 0))
            if 'Marks' in cell_val or '货物' in cell_val.lower():
                # 填充货物信息到右侧
                if shipment.product_name_cn:
                    ws.write(r + 1, 2, shipment.product_name_cn)
                if shipment.package_info:
                    pkg = shipment.package_info
                    if pkg.quantity_kg:
                        ws.write(r + 1, 5, f"{int(pkg.quantity_kg)} KG")
                    if pkg.gross_weight_kg:
                        ws.write(r + 1, 7, f"{pkg.gross_weight_kg} KG")
                break

    def _generate_simple(self, shipment: ShipmentData, output_path: str):
        """生成简化版订舱单（当模板不存在时）"""
        try:
            import xlwt
            from xlwt import Workbook

            wb = Workbook()
            ws = wb.add_sheet('订舱单')

            # 标题
            ws.write(0, 0, '订舱单')
            ws.write(0, 3, f'PI号: {shipment.pi_no or ""}')

            # 发货人
            row = 2
            ws.write(row, 0, '发货人 (Shipper):')
            ws.write(row, 1, shipment.shipper or "")
            row += 1
            ws.write(row, 0, '地址:')
            ws.write(row, 1, shipment.shipper_address or "")

            # 收货人
            row += 2
            ws.write(row, 0, '收货人 (Consignee):')
            ws.write(row, 1, shipment.consignee or "")

            # 通知人
            row += 2
            ws.write(row, 0, '通知人 (Notify Party):')
            ws.write(row, 1, shipment.notifier or "")

            # 货物信息
            row += 2
            ws.write(row, 0, '品名:')
            ws.write(row, 1, shipment.product_name_cn or "")
            row += 1
            ws.write(row, 0, '内部编号:')
            ws.write(row, 1, shipment.internal_code or "")
            row += 1
            ws.write(row, 0, 'H.S. Code:')
            ws.write(row, 1, shipment.hs_code or "")

            # 包装信息
            if shipment.package_info:
                row += 2
                pkg = shipment.package_info
                ws.write(row, 0, '包装类型:')
                ws.write(row, 1, pkg.package_type or "")
                row += 1
                ws.write(row, 0, '数量:')
                ws.write(row, 1, f"{pkg.quantity_kg} KG" if pkg.quantity_kg else "")
                row += 1
                ws.write(row, 0, '桶数:')
                ws.write(row, 1, str(pkg.drums) if pkg.drums else "")
                row += 1
                ws.write(row, 0, '卡板数:')
                ws.write(row, 1, str(pkg.pallets) if pkg.pallets else "")
                row += 1
                ws.write(row, 0, '体积:')
                ws.write(row, 1, f"{pkg.volume_cbm} CBM" if pkg.volume_cbm else "")
                row += 1
                ws.write(row, 0, '毛重:')
                ws.write(row, 1, f"{pkg.gross_weight_kg} KG" if pkg.gross_weight_kg else "")

            # 卸货港
            row += 2
            ws.write(row, 0, '卸货港:')
            ws.write(row, 1, shipment.destination or "")

            wb.save(output_path)

        except ImportError:
            raise ImportError("需要xlwt库来生成Excel文件: pip install xlwt")