# -*- coding: utf-8 -*-
"""
订舱出货模块 - Phase 2 主界面
左右分栏布局：
- 左侧：数据看板（展示提取的结构化数据）
- 右侧：模板编辑区（嵌入Word ActiveX，支持WYSIWYG编辑）
"""

from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
                             QLabel, QGroupBox, QScrollArea, QTableWidget,
                             QTableWidgetItem, QHeaderView, QTextEdit, QComboBox,
                             QLineEdit, QTabWidget, QMessageBox, QFileDialog,
                             QGridLayout, QTableView, QAbstractItemView)
from PyQt5.QtCore import Qt, QSize
from PyQt5.QtGui import QFont
from PyQt5.QAxContainer import QAxWidget
import os
import sys

# 添加父目录到路径以支持导入
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.phase2 import (
    ShipmentData,
    DataMerger,
    ReportParser,
    MSDSParser,
    BookingGenerator,
    MSDSGenerator,
    LOIGenerator,
)


class BookingWidget(QWidget):
    """
    Phase 2 订舱出货模块主界面
    左右分栏布局
    """

    def __init__(self, parent=None, phase1_data=None):
        super().__init__(parent)
        self.phase1_data = phase1_data or {}
        self.shipment = None
        self.data_merger = DataMerger()
        self.report_parser = ReportParser()
        self.msds_parser = MSDSParser()

        self._init_ui()

    def _init_ui(self):
        """初始化UI"""
        layout = QHBoxLayout()
        self.setLayout(layout)

        # 左侧：数据看板
        left_widget = self._create_left_panel()
        layout.addWidget(left_widget, 1)

        # 分隔线
        splitter = QWidget()
        splitter.setFixedWidth(5)
        splitter.setStyleSheet("background-color: #ddd;")
        layout.addWidget(splitter)

        # 右侧：模板编辑区
        right_widget = self._create_right_panel()
        layout.addWidget(right_widget, 2)

    def _create_left_panel(self) -> QWidget:
        """创建左侧面板：数据看板"""
        panel = QWidget()
        layout = QVBoxLayout()
        panel.setLayout(layout)

        # 标题
        title = QLabel("数据看板")
        title.setStyleSheet("font-size: 16px; font-weight: bold; padding: 10px;")
        layout.addWidget(title)

        # 滚动区域
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setMinimumWidth(350)

        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout()

        # === Phase 1 继承数据 ===
        phase1_group = QGroupBox("继承自 Phase 1 (PI合同)")
        phase1_layout = QGridLayout()

        self.lbl_shipper = QLabel("发货人: -")
        self.lbl_consignee = QLabel("收货人: -")
        self.lbl_notifier = QLabel("通知人: -")
        self.lbl_pi_no = QLabel("PI号: -")
        self.lbl_internal_code = QLabel("内部编号: -")
        self.lbl_product_name = QLabel("产品名称: -")
        self.lbl_hs_code = QLabel("H.S.Code: -")
        self.lbl_destination = QLabel("卸货港: -")

        phase1_layout.addWidget(self.lbl_shipper, 0, 0)
        phase1_layout.addWidget(self.lbl_consignee, 0, 1)
        phase1_layout.addWidget(self.lbl_notifier, 1, 0)
        phase1_layout.addWidget(self.lbl_pi_no, 1, 1)
        phase1_layout.addWidget(self.lbl_internal_code, 2, 0)
        phase1_layout.addWidget(self.lbl_product_name, 2, 1)
        phase1_layout.addWidget(self.lbl_hs_code, 3, 0)
        phase1_layout.addWidget(self.lbl_destination, 3, 1)

        phase1_group.setLayout(phase1_layout)
        scroll_layout.addWidget(phase1_group)

        # === 运输鉴定报告 ===
        self.cargo_group = QGroupBox("运输鉴定报告")
        cargo_layout = QGridLayout()

        self.lbl_report_no = QLabel("报告编号: -")
        self.lbl_transport_type = QLabel("运输类型: -")
        self.lbl_appearance_cn = QLabel("外观(中): -")
        self.lbl_appearance_en = QLabel("外观(英): -")
        self.lbl_conclusion = QLabel("鉴定结论: -")

        cargo_layout.addWidget(self.lbl_report_no, 0, 0)
        cargo_layout.addWidget(self.lbl_transport_type, 0, 1)
        cargo_layout.addWidget(self.lbl_appearance_cn, 1, 0)
        cargo_layout.addWidget(self.lbl_appearance_en, 1, 1)
        cargo_layout.addWidget(self.lbl_conclusion, 2, 0, 1, 2)

        self.cargo_group.setLayout(cargo_layout)
        scroll_layout.addWidget(self.cargo_group)

        # === MSDS 信息 ===
        self.msds_group = QGroupBox("MSDS 信息")
        msds_layout = QVBoxLayout()

        self.msds_product_label = QLabel("产品名称: -")
        msds_layout.addWidget(self.msds_product_label)

        # 成分表
        self.msds_components_table = QTableWidget()
        self.msds_components_table.setColumnCount(3)
        self.msds_components_table.setHorizontalHeaderLabels(["组分", "CAS号", "含量"])
        self.msds_components_table.setMaximumHeight(120)
        self.msds_components_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        msds_layout.addWidget(self.msds_components_table)

        # 理化特性
        pc_layout = QGridLayout()
        self.lbl_ph = QLabel("pH: -")
        self.lbl_boiling = QLabel("沸点: -")
        self.lbl_density = QLabel("密度: -")
        self.lbl_flash = QLabel("闪点: -")

        pc_layout.addWidget(self.lbl_ph, 0, 0)
        pc_layout.addWidget(self.lbl_boiling, 0, 1)
        pc_layout.addWidget(self.lbl_density, 1, 0)
        pc_layout.addWidget(self.lbl_flash, 1, 1)
        msds_layout.addLayout(pc_layout)

        self.msds_group.setLayout(msds_layout)
        scroll_layout.addWidget(self.msds_group)

        # === 出口商品编码 ===
        self.export_code_group = QGroupBox("出口商品编码")
        ec_layout = QGridLayout()

        self.lbl_export_hs = QLabel("海关编码: -")
        self.lbl_export_customs = QLabel("报关名称: -")
        self.lbl_export_composition = QLabel("报关成分: -")

        ec_layout.addWidget(self.lbl_export_hs, 0, 0)
        ec_layout.addWidget(self.lbl_export_customs, 0, 1)
        ec_layout.addWidget(self.lbl_export_composition, 1, 0, 1, 2)

        self.export_code_group.setLayout(ec_layout)
        scroll_layout.addWidget(self.export_code_group)

        # === 包装信息 ===
        self.package_group = QGroupBox("包装信息")
        pkg_layout = QGridLayout()

        self.lbl_pkg_type = QLabel("包装类型: -")
        self.lbl_pkg_qty = QLabel("数量: -")
        self.lbl_pkg_drums = QLabel("桶数: -")
        self.lbl_pkg_pallets = QLabel("卡板数: -")
        self.lbl_pkg_volume = QLabel("体积: -")
        self.lbl_pkg_weight = QLabel("毛重: -")

        pkg_layout.addWidget(self.lbl_pkg_type, 0, 0)
        pkg_layout.addWidget(self.lbl_pkg_qty, 0, 1)
        pkg_layout.addWidget(self.lbl_pkg_drums, 1, 0)
        pkg_layout.addWidget(self.lbl_pkg_pallets, 1, 1)
        pkg_layout.addWidget(self.lbl_pkg_volume, 2, 0)
        pkg_layout.addWidget(self.lbl_pkg_weight, 2, 1)

        self.package_group.setLayout(pkg_layout)
        scroll_layout.addWidget(self.package_group)

        scroll_widget.setLayout(scroll_layout)
        scroll.setWidget(scroll_widget)
        layout.addWidget(scroll, 1)

        # 底部按钮
        btn_layout = QHBoxLayout()
        self.btn_import_phase1 = QPushButton("从 Phase 1 导入")
        self.btn_import_phase1.clicked.connect(self._import_from_phase1)
        btn_layout.addWidget(self.btn_import_phase1)

        self.btn_load_codes = QPushButton("加载商品编码")
        self.btn_load_codes.clicked.connect(self._load_export_codes)
        btn_layout.addWidget(self.btn_load_codes)

        layout.addLayout(btn_layout)

        return panel

    def _create_right_panel(self) -> QWidget:
        """创建右侧面板：模板编辑区"""
        panel = QWidget()
        layout = QVBoxLayout()
        panel.setLayout(layout)

        # 标题
        title = QLabel("模板编辑区")
        title.setStyleSheet("font-size: 16px; font-weight: bold; padding: 10px;")
        layout.addWidget(title)

        # Tab选项卡
        self.template_tabs = QTabWidget()
        self.template_tabs.setMinimumHeight(400)

        # Tab 1: 订舱单
        self.booking_tab = QWidget()
        self._setup_booking_tab()
        self.template_tabs.addTab(self.booking_tab, "订舱单")

        # Tab 2: MSDS
        self.msds_tab = QWidget()
        self._setup_msds_tab()
        self.template_tabs.addTab(self.msds_tab, "MSDS")

        # Tab 3: 保函
        self.loi_tab = QWidget()
        self._setup_loi_tab()
        self.template_tabs.addTab(self.loi_tab, "保函")

        layout.addWidget(self.template_tabs, 1)

        # 底部按钮
        btn_layout = QHBoxLayout()

        self.btn_generate_booking = QPushButton("生成订舱单")
        self.btn_generate_booking.clicked.connect(self._generate_booking)
        self.btn_generate_booking.setEnabled(False)
        btn_layout.addWidget(self.btn_generate_booking)

        self.btn_generate_msds = QPushButton("生成 MSDS")
        self.btn_generate_msds.clicked.connect(self._generate_msds)
        self.btn_generate_msds.setEnabled(False)
        btn_layout.addWidget(self.btn_generate_msds)

        self.btn_generate_loi = QPushButton("生成保函")
        self.btn_generate_loi.clicked.connect(self._generate_loi)
        self.btn_generate_loi.setEnabled(False)
        btn_layout.addWidget(self.btn_generate_loi)

        btn_layout.addStretch()

        layout.addLayout(btn_layout)

        return panel

    def _setup_booking_tab(self):
        """设置订舱单Tab"""
        layout = QVBoxLayout()
        self.booking_tab.setLayout(layout)

        info_label = QLabel("订舱单预览（只读）")
        info_label.setStyleSheet("color: #666; padding: 5px;")
        layout.addWidget(info_label)

        # 订舱单预览区域（简化版本，使用表格）
        self.booking_preview = QTextEdit()
        self.booking_preview.setReadOnly(True)
        self.booking_preview.setPlaceholderText("点击「生成订舱单」后，预览将显示在这里...")
        layout.addWidget(self.booking_preview)

    def _setup_msds_tab(self):
        """设置MSDS Tab"""
        layout = QVBoxLayout()
        self.msds_tab.setLayout(layout)

        info_label = QLabel("中文/英文 MSDS 生成")
        info_label.setStyleSheet("color: #666; padding: 5px;")
        layout.addWidget(info_label)

        # MSDS预览（简化版本）
        self.msds_preview = QTextEdit()
        self.msds_preview.setReadOnly(True)
        self.msds_preview.setPlaceholderText("点击「生成 MSDS」后，预览将显示在这里...")
        layout.addWidget(self.msds_preview)

    def _setup_loi_tab(self):
        """设置保函Tab"""
        layout = QVBoxLayout()
        self.loi_tab.setLayout(layout)

        info_label = QLabel("LOI 保函模板选择")
        info_label.setStyleSheet("color: #666; padding: 5px;")
        layout.addWidget(info_label)

        # 模板类型选择
        template_layout = QHBoxLayout()
        template_layout.addWidget(QLabel("模板类型:"))

        self.loi_template_combo = QComboBox()
        self.loi_template_combo.addItems(["非危险品保函", "液体保函"])
        template_layout.addWidget(self.loi_template_combo)
        template_layout.addStretch()

        layout.addLayout(template_layout)

        # 保函预览
        self.loi_preview = QTextEdit()
        self.loi_preview.setReadOnly(True)
        self.loi_preview.setPlaceholderText("点击「生成保函」后，预览将显示在这里...")
        layout.addWidget(self.loi_preview)

    def _import_from_phase1(self):
        """从Phase 1导入数据"""
        if not self.phase1_data:
            QMessageBox.warning(self, "警告", "没有可导入的 Phase 1 数据")
            return

        # 创建ShipmentData
        self.shipment = ShipmentData.from_phase1(self.phase1_data)

        # 设置发货人（默认宏昊）
        self.shipment.set_shipper("宏昊")

        # 更新UI
        self._update_phase1_ui()

        QMessageBox.information(self, "成功", "已从 Phase 1 导入数据")

    def _load_export_codes(self):
        """加载出口商品编码"""
        default_path = os.path.join(
            os.path.dirname(os.path.dirname(os.path.dirname(__file__))),
            "02.订舱出货",
            "2024.12.5 最新出口商品编码及报关成分.xlsx"
        )

        filepath, _ = QFileDialog.getOpenFileName(
            self, "选择出口商品编码Excel", default_path,
            "Excel Files (*.xlsx *.xls);;All Files (*)"
        )

        if filepath:
            if self.data_merger.load_export_codes(filepath):
                QMessageBox.information(self, "成功", f"加载成功，共 {self.data_merger.export_codes_loader.get_count()} 条数据")
            else:
                QMessageBox.warning(self, "失败", "加载出口商品编码失败")

    def _update_phase1_ui(self):
        """更新Phase 1数据UI"""
        if not self.shipment:
            return

        self.lbl_shipper.setText(f"发货人: {self.shipment.shipper}")
        self.lbl_consignee.setText(f"收货人: {self.shipment.consignee}")
        self.lbl_notifier.setText(f"通知人: {self.shipment.notifier}")
        self.lbl_pi_no.setText(f"PI号: {self.shipment.pi_no}")
        self.lbl_internal_code.setText(f"内部编号: {self.shipment.internal_code}")
        self.lbl_product_name.setText(f"产品名称: {self.shipment.product_name_cn}")
        self.lbl_hs_code.setText(f"H.S.Code: {self.shipment.hs_code}")
        self.lbl_destination.setText(f"卸货港: {self.shipment.destination}")

        # 自动匹配出口商品编码
        if self.shipment.internal_code and self.data_merger.export_codes_loader.is_loaded():
            self.data_merger.assign_export_code(self.shipment)
            self._update_export_code_ui()

        # 启用生成按钮
        self._enable_generate_buttons()

    def _update_export_code_ui(self):
        """更新出口商品编码UI"""
        if not self.shipment or not self.shipment.export_code:
            return

        ec = self.shipment.export_code
        self.lbl_export_hs.setText(f"海关编码: {ec.hs_code}")
        self.lbl_export_customs.setText(f"报关名称: {ec.customs_name}")
        self.lbl_export_composition.setText(f"报关成分: {ec.composition}")

    def _update_cargo_ui(self):
        """更新货物信息UI"""
        if not self.shipment or not self.shipment.cargo:
            return

        cargo = self.shipment.cargo
        self.lbl_report_no.setText(f"报告编号: {cargo.report_no}")
        self.lbl_transport_type.setText(f"运输类型: {cargo.transport_type}")
        self.lbl_appearance_cn.setText(f"外观(中): {cargo.appearance_cn}")
        self.lbl_appearance_en.setText(f"外观(英): {cargo.appearance_en}")
        self.lbl_conclusion.setText(f"鉴定结论: {cargo.conclusion}")

        # 更新模板选项卡显示
        if cargo.appearance_cn and '液' in cargo.appearance_cn:
            self.loi_template_combo.setCurrentIndex(1)  # 液体保函
        else:
            self.loi_template_combo.setCurrentIndex(0)  # 非危险品保函

    def _update_msds_ui(self):
        """更新MSDS信息UI"""
        if not self.shipment or not self.shipment.msds:
            return

        msds = self.shipment.msds
        self.msds_product_label.setText(f"产品名称: {msds.product_name}")

        # 更新成分表
        self.msds_components_table.setRowCount(len(msds.components))
        for i, comp in enumerate(msds.components):
            self.msds_components_table.setItem(i, 0, QTableWidgetItem(comp.name))
            self.msds_components_table.setItem(i, 1, QTableWidgetItem(comp.cas_no))
            self.msds_components_table.setItem(i, 2, QTableWidgetItem(comp.concentration))

        # 理化特性
        if msds.physical_chemical:
            pc = msds.physical_chemical
            self.lbl_ph.setText(f"pH: {pc.ph}")
            self.lbl_boiling.setText(f"沸点: {pc.boiling_point}")
            self.lbl_density.setText(f"密度: {pc.density}")
            self.lbl_flash.setText(f"闪点: {pc.flash_point}")

    def _enable_generate_buttons(self):
        """启用生成按钮"""
        if self.shipment:
            self.btn_generate_booking.setEnabled(True)
            self.btn_generate_msds.setEnabled(True)
            self.btn_generate_loi.setEnabled(True)

    def _generate_booking(self):
        """生成订舱单"""
        if not self.shipment:
            QMessageBox.warning(self, "警告", "请先导入数据")
            return

        try:
            generator = BookingGenerator()
            output_dir = os.path.join(os.path.expanduser("~"), "Desktop")
            output_path = generator.generate(self.shipment, output_dir)

            # 预览
            preview_text = f"订舱单已生成:\n{output_path}"
            self.booking_preview.setPlainText(preview_text)

            QMessageBox.information(self, "成功", f"订舱单已保存到桌面:\n{os.path.basename(output_path)}")

        except Exception as e:
            QMessageBox.warning(self, "失败", f"生成订舱单失败:\n{str(e)}")

    def _generate_msds(self):
        """生成MSDS"""
        if not self.shipment:
            QMessageBox.warning(self, "警告", "请先导入数据")
            return

        try:
            generator = MSDSGenerator()
            output_dir = os.path.join(os.path.expanduser("~"), "Desktop")

            cn_path, en_path = generator.generate_both(self.shipment, output_dir)

            # 预览
            preview_text = f"中文MSDS: {os.path.basename(cn_path)}\n英文MSDS: {os.path.basename(en_path)}"
            self.msds_preview.setPlainText(preview_text)

            QMessageBox.information(self, "成功", f"MSDS已保存到桌面")

        except Exception as e:
            QMessageBox.warning(self, "失败", f"生成MSDS失败:\n{str(e)}")

    def _generate_loi(self):
        """生成保函"""
        if not self.shipment:
            QMessageBox.warning(self, "警告", "请先导入数据")
            return

        try:
            generator = LOIGenerator()
            output_dir = os.path.join(os.path.expanduser("~"), "Desktop")

            output_path = generator.generate(self.shipment, output_dir)

            # 预览
            preview_text = f"保函已生成:\n{output_path}"
            self.loi_preview.setPlainText(preview_text)

            QMessageBox.information(self, "成功", f"保函已保存到桌面:\n{os.path.basename(output_path)}")

        except Exception as e:
            QMessageBox.warning(self, "失败", f"生成保函失败:\n{str(e)}")

    def set_phase1_data(self, data: dict):
        """设置Phase 1数据（从外部调用）"""
        self.phase1_data = data
        if data:
            self._import_from_phase1()

    def set_package_data(self, pkg_data: dict):
        """设置包装数据（从Phase 1的包装计算结果）"""
        if self.shipment and pkg_data:
            from core.phase2 import PackageInfo
            pkg = PackageInfo(
                package_type=pkg_data.get("package_type", ""),
                quantity_kg=float(pkg_data.get("quantity_kg", 0)),
                drums=int(pkg_data.get("drums", 0)),
                pallets=int(pkg_data.get("pallets", 0)),
                volume_cbm=float(pkg_data.get("volume_cbm", 0)),
                gross_weight_kg=float(pkg_data.get("gross_weight_kg", 0)),
                pallet_type=pkg_data.get("pallet_type", "")
            )
            self.shipment.package_info = pkg
            self._update_package_ui()

    def _update_package_ui(self):
        """更新包装信息UI"""
        if not self.shipment or not self.shipment.package_info:
            return

        pkg = self.shipment.package_info
        self.lbl_pkg_type.setText(f"包装类型: {pkg.package_type}")
        self.lbl_pkg_qty.setText(f"数量: {pkg.quantity_kg} kg")
        self.lbl_pkg_drums.setText(f"桶数: {pkg.drums}")
        self.lbl_pkg_pallets.setText(f"卡板数: {pkg.pallets}")
        self.lbl_pkg_volume.setText(f"体积: {pkg.volume_cbm} CBM")
        self.lbl_pkg_weight.setText(f"毛重: {pkg.gross_weight_kg} kg")