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
                             QGridLayout, QTableView, QAbstractItemView, QSizePolicy)
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
        self.output_dir = None  # PI号对应的输出目录

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
        scroll.setMinimumWidth(450)

        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout()

        # === Phase 1 继承数据 ===
        phase1_group = QGroupBox("继承自 Phase 1 (PI合同)")
        phase1_layout = QGridLayout()

        # 发货人（可编辑）
        phase1_layout.addWidget(QLabel("发货人:"), 0, 0)
        self.edit_shipper = QLineEdit("-")
        self.edit_shipper.textChanged.connect(self._on_shipper_changed)
        phase1_layout.addWidget(self.edit_shipper, 0, 1)

        # 收货人（可编辑）
        phase1_layout.addWidget(QLabel("收货人:"), 0, 2)
        self.edit_consignee = QLineEdit("-")
        self.edit_consignee.textChanged.connect(self._on_consignee_changed)
        phase1_layout.addWidget(self.edit_consignee, 0, 3)

        # 通知人（可编辑）
        phase1_layout.addWidget(QLabel("通知人:"), 1, 0)
        self.edit_notifier = QLineEdit("-")
        self.edit_notifier.textChanged.connect(self._on_notifier_changed)
        phase1_layout.addWidget(self.edit_notifier, 1, 1)

        # PI号（可编辑）
        phase1_layout.addWidget(QLabel("PI号:"), 1, 2)
        self.edit_pi_no = QLineEdit("-")
        self.edit_pi_no.textChanged.connect(self._on_pi_no_changed)
        phase1_layout.addWidget(self.edit_pi_no, 1, 3)

        # 内部编号（可编辑）
        phase1_layout.addWidget(QLabel("内部编号:"), 2, 0)
        self.edit_internal_code = QLineEdit("-")
        self.edit_internal_code.textChanged.connect(self._on_internal_code_changed)
        phase1_layout.addWidget(self.edit_internal_code, 2, 1)

        # 产品名称（可编辑）
        phase1_layout.addWidget(QLabel("产品名称:"), 2, 2)
        self.edit_product_name = QLineEdit("-")
        self.edit_product_name.textChanged.connect(self._on_product_name_changed)
        phase1_layout.addWidget(self.edit_product_name, 2, 3)

        # H.S.Code（只读）
        phase1_layout.addWidget(QLabel("H.S.Code:"), 3, 0)
        self.lbl_hs_code = QLabel("-")
        phase1_layout.addWidget(self.lbl_hs_code, 3, 1)

        # 卸货港（只读）
        phase1_layout.addWidget(QLabel("卸货港:"), 3, 2)
        self.lbl_destination = QLabel("-")
        phase1_layout.addWidget(self.lbl_destination, 3, 3)

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

        # 加载鉴定报告按钮
        self.btn_load_cargo = QPushButton("加载鉴定报告")
        self.btn_load_cargo.clicked.connect(self._load_cargo_report)
        cargo_layout.addWidget(self.btn_load_cargo, 3, 0, 1, 2)

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
        # 使用Stretch模式让列宽自适应填满
        self.msds_components_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.msds_components_table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
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

        btn_layout.addStretch()

        layout.addLayout(btn_layout)

        return panel

    def _setup_booking_tab(self):
        """设置订舱单Tab"""
        layout = QVBoxLayout()
        self.booking_tab.setLayout(layout)

        # 订舱单预览表格（显示Excel模板内容）
        self.booking_table = QTableWidget()
        self.booking_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.booking_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.booking_table.setFont(QFont("Microsoft YaHei", 9))
        layout.addWidget(self.booking_table)

        # 底部按钮
        btn_layout = QHBoxLayout()
        self.btn_load_template = QPushButton("加载模板")
        self.btn_load_template.clicked.connect(self._load_booking_template)
        btn_layout.addWidget(self.btn_load_template)

        self.btn_fill_data = QPushButton("填充数据")
        self.btn_fill_data.clicked.connect(self._fill_booking_from_shipment)
        self.btn_fill_data.setEnabled(False)
        btn_layout.addWidget(self.btn_fill_data)

        btn_layout.addStretch()
        layout.addLayout(btn_layout)

    def _load_booking_template(self):
        """加载BOOKING模板Excel"""
        template_path = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            "core", "phase2", "长晟出口海运BOOKING模板.xls"
        )

        if not os.path.exists(template_path):
            QMessageBox.warning(self, "失败", f"模板文件不存在:\n{template_path}")
            return

        try:
            import xlrd
            wb = xlrd.open_workbook(template_path)
            sh = wb.sheet_by_index(0)

            # 设置表格行列数
            self.booking_table.setRowCount(sh.nrows)
            self.booking_table.setColumnCount(sh.ncols)

            # 填充数据
            for r in range(sh.nrows):
                for c in range(sh.ncols):
                    cell_value = sh.cell_value(r, c)
                    item = QTableWidgetItem(str(cell_value) if cell_value else "")
                    self.booking_table.setItem(r, c, item)

            # 调整列宽
            self.booking_table.resizeColumnsToContents()

            self.btn_fill_data.setEnabled(True)
            QMessageBox.information(self, "成功", "模板已加载")

        except Exception as e:
            QMessageBox.warning(self, "失败", f"加载模板失败:\n{str(e)}")

    def _fill_booking_from_shipment(self):
        """用Shipment数据填充表格"""
        if not self.shipment:
            QMessageBox.warning(self, "警告", "请先导入数据")
            return

        # 根据模板结构填充数据
        # Row 1: Shipper发货人
        # Row 8: Consignee收货人
        # Row 13: Notify通知人
        # Row 24: Port of loading
        # Row 27: Port of Discharge

        for r in range(self.booking_table.rowCount()):
            for c in range(self.booking_table.columnCount()):
                item = self.booking_table.item(r, c)
                if not item:
                    continue
                cell_text = item.text()

                # 发货人
                if 'Shipper' in cell_text or '发货人' in cell_text:
                    if self.shipment.shipper:
                        self.booking_table.setItem(r, c + 1, QTableWidgetItem(self.shipment.shipper))

                # 收货人
                if 'Consignee' in cell_text or '收货人' in cell_text:
                    if self.shipment.consignee:
                        self.booking_table.setItem(r, c + 1, QTableWidgetItem(self.shipment.consignee))

                # 通知人
                if 'Notify' in cell_text or '通知人' in cell_text:
                    if self.shipment.notifier:
                        self.booking_table.setItem(r, c + 1, QTableWidgetItem(self.shipment.notifier))

                # 卸货港
                if 'Port of' in cell_text and 'Discharge' in cell_text:
                    if self.shipment.destination:
                        self.booking_table.setItem(r, c + 1, QTableWidgetItem(self.shipment.destination))

        # 填充货物信息（Row 31-32附近）
        for r in range(self.booking_table.rowCount()):
            item = self.booking_table.item(r, 0)
            if item and ('Marks' in item.text() or '货物' in item.text()):
                # 填充品名
                if self.shipment.product_name_cn:
                    self.booking_table.setItem(r + 1, 2, QTableWidgetItem(self.shipment.product_name_cn))
                # 填充数量
                if self.shipment.package_info:
                    pkg = self.shipment.package_info
                    if pkg.quantity_kg:
                        self.booking_table.setItem(r + 1, 5, QTableWidgetItem(f"{int(pkg.quantity_kg)} KG"))
                    if pkg.gross_weight_kg:
                        self.booking_table.setItem(r + 1, 7, QTableWidgetItem(f"{pkg.gross_weight_kg} KG"))
                break

        QMessageBox.information(self, "成功", "数据已填充")

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
        # 不使用硬编码路径，让用户选择文件
        filepath, _ = QFileDialog.getOpenFileName(
            self, "选择出口商品编码Excel", "",
            "Excel Files (*.xlsx *.xls);;All Files (*)"
        )

        if filepath:
            if self.data_merger.load_export_codes(filepath):
                QMessageBox.information(self, "成功", f"加载成功，共 {self.data_merger.export_codes_loader.get_count()} 条数据")
            else:
                QMessageBox.warning(self, "失败", "加载出口商品编码失败")

    def _load_cargo_report(self):
        """加载运输鉴定报告"""
        filepath, _ = QFileDialog.getOpenFileName(
            self, "选择运输鉴定报告", "",
            "Excel Files (*.xlsx *.xls);;All Files (*)"
        )

        if filepath:
            try:
                # 使用ReportParser解析鉴定报告
                cargo = self.report_parser.parse_file(filepath)
                if cargo:
                    self.shipment.cargo = cargo
                    self._update_cargo_ui()
                    QMessageBox.information(self, "成功", "鉴定报告已加载")
                else:
                    QMessageBox.warning(self, "失败", "无法解析鉴定报告文件")
            except Exception as e:
                QMessageBox.warning(self, "失败", f"加载鉴定报告失败:\n{str(e)}")

    def _update_phase1_ui(self):
        """更新Phase 1数据UI"""
        if not self.shipment:
            return

        self.edit_shipper.setText(self.shipment.shipper or "-")
        self.edit_consignee.setText(self.shipment.consignee or "-")
        self.edit_notifier.setText(self.shipment.notifier or "-")
        self.edit_pi_no.setText(self.shipment.pi_no or "-")
        self.edit_internal_code.setText(self.shipment.internal_code or "-")
        self.edit_product_name.setText(self.shipment.product_name_cn or "-")

        # 创建PI号对应的输出文件夹
        self._create_output_folder()
        self.lbl_hs_code.setText(self.shipment.hs_code or "-")
        self.lbl_destination.setText(self.shipment.destination or "-")

        # 自动匹配出口商品编码
        if self.shipment.internal_code and self.data_merger.export_codes_loader.is_loaded():
            self.data_merger.assign_export_code(self.shipment)
            self._update_export_code_ui()

        # 启用生成按钮
        self._enable_generate_buttons()

    def _on_shipper_changed(self, text):
        if self.shipment:
            self.shipment.shipper = text

    def _on_consignee_changed(self, text):
        if self.shipment:
            self.shipment.consignee = text

    def _on_notifier_changed(self, text):
        if self.shipment:
            self.shipment.notifier = text

    def _on_pi_no_changed(self, text):
        if self.shipment:
            self.shipment.pi_no = text

    def _on_internal_code_changed(self, text):
        if self.shipment:
            self.shipment.internal_code = text

    def _on_product_name_changed(self, text):
        if self.shipment:
            self.shipment.product_name_cn = text

    def _update_export_code_ui(self):
        """更新出口商品编码UI"""
        if not self.shipment or not self.shipment.export_code:
            return

        ec = self.shipment.export_code
        self.lbl_export_hs.setText(f"海关编码: {ec.hs_code}")
        self.lbl_export_customs.setText(f"报关名称: {ec.customs_name}")
        self.lbl_export_composition.setText(f"报关成分: {ec.composition}")

    def _create_output_folder(self):
        """创建PI号对应的输出文件夹"""
        if not self.shipment or not self.shipment.pi_no:
            return

        pi_no = self.shipment.pi_no
        if not pi_no or pi_no == "-":
            return

        # 桌面路径作为基础
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        self.output_dir = os.path.join(desktop, f"订舱文件_{pi_no}")

        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)

        # 显示在父窗口的状态栏
        if self.parent() and hasattr(self.parent(), 'statusBar'):
            self.parent().statusBar().showMessage(f"文件输出目录: {self.output_dir}", 3000)

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

    def _generate_booking(self):
        """生成订舱单"""
        if not self.shipment:
            QMessageBox.warning(self, "警告", "请先导入数据")
            return

        try:
            import xlwt
            from xlwt import Workbook

            output_dir = self.output_dir if self.output_dir else os.path.join(os.path.expanduser("~"), "Desktop")
            pi_no = self.shipment.pi_no or "Booking"
            filename = f"订舱单-{pi_no}.xls"
            output_path = os.path.join(output_dir, filename)

            # 从表格创建Excel
            wb = Workbook()
            ws = wb.add_sheet('订舱单')

            for r in range(self.booking_table.rowCount()):
                for c in range(self.booking_table.columnCount()):
                    item = self.booking_table.item(r, c)
                    if item:
                        ws.write(r, c, item.text())

            wb.save(output_path)
            QMessageBox.information(self, "成功", f"订舱单已保存到:\n{output_path}")

        except ImportError:
            QMessageBox.warning(self, "失败", "需要安装xlwt库: pip install xlwt")
        except Exception as e:
            QMessageBox.warning(self, "失败", f"生成订舱单失败:\n{str(e)}")

    def _generate_msds(self):
        """生成MSDS"""
        if not self.shipment:
            QMessageBox.warning(self, "警告", "请先导入数据")
            return

        try:
            generator = MSDSGenerator()
            output_dir = self.output_dir if self.output_dir else os.path.join(os.path.expanduser("~"), "Desktop")

            cn_path, en_path = generator.generate_both(self.shipment, output_dir)

            # 预览
            preview_text = f"中文MSDS: {cn_path}\n英文MSDS: {en_path}"
            self.msds_preview.setPlainText(preview_text)

            QMessageBox.information(self, "成功", f"MSDS已保存到:\n{output_dir}")

        except Exception as e:
            QMessageBox.warning(self, "失败", f"生成MSDS失败:\n{str(e)}")

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