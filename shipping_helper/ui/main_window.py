# shipping_helper/ui/main_window.py
# -*- coding: utf-8 -*-
"""
主窗口 - Phase 1 左右分栏布局
"""

from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QTextEdit, QPushButton, QLabel, QMessageBox,
                             QTableWidget, QTableWidgetItem, QScrollArea,
                             QGroupBox, QLineEdit, QApplication)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont, QClipboard
import json
import os

from core.order_parser import OrderParser
from core.pi_extractor import PIExtractor
from core.code_matcher import CodeMatcher
from core.package_calculator import PackageCalculator


class MainWindow(QMainWindow):
    """主窗口"""

    def __init__(self):
        super().__init__()
        self.order_parser = OrderParser()
        self.pi_extractor = PIExtractor()
        self.code_matcher = None
        self.package_calculator = None
        self.merged_data = {}
        self.package_result = {}

        self._init_ui()
        self._load_knowledge()

    def _init_ui(self):
        """初始化UI"""
        self.setWindowTitle("ShippingHelper - Phase 1: 订单数据提取与包装计算")
        self.setGeometry(100, 100, 1200, 700)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QHBoxLayout()
        central_widget.setLayout(main_layout)

        left_widget = self._create_input_panel()
        main_layout.addWidget(left_widget, 1)

        right_widget = self._create_result_panel()
        main_layout.addWidget(right_widget, 1)

        bottom_layout = QHBoxLayout()
        self.btn_phase2 = QPushButton("进入 Phase 2")
        self.btn_phase2.setEnabled(False)
        self.btn_phase2.clicked.connect(self.enter_phase2)
        bottom_layout.addStretch()
        bottom_layout.addWidget(self.btn_phase2)

        main_layout.addLayout(bottom_layout)

    def _create_input_panel(self) -> QWidget:
        """创建输入面板"""
        panel = QGroupBox("数据输入")
        layout = QVBoxLayout()

        layout.addWidget(QLabel("外贸销售订单表粘贴:"))
        self.order_text_edit = QTextEdit()
        self.order_text_edit.setPlaceholderText("从在线表格复制一行数据，粘贴至此...")
        self.order_text_edit.setMinimumHeight(150)
        layout.addWidget(self.order_text_edit)

        layout.addWidget(QLabel("PI文件 (.xls):"))
        pi_layout = QHBoxLayout()
        self.pi_path_edit = QLineEdit()
        self.pi_path_edit.setPlaceholderText("选择PI文件路径...")
        self.btn_select_pi = QPushButton("选择文件")
        self.btn_select_pi.clicked.connect(self.select_pi_file)
        pi_layout.addWidget(self.pi_path_edit)
        pi_layout.addWidget(self.btn_select_pi)
        layout.addWidget(pi_layout)

        self.btn_calculate = QPushButton("开始计算")
        self.btn_calculate.clicked.connect(self.calculate)
        layout.addWidget(self.btn_calculate)

        self.knowledge_label = QLabel("知识库状态：未加载")
        layout.addWidget(self.knowledge_label)

        panel.setLayout(layout)
        return panel

    def _create_result_panel(self) -> QWidget:
        """创建结果面板"""
        panel = QGroupBox("结果展示")
        layout = QVBoxLayout()

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout()

        fields_group = QGroupBox("合并字段结果")
        fields_layout = QVBoxLayout()
        self.fields_table = QTableWidget()
        self.fields_table.setColumnCount(2)
        self.fields_table.setHorizontalHeaderLabels(["字段名", "值"])
        self.fields_table.horizontalHeader().setStretchToSection(True)
        self.fields_table.setRowCount(23)
        self.fields_table.cellClicked.connect(self.copy_cell_value)
        fields_layout.addWidget(self.fields_table)
        fields_group.setLayout(fields_layout)
        scroll_layout.addWidget(fields_group)

        package_group = QGroupBox("包装计算结果")
        package_layout = QVBoxLayout()
        self.package_table = QTableWidget()
        self.package_table.setColumnCount(2)
        self.package_table.setHorizontalHeaderLabels(["项目", "值"])
        self.package_table.setRowCount(8)
        self.package_table.cellClicked.connect(self.copy_cell_value)
        package_layout.addWidget(self.package_table)
        package_group.setLayout(package_layout)
        scroll_layout.addWidget(package_group)

        scroll_widget.setLayout(scroll_layout)
        scroll.setWidget(scroll_widget)
        layout.addWidget(scroll)

        panel.setLayout(layout)
        return panel

    def _load_knowledge(self):
        """加载知识库"""
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        products_file = os.path.join(base_dir, 'knowledge', 'products_knowledge.json')
        packaging_file = os.path.join(base_dir, 'knowledge', 'packaging_data.json')

        self.code_matcher = CodeMatcher(products_file)
        if self.code_matcher.load(products_file):
            version = self.code_matcher.get_version()
            self.knowledge_label.setText(f"知识库状态：已加载 (v{version})")
        else:
            self.knowledge_label.setText("知识库状态：加载失败")

        self.package_calculator = PackageCalculator(packaging_file)

    def select_pi_file(self):
        """选择PI文件"""
        from PyQt5.QtWidgets import QFileDialog
        filepath, _ = QFileDialog.getOpenFileName(
            self, "选择PI文件", "", "Excel Files (*.xls);;All Files (*)"
        )
        if filepath:
            self.pi_path_edit.setText(filepath)

    def calculate(self):
        """执行计算"""
        order_text = self.order_text_edit.toPlainText().strip()
        if not order_text:
            QMessageBox.warning(self, "警告", "请粘贴订单表数据")
            return

        order_data = self.order_parser.parse(order_text)
        is_valid, msg = self.order_parser.validate()
        if not is_valid:
            QMessageBox.warning(self, "警告", msg)
            return

        pi_path = self.pi_path_edit.text().strip()
        if pi_path and os.path.exists(pi_path):
            pi_data = self.pi_extractor.parse_file(pi_path)
        else:
            pi_data = {}

        internal_code = self.order_parser.get_internal_code()
        if internal_code:
            code_info = self.code_matcher.get_all_info(internal_code)
        else:
            code_info = {}

        self.merged_data = self._build_merged_data(order_data, pi_data, code_info)

        order_req = order_data.get('订单要求', '')
        order_qty = float(order_data.get('订单量kg', 0) or 0)
        self.package_result = self._calculate_package(order_req, order_qty)

        self._update_result_ui()

        self.btn_phase2.setEnabled(True)

    def _build_merged_data(self, order_data, pi_data, code_info) -> dict:
        """构建合并数据"""
        result = {}
        for key in ['业务员', '客户编号', '内部编号', '产品中文名', '报关名称',
                    '规格kg', '订单量kg', '是否调价', '有无样品', '订单要求',
                    '交货日期', '审核', '销售区域', '订单号', '出货抬头',
                    '单据类型', '跟单员', '下单日期', '确认下单', '生产交期',
                    '出货渠道', '出货方式', '规格异常']:
            result[key] = order_data.get(key, '')

        result['收货人'] = pi_data.get('收货人', '')
        result['品名英文'] = pi_data.get('品名英文', '')
        result['数量'] = pi_data.get('数量', '')
        result['单价'] = pi_data.get('单价', '')
        result['金额'] = pi_data.get('金额', '')
        result['H.S.Code'] = pi_data.get('H.S.Code', '')
        result['卸货港'] = pi_data.get('卸货港', '')

        if not result.get('报关名称'):
            result['报关名称'] = code_info.get('报关名称', '')
        result['海关编码'] = code_info.get('海关编码', '')
        result['报关成分'] = code_info.get('报关成分', '')

        return result

    def _calculate_package(self, order_req: str, total_qty: float) -> dict:
        """计算包装"""
        if total_qty <= 0:
            return {'error': '订单量无效'}

        drum_type = "50kg蓝桶(细口)"
        pallet_type = "1.1*1.1m卡板"

        result = self.package_calculator.calculate_with_pallet(
            drum_type, pallet_type, total_qty
        )
        return result

    def _update_result_ui(self):
        """更新结果UI"""
        fields = [
            ('业务员', self.merged_data.get('业务员', '')),
            ('客户编号', self.merged_data.get('客户编号', '')),
            ('内部编号', self.merged_data.get('内部编号', '')),
            ('产品中文名', self.merged_data.get('产品中文名', '')),
            ('报关名称', self.merged_data.get('报关名称', '')),
            ('规格kg', self.merged_data.get('规格kg', '')),
            ('订单量kg', self.merged_data.get('订单量kg', '')),
            ('是否调价', self.merged_data.get('是否调价', '')),
            ('有无样品', self.merged_data.get('有无样品', '')),
            ('订单要求', self.merged_data.get('订单要求', '')),
            ('交货日期', self.merged_data.get('交货日期', '')),
            ('审核', self.merged_data.get('审核', '')),
            ('销售区域', self.merged_data.get('销售区域', '')),
            ('订单号', self.merged_data.get('订单号', '')),
            ('出货抬头', self.merged_data.get('出货抬头', '')),
            ('单据类型', self.merged_data.get('单据类型', '')),
            ('跟单员', self.merged_data.get('跟单员', '')),
            ('下单日期', self.merged_data.get('下单日期', '')),
            ('确认下单', self.merged_data.get('确认下单', '')),
            ('生产交期', self.merged_data.get('生产交期', '')),
            ('出货渠道', self.merged_data.get('出货渠道', '')),
            ('出货方式', self.merged_data.get('出货方式', '')),
            ('规格异常', self.merged_data.get('规格异常', '')),
        ]

        for i, (name, value) in enumerate(fields):
            self.fields_table.setItem(i, 0, QTableWidgetItem(name))
            item = QTableWidgetItem(str(value))
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            self.fields_table.setItem(i, 1, item)

        pkg = self.package_result
        if 'error' in pkg:
            package_items = [('错误', pkg['error'])]
        else:
            res = pkg.get('result', {})
            package_items = [
                ('桶类型', pkg.get('input', {}).get('drum_type', '')),
                ('桶数', res.get('total_drums', '')),
                ('卡板类型', pkg.get('input', {}).get('pallet_type', '')),
                ('卡板数', res.get('total_pallets', '')),
                ('产品净重(kg)', res.get('product_weight_kg', '')),
                ('毛重(kg)', res.get('gross_weight_kg', '')),
                ('总体积(CBM)', res.get('total_volume_cbm', '')),
                ('20GP装裁', '是' if pkg.get('container_fit', {}).get('fits_20gp') else '否'),
            ]

        for i, (name, value) in enumerate(package_items):
            self.package_table.setItem(i, 0, QTableWidgetItem(name))
            item = QTableWidgetItem(str(value))
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            self.package_table.setItem(i, 1, item)

    def copy_cell_value(self, row, col):
        """点击单元格复制值"""
        if col == 1:
            item = self.fields_table.item(row, col) or self.package_table.item(row, col)
            if item:
                clipboard = QApplication.clipboard()
                clipboard.setText(item.text())
                self.statusBar().showMessage(f"已复制: {item.text()}", 2000)

    def enter_phase2(self):
        """进入Phase 2（预留接口）"""
        QMessageBox.information(
            self, "Phase 2",
            f"合并数据: {len(self.merged_data)} 字段\n"
            f"包装结果: {self.package_result}\n\n"
            "Phase 2 功能待开发"
        )