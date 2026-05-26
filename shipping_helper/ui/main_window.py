# shipping_helper/ui/main_window.py
# -*- coding: utf-8 -*-
"""
主窗口 - Phase 1 左右分栏布局
"""

from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QTextEdit, QPushButton, QLabel, QMessageBox,
                             QTableWidget, QTableWidgetItem, QScrollArea,
                             QGroupBox, QLineEdit, QApplication, QHeaderView,
                             QComboBox, QCheckBox)
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
        self.merged_data = {}
        self.package_result = {}
        self._current_order_qty = 0
        self._current_order_req = ''

        self._load_knowledge()
        self._init_ui()

    def _init_ui(self):
        """初始化UI"""
        self.setWindowTitle("ShippingHelper - Phase 1: 订单数据提取与包装计算")
        self.setGeometry(100, 100, 1200, 700)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QHBoxLayout()
        central_widget.setLayout(main_layout)

        left_widget = self._create_input_panel()
        main_layout.addWidget(left_widget, 2)  # 左侧占比较小

        right_widget = self._create_result_panel()
        main_layout.addWidget(right_widget, 3)  # 右侧占比较大

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
        self.order_text_edit.setMinimumHeight(200)
        self.order_text_edit.setFont(QFont("Microsoft YaHei", 10))
        self.order_text_edit.setLineWrapMode(QTextEdit.WidgetWidth)  # 自动换行
        layout.addWidget(self.order_text_edit)

        layout.addWidget(QLabel("PI文件 (.xls):"))
        pi_layout = QHBoxLayout()
        self.pi_path_edit = QLineEdit()
        self.pi_path_edit.setPlaceholderText("选择PI文件路径...")
        self.btn_select_pi = QPushButton("选择文件")
        self.btn_select_pi.clicked.connect(self.select_pi_file)
        pi_layout.addWidget(self.pi_path_edit)
        pi_layout.addWidget(self.btn_select_pi)
        layout.addLayout(pi_layout)

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
        self.fields_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.fields_table.setRowCount(23)
        self.fields_table.setWordWrap(True)
        self.fields_table.cellClicked.connect(self.copy_cell_value)
        # 设置默认行高
        for i in range(23):
            self.fields_table.setRowHeight(i, 25)
        fields_layout.addWidget(self.fields_table)

        # 右移按钮：当订单要求字段缺失时使用
        shift_layout = QHBoxLayout()
        self.btn_shift_right = QPushButton("右移（订单要求缺失时使用）")
        self.btn_shift_right.clicked.connect(self._shift_values_right)
        shift_layout.addWidget(self.btn_shift_right)
        shift_layout.addStretch()
        fields_layout.addLayout(shift_layout)

        fields_group.setLayout(fields_layout)
        scroll_layout.addWidget(fields_group)

        package_group = QGroupBox("包装计算")
        package_layout = QVBoxLayout()

        # 桶类型选择
        drum_row = QHBoxLayout()
        drum_row.addWidget(QLabel("桶类型:"))
        self.drum_combo = QComboBox()
        self.drum_combo.addItems(self.package_calculator.get_package_options())
        self.drum_combo.currentTextChanged.connect(self._on_package_changed)
        drum_row.addWidget(self.drum_combo)
        drum_row.addStretch()
        package_layout.addLayout(drum_row)

        # 卡板选择
        pallet_row = QHBoxLayout()
        pallet_row.addWidget(QLabel("卡板:"))
        self.pallet_combo = QComboBox()
        self.pallet_combo.addItems(self.package_calculator.get_pallet_options())
        self.pallet_combo.currentTextChanged.connect(self._on_package_changed)
        pallet_row.addWidget(self.pallet_combo)
        pallet_row.addStretch()
        package_layout.addLayout(pallet_row)

        # 不打卡板选项
        self.no_pallet_check = QCheckBox("不打卡板")
        self.no_pallet_check.stateChanged.connect(self._on_package_changed)
        package_layout.addWidget(self.no_pallet_check)

        # 计算结果标签
        self.package_result_label = QLabel("请先粘贴订单数据")
        self.package_result_label.setWordWrap(True)
        self.package_result_label.setFont(QFont("Microsoft YaHei", 10))
        package_layout.addWidget(self.package_result_label)

        # 重新计算按钮
        self.btn_recalc = QPushButton("重新计算包装")
        self.btn_recalc.clicked.connect(self._recalculate_package)
        package_layout.addWidget(self.btn_recalc)

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
            # 延迟更新label，等UI初始化完成
            if hasattr(self, 'knowledge_label'):
                self.knowledge_label.setText(f"知识库状态：已加载 (v{version})")
        else:
            if hasattr(self, 'knowledge_label'):
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
        self._current_order_qty = order_qty
        self._current_order_req = order_req

        # 智能推荐桶类型
        if order_req and '粉' in order_req:
            # 粉类优先用纸桶或编织袋
            for i, name in enumerate(self.drum_combo.currentText()):
                if '纸桶' in name or '编织' in name or '25kg' in name:
                    self.drum_combo.setCurrentIndex(i)
                    break

        self._recalculate_package()

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
            self.fields_table.setItem(i, 1, item)

    def _shift_values_right(self):
        """右移值：从订单要求开始，将值往后移一位"""
        # 从后往前移，避免覆盖
        for i in range(21, 9, -1):
            item = self.fields_table.item(i, 1)
            value = item.text() if item else ''
            new_item = QTableWidgetItem(value)
            self.fields_table.setItem(i + 1, 1, new_item)

        # 订单要求位置清空
        self.fields_table.setItem(10, 1, QTableWidgetItem(''))
        self.statusBar().showMessage("已右移，请检查并手动修正", 3000)

    def copy_cell_value(self, row, col):
        """点击单元格复制值"""
        if col == 1:
            item = self.fields_table.item(row, col)
            if item:
                clipboard = QApplication.clipboard()
                clipboard.setText(item.text())
                self.statusBar().showMessage(f"已复制: {item.text()}", 2000)

    def _on_package_changed(self):
        """当包装选项变化时重新计算"""
        if hasattr(self, '_current_order_qty') and self._current_order_qty > 0:
            self._recalculate_package()

    def _recalculate_package(self):
        """重新计算包装"""
        if not hasattr(self, '_current_order_qty') or self._current_order_qty <= 0:
            self.package_result_label.setText("无效的订单量")
            return

        drum_type = self.drum_combo.currentText()
        pallet_type = self.pallet_combo.currentText()
        total_qty = self._current_order_qty

        if self.no_pallet_check.isChecked():
            result = self.package_calculator.calculate_no_pallet(drum_type, total_qty)
        else:
            result = self.package_calculator.calculate_with_pallet(drum_type, pallet_type, total_qty)

        self.package_result = result

        if 'error' in result:
            self.package_result_label.setText(f"错误: {result['error']}")
        else:
            res = result.get('result', {})
            container = result.get('container_fit', {})
            text = f"""
桶类型: {drum_type}
桶数: {res.get('total_drums', 0)}
卡板数: {'无' if self.no_pallet_check.isChecked() else res.get('total_pallets', 0)}
产品净重: {res.get('product_weight_kg', 0)} kg
毛重: {res.get('gross_weight_kg', 0)} kg
总体积: {res.get('total_volume_cbm', 0)} CBM
{'能装入20GP' if container.get('fits_20gp') else '不能装入20GP'}
"""
            self.package_result_label.setText(text.strip())

    def enter_phase2(self):
        """进入Phase 2（预留接口）"""
        QMessageBox.information(
            self, "Phase 2",
            f"合并数据: {len(self.merged_data)} 字段\n"
            f"包装结果: {self.package_result}\n\n"
            "Phase 2 功能待开发"
        )