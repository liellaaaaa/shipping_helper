# shipping_helper/ui/main_window.py
# -*- coding: utf-8 -*-
"""
主窗口 - Phase 1 左右分栏布局
"""

from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QTextEdit, QPushButton, QLabel, QMessageBox,
                             QTableWidget, QTableWidgetItem, QScrollArea,
                             QGroupBox, QLineEdit, QApplication, QHeaderView,
                             QComboBox, QCheckBox, QGridLayout)
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
        self._suppress_text_changed = False  # 防止setPlainText触发textChanged死循环

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
        panel = QWidget()
        layout = QVBoxLayout()
        panel.setLayout(layout)

        order_label = QLabel("外贸销售订单表粘贴:")
        order_label.setStyleSheet("font-weight: bold;")
        layout.addWidget(order_label)

        self.order_text_edit = QTextEdit()
        self.order_text_edit.setPlaceholderText("从在线表格复制一行数据，粘贴至此...")
        self.order_text_edit.setMinimumHeight(120)
        self.order_text_edit.setFont(QFont("Microsoft YaHei", 10))
        self.order_text_edit.setLineWrapMode(QTextEdit.WidgetWidth)
        self.order_text_edit.setMaximumWidth(400)
        self.order_text_edit.textChanged.connect(self._on_order_text_changed)
        layout.addWidget(self.order_text_edit)

        pi_label = QLabel("PI文件 (.xls):")
        pi_label.setStyleSheet("font-weight: bold; margin-top: 10px;")
        layout.addWidget(pi_label)

        pi_layout = QHBoxLayout()
        self.pi_path_edit = QLineEdit()
        self.pi_path_edit.setPlaceholderText("选择PI文件路径...")
        self.btn_select_pi = QPushButton("选择文件")
        self.btn_select_pi.clicked.connect(self.select_pi_file)
        pi_layout.addWidget(self.pi_path_edit)
        pi_layout.addWidget(self.btn_select_pi)
        layout.addLayout(pi_layout)

        self.btn_calculate = QPushButton("开始解析")
        self.btn_calculate.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-weight: bold;
                padding: 8px;
                border-radius: 4px;
            }
            QPushButton:hover { background-color: #45a049; }
        """)
        self.btn_calculate.clicked.connect(self.calculate)
        layout.addWidget(self.btn_calculate)

        # 解析结果预览区域（两个预览合并）
        preview_label = QLabel("解析结果预览:")
        preview_label.setStyleSheet("font-weight: bold; margin-top: 15px;")
        layout.addWidget(preview_label)

        # 预览区域使用QWidget包含两个GroupBox
        self.preview_container = QWidget()
        preview_main_layout = QVBoxLayout()
        self.preview_container.setLayout(preview_main_layout)

        # 订单预览
        order_preview_group = QGroupBox("销售订单表解析结果")
        order_preview_layout = QGridLayout()
        order_preview_layout.setSpacing(5)
        self.order_preview_labels = {}
        order_preview_fields = [
            ('业务员', 0, 0), ('客户编号', 0, 2),
            ('内部编号', 1, 0), ('产品中文名', 1, 2),
            ('报关名称', 2, 0), ('规格kg', 2, 2),
            ('订单量kg', 3, 0), ('订单要求', 3, 2),
            ('交货日期', 4, 0), ('订单号', 4, 2),
        ]
        for field, row, col in order_preview_fields:
            lbl = QLineEdit("-")
            lbl.setStyleSheet("background-color: #f5f5f5; padding: 2px; border-radius: 3px;")
            lbl.setReadOnly(False)
            order_preview_layout.addWidget(QLabel(f"{field}:"), row, col)
            order_preview_layout.addWidget(lbl, row, col + 1)
            self.order_preview_labels[field] = lbl
        order_preview_group.setLayout(order_preview_layout)
        preview_main_layout.addWidget(order_preview_group)

        # PI预览
        pi_preview_group = QGroupBox("PI文件解析结果")
        pi_preview_layout = QGridLayout()
        pi_preview_layout.setSpacing(5)
        self.pi_preview_labels = {}
        pi_preview_fields = [
            ('收货人', 0, 0), ('收货人地址', 0, 2),
            ('日期', 1, 0), ('PI号', 1, 2),
            ('品名英文', 2, 0), ('数量', 2, 2),
            ('单价', 3, 0), ('金额', 3, 2),
            ('H.S.Code', 4, 0), ('卸货港', 4, 2),
        ]
        for field, row, col in pi_preview_fields:
            lbl = QLineEdit("-")
            lbl.setStyleSheet("background-color: #f5f5f5; padding: 2px; border-radius: 3px;")
            lbl.setReadOnly(False)
            pi_preview_layout.addWidget(QLabel(f"{field}:"), row, col)
            pi_preview_layout.addWidget(lbl, row, col + 1)
            self.pi_preview_labels[field] = lbl
        pi_preview_group.setLayout(pi_preview_layout)
        preview_main_layout.addWidget(pi_preview_group)

        layout.addWidget(self.preview_container)

        self.knowledge_label = QLabel("知识库状态：未加载")
        layout.addWidget(self.knowledge_label)

        layout.addStretch()
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
        # 第一列(字段名)固定宽度100，第二列(值)固定宽度300，超出自动换行
        self.fields_table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.fields_table.verticalHeader().setDefaultSectionSize(30)  # 默认行高30
        self.fields_table.setWordWrap(True)
        self.fields_table.cellClicked.connect(self.copy_cell_value)
        self.fields_table.setColumnWidth(0, 100)
        self.fields_table.setColumnWidth(1, 300)
        self.fields_table.setRowCount(23)
        # 让行高自动适应内容
        self.fields_table.resizeRowsToContents()
        fields_layout.addWidget(self.fields_table)

        # 移动按钮：根据选中行移动
        shift_layout = QHBoxLayout()
        self.btn_shift_up = QPushButton("上移")
        self.btn_shift_up.clicked.connect(self._shift_values_up)
        shift_layout.addWidget(self.btn_shift_up)

        self.btn_shift_down = QPushButton("下移")
        self.btn_shift_down.clicked.connect(self._shift_values_down)
        shift_layout.addWidget(self.btn_shift_down)

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

    def _on_order_text_changed(self):
        """订单文本变化时，将tab替换为换行"""
        if self._suppress_text_changed:
            return
        text = self.order_text_edit.toPlainText()
        if '\t' in text:
            self._suppress_text_changed = True
            self.order_text_edit.setPlainText(text.replace('\t', '\n'))
            self._suppress_text_changed = False

    def _update_preview(self, order_data: dict, pi_data: dict):
        """更新预览区域（订单+PI）"""
        # 更新订单预览
        for field, lbl in self.order_preview_labels.items():
            value = order_data.get(field, '-')
            if isinstance(value, float):
                value = f"{value:.0f}"
            lbl.setText(str(value) if value else '-')

        # 更新PI预览
        for field, lbl in self.pi_preview_labels.items():
            value = pi_data.get(field, '-')
            if isinstance(value, float):
                value = f"{value:.2f}" if field in ['单价', '金额'] else f"{value:.0f}"
            lbl.setText(str(value) if value else '-')

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

        # 将Tab替换为换行，便于查看（但保留Tab用于解析）
        order_text_for_display = order_text.replace('\t', '\n')
        self.order_text_edit.setPlainText(order_text_for_display)

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

        # 更新预览
        self._update_preview(order_data, pi_data)

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
        self.fields_table.resizeRowsToContents()

    def _shift_values_down(self):
        """下移值：将选中行及之后的值往下移一位"""
        current_row = self.fields_table.currentRow()
        if current_row < 0:
            current_row = 10  # 默认从第10行开始
        if current_row >= 22:
            self.statusBar().showMessage("已到最后一行，无法继续下移", 3000)
            return

        # 保存选中行的值
        current_item = self.fields_table.item(current_row, 1)
        original_value = current_item.text() if current_item else ''

        # 从current_row开始到结尾，所有值上移一位（填到选中行）
        for i in range(current_row, 22):
            item = self.fields_table.item(i + 1, 1)
            value = item.text() if item else ''
            self.fields_table.setItem(i, 1, QTableWidgetItem(value))
        # 最后一行清空
        self.fields_table.setItem(22, 1, QTableWidgetItem(''))
        # 选中行的值放到下一行
        self.fields_table.setItem(current_row + 1, 1, QTableWidgetItem(original_value))
        self.fields_table.setCurrentCell(current_row + 1, 1)
        self.statusBar().showMessage(f"已将第{current_row+1}行下移", 3000)

    def _shift_values_up(self):
        """上移值：将选中行及之后的值往上移一位"""
        current_row = self.fields_table.currentRow()
        if current_row < 0:
            current_row = 10  # 默认从第10行开始
        if current_row <= 0:
            self.statusBar().showMessage("已到第一行，无法继续上移", 3000)
            return

        # 保存选中行的值
        current_item = self.fields_table.item(current_row, 1)
        original_value = current_item.text() if current_item else ''

        # 从第0行到current_row-1，所有值下移一位（填到选中行）
        for i in range(current_row - 1, -1, -1):
            item = self.fields_table.item(i, 1)
            value = item.text() if item else ''
            self.fields_table.setItem(i + 1, 1, QTableWidgetItem(value))
        # 第一行清空
        self.fields_table.setItem(0, 1, QTableWidgetItem(''))
        # 选中行的值放到上一行
        self.fields_table.setItem(current_row - 1, 1, QTableWidgetItem(original_value))
        self.fields_table.setCurrentCell(current_row - 1, 1)
        self.statusBar().showMessage(f"已将第{current_row+1}行上移", 3000)

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
            fits_20gp = container.get('fits_20gp', False)
            full_loads = container.get('full_20gp_loads', 0)

            if fits_20gp:
                container_text = f"能装入20GP，需要 {full_loads} 个货柜"
            else:
                container_text = "不能装入20GP"

            text = f"""
桶类型: {drum_type}
桶数: {res.get('total_drums', 0)}
卡板数: {'无' if self.no_pallet_check.isChecked() else res.get('total_pallets', 0)}
产品净重: {res.get('product_weight_kg', 0)} kg
毛重: {res.get('gross_weight_kg', 0)} kg
总体积: {res.get('total_volume_cbm', 0)} CBM
{container_text}
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