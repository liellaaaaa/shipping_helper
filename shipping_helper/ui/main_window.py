# -*- coding: utf-8 -*-
"""
主窗口 - Phase 1 左右分栏布局
"""

from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QTextEdit, QPushButton, QLabel, QMessageBox,
                             QTableWidget, QTableWidgetItem, QScrollArea,
                             QGroupBox, QLineEdit, QApplication, QHeaderView,
                             QComboBox, QCheckBox, QGridLayout, QSizePolicy,
                             QRadioButton, QButtonGroup, QTabWidget)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont, QClipboard
import json
import os

from core.order_parser import OrderParser
from core.pi_extractor import PIExtractor
from core.code_matcher import CodeMatcher
from core.package_calculator import PackageCalculator
from ui.booking_widget import BookingWidget


class MainWindow(QMainWindow):
    """主窗口 - Phase 1 & Phase 2 整合"""

    def __init__(self):
        super().__init__()
        self.order_parser = OrderParser()
        self.pi_extractor = PIExtractor()
        self.merged_data = {}
        self.package_result = {}
        self._current_order_qty = 0
        self._current_order_req = ''
        self._suppress_text_changed = False
        self._order_items = []  # 存储多个订单项
        self.booking_widget = None  # Phase 2 widget

        self._load_knowledge()
        self._init_ui()

    def _init_ui(self):
        """初始化UI"""
        self.setWindowTitle("ShippingHelper - Phase 1 & Phase 2 整合")
        self.setGeometry(100, 100, 1400, 800)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # 创建Tab组件
        self.main_tabs = QTabWidget()
        central_layout = QVBoxLayout()
        central_layout.addWidget(self.main_tabs)
        central_widget.setLayout(central_layout)

        # Tab 1: Phase 1 - 订单处理
        self.phase1_widget = self._create_phase1_widget()
        self.main_tabs.addTab(self.phase1_widget, "Phase 1 - 订单处理")

        # Tab 2: Phase 2 - 订舱出货
        self.booking_widget = BookingWidget(self, {})
        self.main_tabs.addTab(self.booking_widget, "Phase 2 - 订舱出货")

        # 初始化时自动填充每板数量
        self._auto_fill_per_pallet()

    def _create_phase1_widget(self) -> QWidget:
        """创建Phase 1主Widget"""
        widget = QWidget()
        layout = QHBoxLayout()
        widget.setLayout(layout)

        # 左侧：输入区 + 合并字段结果
        left_widget = self._create_left_panel()
        layout.addWidget(left_widget, 2)

        # 右侧：订单要求 + 包装计算
        right_widget = self._create_right_panel()
        layout.addWidget(right_widget, 3)

        return widget

    def _create_left_panel(self) -> QWidget:
        """创建左侧面板：输入区 + 合并字段结果"""
        panel = QWidget()
        layout = QVBoxLayout()
        panel.setLayout(layout)

        # ===== 输入区 =====
        input_group = QGroupBox("数据输入")
        input_layout = QVBoxLayout()

        order_label = QLabel("外贸销售订单表粘贴:")
        order_label.setStyleSheet("font-weight: bold;")
        input_layout.addWidget(order_label)

        self.order_text_edit = QTextEdit()
        self.order_text_edit.setPlaceholderText("从在线表格复制一行数据，粘贴至此...")
        self.order_text_edit.setMinimumHeight(80)
        self.order_text_edit.setFont(QFont("Microsoft YaHei", 10))
        self.order_text_edit.setLineWrapMode(QTextEdit.WidgetWidth)
        self.order_text_edit.textChanged.connect(self._on_order_text_changed)
        input_layout.addWidget(self.order_text_edit)

        pi_label = QLabel("PI文件 (.xls):")
        pi_label.setStyleSheet("font-weight: bold; margin-top: 8px;")
        input_layout.addWidget(pi_label)

        pi_layout = QHBoxLayout()
        self.pi_path_edit = QLineEdit()
        self.pi_path_edit.setPlaceholderText("选择PI文件路径...")
        self.btn_select_pi = QPushButton("选择文件")
        self.btn_select_pi.clicked.connect(self.select_pi_file)
        pi_layout.addWidget(self.pi_path_edit)
        pi_layout.addWidget(self.btn_select_pi)
        input_layout.addLayout(pi_layout)

        btn_row = QHBoxLayout()
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

        btn_row.addWidget(self.btn_calculate)
        btn_row.addStretch()
        input_layout.addLayout(btn_row)

        self.knowledge_label = QLabel("知识库状态：未加载")
        input_layout.addWidget(self.knowledge_label)

        input_group.setLayout(input_layout)
        layout.addWidget(input_group, 1)  # 输入区占1份

        # ===== 合并字段结果 =====
        self.fields_group = QGroupBox("合并字段结果")
        fields_layout = QVBoxLayout()

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setMinimumHeight(400)
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout()

        # 外贸销售订单表解析结果
        order_fields_group = QGroupBox("外贸销售订单表解析结果")
        order_fields_layout = QVBoxLayout()
        self.order_fields_table = QTableWidget()
        self.order_fields_table.setColumnCount(2)
        self.order_fields_table.setHorizontalHeaderLabels(["字段名", "值"])
        self.order_fields_table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.order_fields_table.verticalHeader().setDefaultSectionSize(25)
        self.order_fields_table.setWordWrap(True)
        self.order_fields_table.cellClicked.connect(self.copy_cell_value)
        self.order_fields_table.setColumnWidth(0, 100)
        self.order_fields_table.setColumnWidth(1, 320)
        self.order_fields_table.setRowCount(23)
        self.order_fields_table.resizeRowsToContents()
        order_fields_layout.addWidget(self.order_fields_table)
        order_fields_group.setLayout(order_fields_layout)
        scroll_layout.addWidget(order_fields_group)

        # PI文件解析结果
        pi_fields_group = QGroupBox("PI文件解析结果")
        pi_fields_layout = QVBoxLayout()
        self.pi_fields_table = QTableWidget()
        self.pi_fields_table.setColumnCount(2)
        self.pi_fields_table.setHorizontalHeaderLabels(["字段名", "值"])
        self.pi_fields_table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.pi_fields_table.verticalHeader().setDefaultSectionSize(25)
        self.pi_fields_table.setWordWrap(True)
        self.pi_fields_table.cellClicked.connect(self.copy_cell_value)
        self.pi_fields_table.setColumnWidth(0, 100)
        self.pi_fields_table.setColumnWidth(1, 320)
        self.pi_fields_table.setRowCount(12)
        self.pi_fields_table.resizeRowsToContents()
        pi_fields_layout.addWidget(self.pi_fields_table)
        pi_fields_group.setLayout(pi_fields_layout)
        scroll_layout.addWidget(pi_fields_group)

        # 上移/下移按钮
        shift_layout = QHBoxLayout()
        self.btn_shift_up = QPushButton("上移")
        self.btn_shift_up.clicked.connect(self._shift_values_up)
        shift_layout.addWidget(self.btn_shift_up)

        self.btn_shift_down = QPushButton("下移")
        self.btn_shift_down.clicked.connect(self._shift_values_down)
        shift_layout.addWidget(self.btn_shift_down)
        shift_layout.addStretch()
        scroll_layout.addLayout(shift_layout)

        scroll_widget.setLayout(scroll_layout)
        scroll.setWidget(scroll_widget)
        fields_layout.addWidget(scroll)

        self.fields_group.setLayout(fields_layout)
        layout.addWidget(self.fields_group, 4)  # 合并字段结果占4份
        layout.addStretch()

        return panel

    def _create_right_panel(self) -> QWidget:
        """创建右侧面板：订单要求 + 包装计算"""
        panel = QWidget()
        layout = QVBoxLayout()
        panel.setLayout(layout)

        # ===== 订单要求 =====
        req_group = QGroupBox("订单要求")
        req_layout = QVBoxLayout()

        self.order_req_edit = QTextEdit()
        self.order_req_edit.setPlaceholderText("订单要求内容（可编辑）...")
        self.order_req_edit.setMinimumHeight(150)
        self.order_req_edit.setFont(QFont("Microsoft YaHei", 10))
        self.order_req_edit.setLineWrapMode(QTextEdit.WidgetWidth)
        req_layout.addWidget(self.order_req_edit)

        req_group.setLayout(req_layout)
        layout.addWidget(req_group)

        # ===== 包装计算 =====
        package_group = QGroupBox("包装计算")
        package_layout = QVBoxLayout()

        # 包装类型选择
        pkg_type_row = QHBoxLayout()
        pkg_type_row.addWidget(QLabel("包装类型:"))
        self.pkg_type_combo = QComboBox()
        self.pkg_type_combo.addItems(self.package_calculator.get_package_options())
        self.pkg_type_combo.setStyleSheet("font-size: 13px; padding: 5px;")
        pkg_type_row.addWidget(self.pkg_type_combo)
        pkg_type_row.addStretch()
        package_layout.addLayout(pkg_type_row)

        # 数量输入
        qty_row = QHBoxLayout()
        qty_row.addWidget(QLabel("数量(kg):"))
        self.pkg_qty_edit = QLineEdit()
        self.pkg_qty_edit.setPlaceholderText("输入数量...")
        self.pkg_qty_edit.setStyleSheet("font-size: 13px; padding: 5px;")
        qty_row.addWidget(self.pkg_qty_edit)
        qty_row.addStretch()
        package_layout.addLayout(qty_row)

        # 卡板选择
        pallet_row = QHBoxLayout()
        pallet_row.addWidget(QLabel("卡板:"))
        self.pallet_combo = QComboBox()
        self.pallet_combo.addItems(self.package_calculator.get_pallet_options())
        self.pallet_combo.setStyleSheet("font-size: 13px; padding: 5px;")
        pallet_row.addWidget(self.pallet_combo)

        pallet_row.addSpacing(20)
        pallet_row.addWidget(QLabel("每板数量:"))
        self.per_pallet_edit = QLineEdit()
        self.per_pallet_edit.setPlaceholderText("留空用默认值")
        self.per_pallet_edit.setStyleSheet("font-size: 13px; padding: 5px; width: 60px;")
        pallet_row.addWidget(self.per_pallet_edit)
        pallet_row.addStretch()
        package_layout.addLayout(pallet_row)

        # 信号连接：选择包装类型或卡板类型时自动填充每板数量
        self.pkg_type_combo.currentIndexChanged.connect(self._auto_fill_per_pallet)
        self.pallet_combo.currentIndexChanged.connect(self._auto_fill_per_pallet)

        # 不打卡板选项
        pallet_options_row = QHBoxLayout()
        self.no_pallet_check = QCheckBox("不打卡板")
        pallet_options_row.addWidget(self.no_pallet_check)
        pallet_options_row.addSpacing(20)

        # 卡板版本选择
        pallet_options_row.addWidget(QLabel("卡板版本:"))
        self.pallet_version_group = QButtonGroup()
        self.pallet_version_new = QRadioButton("新卡板(27kg)")
        self.pallet_version_old = QRadioButton("旧卡板")
        self.pallet_version_new.setChecked(True)  # 默认选中新技术
        self.pallet_version_group.addButton(self.pallet_version_new)
        self.pallet_version_group.addButton(self.pallet_version_old)
        pallet_options_row.addWidget(self.pallet_version_new)
        pallet_options_row.addWidget(self.pallet_version_old)
        pallet_options_row.addStretch()
        package_layout.addLayout(pallet_options_row)

        # 信号连接：不打卡板时禁用卡板版本选择
        self.no_pallet_check.toggled.connect(self._on_no_pallet_toggled)

        # 包装计算表格（多产品支持）
        # 用滚动区域包裹表格，这样合计行可以固定在下方
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setMinimumHeight(200)
        scroll_area.setStyleSheet("QScrollArea { border: 1px solid #ddd; }")

        self.package_table = QTableWidget()
        self.package_table.setColumnCount(9)
        self.package_table.setHorizontalHeaderLabels([
            "序号", "包装类型", "数量(kg)", "桶数", "卡板数", "体积(CBM)", "毛重(kg)", "20GP", "操作"
        ])
        self.package_table.verticalHeader().setDefaultSectionSize(28)
        self.package_table.setColumnWidth(0, 60)
        self.package_table.setColumnWidth(1, 90)
        self.package_table.setColumnWidth(2, 60)
        self.package_table.setColumnWidth(3, 40)
        self.package_table.setColumnWidth(4, 55)
        self.package_table.setColumnWidth(5, 65)
        self.package_table.setColumnWidth(6, 65)
        self.package_table.setColumnWidth(7, 55)
        self.package_table.setColumnWidth(8, 50)
        self.package_table.setRowCount(0)
        self.package_table.resizeRowsToContents()

        scroll_area.setWidget(self.package_table)
        package_layout.addWidget(scroll_area)

        # 固定合计行（不随滚动消失）
        self.totals_widget = QWidget()
        totals_layout = QHBoxLayout()
        totals_layout.setContentsMargins(5, 8, 5, 8)
        totals_layout.setSpacing(15)

        # 样式：浅黄色背景，加粗字体
        totals_style = "border-top: 2px0; padding: 5px;"
        self.totals_widget.setStyleSheet(totals_style)

        # 合计标签
        totals_title = QLabel("合计:")
        title_font = QFont("Microsoft YaHei", 10)
        title_font.setBold(True)
        totals_title.setFont(title_font)
        totals_layout.addWidget(totals_title)

        # 桶数合计
        totals_layout.addWidget(QLabel("总桶数:"))
        self.totals_drums = QLabel("0")
        self.totals_drums.setStyleSheet("font-weight: bold;")
        totals_layout.addWidget(self.totals_drums)

        # 卡板数合计
        totals_layout.addWidget(QLabel("总卡板数:"))
        self.totals_pallets = QLabel("0")
        self.totals_pallets.setStyleSheet("font-weight: bold;")
        totals_layout.addWidget(self.totals_pallets)

        # 体积合计
        totals_layout.addWidget(QLabel("总体积:"))
        self.totals_cbm = QLabel("0.00 CBM")
        self.totals_cbm.setStyleSheet("font-weight: bold;")
        totals_layout.addWidget(self.totals_cbm)

        # 毛重合计
        totals_layout.addWidget(QLabel("总毛重:"))
        self.totals_weight = QLabel("0.0 kg")
        self.totals_weight.setStyleSheet("font-weight: bold;")
        totals_layout.addWidget(self.totals_weight)

        # 20GP判定
        totals_layout.addWidget(QLabel("20GP:"))
        self.totals_20gp = QLabel("0个")
        self.totals_20gp.setStyleSheet("font-weight: bold;")
        totals_layout.addWidget(self.totals_20gp)

        totals_layout.addStretch()
        self.totals_widget.setLayout(totals_layout)
        package_layout.addWidget(self.totals_widget)

        # 按钮行
        btn_row2 = QHBoxLayout()
        self.btn_calculate_item = QPushButton("计算")
        self.btn_calculate_item.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-weight: bold;
                padding: 8px 16px;
                border-radius: 4px;
            }
            QPushButton:hover { background-color: #45a049; }
        """)
        self.btn_calculate_item.clicked.connect(self._calculate_package_item)
        btn_row2.addWidget(self.btn_calculate_item)

        self.btn_clear_items = QPushButton("清除")
        self.btn_clear_items.setStyleSheet("""
            QPushButton {
                background-color: #f44336;
                color: white;
                font-weight: bold;
                padding: 8px 16px;
                border-radius: 4px;
            }
            QPushButton:hover { background-color: #d32f2f; }
        """)
        self.btn_clear_items.clicked.connect(self._clear_package_items)
        btn_row2.addWidget(self.btn_clear_items)
        btn_row2.addStretch()
        package_layout.addLayout(btn_row2)

        package_group.setLayout(package_layout)
        layout.addWidget(package_group)

        return panel

    def _load_knowledge(self):
        """加载知识库"""
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        products_file = os.path.join(base_dir, 'knowledge', 'products_knowledge.json')
        packaging_file = os.path.join(base_dir, 'knowledge', 'packaging_data.json')

        self.code_matcher = CodeMatcher(products_file)
        if self.code_matcher.load(products_file):
            version = self.code_matcher.get_version()
            if hasattr(self, 'knowledge_label'):
                self.knowledge_label.setText(f"知识库状态：已加载 (v{version})")
        else:
            if hasattr(self, 'knowledge_label'):
                self.knowledge_label.setText("知识库状态：加载失败")

        self.package_calculator = PackageCalculator(packaging_file)

    def _auto_fill_per_pallet(self):
        """自动填充每板数量（根据选择的包装类型和卡板类型）"""
        pkg_type = self.pkg_type_combo.currentText()
        pallet_type = self.pallet_combo.currentText()

        # 获取卡板尺寸key
        pallet_key = None
        if '1.0' in pallet_type:
            pallet_key = '1.0*1.0'
        elif '1.1' in pallet_type:
            pallet_key = '1.1*1.1'

        drums_per_pallet = None

        if pallet_key:
            # 从pallet_capacity获取标准每板数量
            capacity = self.package_calculator.pallet_capacity.get(pkg_type, {}).get(pallet_key)
            if capacity:
                # capacity可能是整数或元组
                if isinstance(capacity, tuple):
                    drums_per_pallet = capacity[0]
                else:
                    drums_per_pallet = capacity

            # 如果当前选中的pallet没有数据，尝试另一个pallet
            if not drums_per_pallet:
                other_key = '1.1*1.1' if pallet_key == '1.0*1.0' else '1.0*1.0'
                capacity = self.package_calculator.pallet_capacity.get(pkg_type, {}).get(other_key)
                if capacity:
                    if isinstance(capacity, tuple):
                        drums_per_pallet = capacity[0]
                    else:
                        drums_per_pallet = capacity

        if drums_per_pallet:
            self.per_pallet_edit.setText(str(drums_per_pallet))
        else:
            self.per_pallet_edit.setText('')

    def _on_no_pallet_toggled(self, checked: bool):
        """不打卡板选项切换时，禁用/启用卡板版本选择"""
        self.pallet_version_new.setEnabled(not checked)
        self.pallet_version_old.setEnabled(not checked)
        if checked:
            self.pallet_version_new.setChecked(False)
            self.pallet_version_old.setChecked(False)
        else:
            self.pallet_version_new.setChecked(True)

    def _on_order_text_changed(self):
        """订单文本变化时，将tab替换为换行"""
        if self._suppress_text_changed:
            return
        text = self.order_text_edit.toPlainText()
        if '\t' in text:
            self._suppress_text_changed = True
            self.order_text_edit.setPlainText(text.replace('\t', '\n'))
            self._suppress_text_changed = False

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

        # 更新订单要求编辑框
        self.order_req_edit.setPlainText(order_req)

        # 智能推荐包装类型
        if order_req and '粉' in order_req:
            for i in range(self.pkg_type_combo.count()):
                name = self.pkg_type_combo.itemText(i)
                if '纸桶' in name or '编织' in name or '25kg' in name:
                    self.pkg_type_combo.setCurrentIndex(i)
                    break

        self._recalculate_package()
        self._update_result_ui()

        # 自动导入数据到 Phase 2
        self._import_data_to_phase2()

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
        result['收货人地址'] = pi_data.get('收货人地址', '')
        result['日期'] = pi_data.get('日期', '')
        result['PI号'] = pi_data.get('PI号', '')
        result['品名英文'] = pi_data.get('品名英文', '')
        result['数量'] = pi_data.get('数量', '')
        result['单价'] = pi_data.get('单价', '')
        result['金额'] = pi_data.get('金额', '')
        result['H.S.Code'] = pi_data.get('H.S.Code', '')
        result['卸货港'] = pi_data.get('卸货港', '')
        result['包装说明'] = pi_data.get('包装说明', '')

        if not result.get('报关名称'):
            result['报关名称'] = code_info.get('报关名称', '')
        result['海关编码'] = code_info.get('海关编码', '')
        result['报关成分'] = code_info.get('报关成分', '')

        return result

    def _update_result_ui(self):
        """更新结果UI"""
        # 外贸销售订单表字段（23个）
        order_fields = [
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

        self.order_fields_table.setRowCount(len(order_fields))
        for i, (name, value) in enumerate(order_fields):
            self.order_fields_table.setItem(i, 0, QTableWidgetItem(name))
            item = QTableWidgetItem(str(value))
            self.order_fields_table.setItem(i, 1, item)
        self.order_fields_table.resizeRowsToContents()

        # PI文件字段
        pi_fields = [
            ('收货人', self.merged_data.get('收货人', '')),
            ('收货人地址', self.merged_data.get('收货人地址', '')),
            ('日期', self.merged_data.get('日期', '')),
            ('PI号', self.merged_data.get('PI号', '')),
            ('品名英文', self.merged_data.get('品名英文', '')),
            ('数量', self.merged_data.get('数量', '')),
            ('单价', self.merged_data.get('单价', '')),
            ('金额', self.merged_data.get('金额', '')),
            ('H.S.Code', self.merged_data.get('H.S.Code', '')),
            ('卸货港', self.merged_data.get('卸货港', '')),
            ('包装说明', self.merged_data.get('包装说明', '')),
            ('海关编码', self.merged_data.get('海关编码', '')),
            ('报关成分', self.merged_data.get('报关成分', '')),
        ]

        self.pi_fields_table.setRowCount(len(pi_fields))
        for i, (name, value) in enumerate(pi_fields):
            self.pi_fields_table.setItem(i, 0, QTableWidgetItem(name))
            item = QTableWidgetItem(str(value))
            self.pi_fields_table.setItem(i, 1, item)
        self.pi_fields_table.resizeRowsToContents()

    def _shift_values_down(self):
        """下移值：将选中行及之后所有行下移一位"""
        table = self.order_fields_table
        current_row = table.currentRow()
        row_count = table.rowCount()
        if current_row < 0:
            current_row = 0
        if current_row >= row_count - 1:
            self.statusBar().showMessage("已到最后一行，无法继续下移", 3000)
            return

        # 检查下方是否有数据
        has_data = any(
            table.item(i, 1).text()
            for i in range(current_row + 1, row_count)
        )
        if has_data:
            reply = QMessageBox.question(
                self, '确认移动',
                '移动目标位置有数据，是否继续？',
                QMessageBox.Yes | QMessageBox.No
            )
            if reply == QMessageBox.No:
                return

        # 下移：current_row到row_count-1所有行下移一位
        for i in range(row_count - 1, current_row, -1):
            item = table.item(i - 1, 1)
            value = item.text() if item else ''
            table.setItem(i, 1, QTableWidgetItem(value))
        table.setItem(current_row, 1, QTableWidgetItem(''))
        table.setCurrentCell(current_row + 1, 1)
        self.statusBar().showMessage(f"已将第{current_row+1}行及之后下移", 3000)

    def _shift_values_up(self):
        """上移值：将选中行及之后所有行上移一位"""
        table = self.order_fields_table
        current_row = table.currentRow()
        row_count = table.rowCount()
        if current_row < 0:
            current_row = 0
        if current_row <= 0:
            self.statusBar().showMessage("已到第一行，无法继续上移", 3000)
            return

        # 检查下方是否有数据（移动后会覆盖）
        has_data = any(
            table.item(i, 1).text()
            for i in range(current_row + 1, row_count)
        )
        if has_data:
            reply = QMessageBox.question(
                self, '确认移动',
                '移动目标位置有数据，是否继续？',
                QMessageBox.Yes | QMessageBox.No
            )
            if reply == QMessageBox.No:
                return

        # 上移：current_row到row_count-1所有行上移一位
        for i in range(current_row, row_count - 1):
            item = table.item(i + 1, 1)
            value = item.text() if item else ''
            table.setItem(i, 1, QTableWidgetItem(value))
        table.setItem(row_count - 1, 1, QTableWidgetItem(''))
        table.setCurrentCell(current_row - 1 if current_row > 0 else 0, 1)
        self.statusBar().showMessage(f"已将第{current_row+1}行及之后上移", 3000)

    def copy_cell_value(self, row, col):
        """点击单元格复制值"""
        if col == 1:
            table = self.sender()
            item = table.item(row, col)
            if item:
                clipboard = QApplication.clipboard()
                clipboard.setText(item.text())
                self.statusBar().showMessage(f"已复制: {item.text()}", 2000)

    def _on_package_changed(self):
        """当包装选项变化时重新计算当前行"""
        self._calculate_package_item()

    def _calculate_package_item(self):
        """计算包装：使用当前选择的类型和数量，添加新行到表格"""
        pkg_type = self.pkg_type_combo.currentText()
        qty_text = self.pkg_qty_edit.text().strip()
        pallet_type = self.pallet_combo.currentText()

        try:
            qty = float(qty_text) if qty_text else 0
        except:
            qty = 0

        if qty <= 0:
            QMessageBox.warning(self, "警告", "请输入有效的数量")
            return

        if self.no_pallet_check.isChecked():
            result = self.package_calculator.calculate_no_pallet(pkg_type, qty)
        else:
            # 检查是否有手动指定的每板数量
            per_pallet_text = self.per_pallet_edit.text().strip()
            drums_per_pallet = None
            if per_pallet_text:
                try:
                    drums_per_pallet = int(per_pallet_text)
                except:
                    pass
            # 获取卡板版本
            pallet_version = 'new' if self.pallet_version_new.isChecked() else 'old'
            result = self.package_calculator.calculate_with_pallet(pkg_type, pallet_type, qty, drums_per_pallet, pallet_version)

        if 'error' in result:
            QMessageBox.warning(self, "错误", result['error'])
            return

        res = result.get('result', {})
        container = result.get('container_fit', {})

        # 添加新行
        row = self.package_table.rowCount()
        self.package_table.insertRow(row)

        # 填充数据
        self.package_table.setItem(row, 0, QTableWidgetItem(str(row + 1)))  # 序号
        self.package_table.setItem(row, 1, QTableWidgetItem(pkg_type))  # 包装类型
        self.package_table.setItem(row, 2, QTableWidgetItem(str(int(qty))))  # 数量
        self.package_table.setItem(row, 3, QTableWidgetItem(str(int(res.get('total_drums', 0)))))  # 桶数

        if self.no_pallet_check.isChecked():
            self.package_table.setItem(row, 4, QTableWidgetItem('无'))
        else:
            pallets_item = QTableWidgetItem(str(int(res.get('total_pallets', 0))))
            pallets_item.setFlags(pallets_item.flags() | Qt.ItemIsEditable)
            self.package_table.setItem(row, 4, pallets_item)  # 卡板数（可编辑）
        self.package_table.setItem(row, 5, QTableWidgetItem(f"{res.get('total_volume_cbm', 0):.2f}"))  # 体积
        self.package_table.setItem(row, 6, QTableWidgetItem(f"{res.get('gross_weight_kg', 0):.1f}"))  # 毛重

        fits_20gp = container.get('fits_20gp', False)
        full_loads = container.get('full_20gp_loads', 0)
        if fits_20gp:
            self.package_table.setItem(row, 7, QTableWidgetItem(f"{full_loads}个"))
        else:
            self.package_table.setItem(row, 7, QTableWidgetItem('不能'))

        # 删除按钮
        del_btn = QPushButton("删除")
        del_btn.clicked.connect(lambda: self._delete_package_row(row))
        self.package_table.setCellWidget(row, 8, del_btn)

        self._update_package_totals()

    def _delete_package_row(self, row):
        """删除指定行"""
        if 0 <= row < self.package_table.rowCount():
            self.package_table.removeRow(row)
            self._update_package_totals()

    def _recalculate_package(self):
        """重新计算包装（初始产品）"""
        if not hasattr(self, '_current_order_qty') or self._current_order_qty <= 0:
            return

        pkg_type = self.pkg_type_combo.currentText()
        pallet_type = self.pallet_combo.currentText()
        total_qty = self._current_order_qty

        if self.no_pallet_check.isChecked():
            result = self.package_calculator.calculate_no_pallet(pkg_type, total_qty)
        else:
            pallet_version = 'new' if self.pallet_version_new.isChecked() else 'old'
            result = self.package_calculator.calculate_with_pallet(pkg_type, pallet_type, total_qty, None, pallet_version)

        self.package_result = result

        if 'error' in result:
            return

        res = result.get('result', {})
        container = result.get('container_fit', {})

        # 添加一行到表格
        row = self.package_table.rowCount()
        self.package_table.insertRow(row)

        self.package_table.setItem(row, 0, QTableWidgetItem(str(row + 1)))
        self.package_table.setItem(row, 1, QTableWidgetItem(pkg_type))
        self.package_table.setItem(row, 2, QTableWidgetItem(str(int(total_qty))))
        self.package_table.setItem(row, 3, QTableWidgetItem(str(int(res.get('total_drums', 0)))))

        if self.no_pallet_check.isChecked():
            self.package_table.setItem(row, 4, QTableWidgetItem('无'))
        else:
            pallets_item = QTableWidgetItem(str(int(res.get('total_pallets', 0))))
            pallets_item.setFlags(pallets_item.flags() | Qt.ItemIsEditable)
            self.package_table.setItem(row, 4, pallets_item)
        self.package_table.setItem(row, 5, QTableWidgetItem(f"{res.get('total_volume_cbm', 0):.2f}"))
        self.package_table.setItem(row, 6, QTableWidgetItem(f"{res.get('gross_weight_kg', 0):.1f}"))

        fits_20gp = container.get('fits_20gp', False)
        full_loads = container.get('full_20gp_loads', 0)
        if fits_20gp:
            self.package_table.setItem(row, 7, QTableWidgetItem(f"{full_loads}个"))
        else:
            self.package_table.setItem(row, 7, QTableWidgetItem('不能'))

        self._update_package_totals()

    def _update_package_totals(self):
        """更新底部固定合计栏"""
        row_count = self.package_table.rowCount()
        total_drums = 0
        total_pallets = 0
        total_cbm = 0.0
        total_gross = 0.0

        for i in range(row_count):
            try:
                drums_item = self.package_table.item(i, 3)
                if drums_item and drums_item.text() and drums_item.text() != '':
                    total_drums += int(drums_item.text())
            except:
                pass

            try:
                pallets_item = self.package_table.item(i, 4)
                if pallets_item and pallets_item.text() and pallets_item.text() not in ['无', '']:
                    total_pallets += int(pallets_item.text())
            except:
                pass

            try:
                cbm_item = self.package_table.item(i, 5)
                if cbm_item and cbm_item.text():
                    total_cbm += float(cbm_item.text())
            except:
                pass

            try:
                gross_item = self.package_table.item(i, 6)
                if gross_item and gross_item.text():
                    total_gross += float(gross_item.text())
            except:
                pass

        # 计算需要的20GP数量
        max_cbm = 33.07
        total_20gp = 0
        if total_cbm > 0:
            total_20gp = int(total_cbm // max_cbm) + (1 if total_cbm % max_cbm > 0 else 0)

        # 更新固定合计栏
        self.totals_drums.setText(str(total_drums))
        self.totals_pallets.setText(str(total_pallets))
        self.totals_cbm.setText(f"{total_cbm:.2f} CBM")
        self.totals_weight.setText(f"{total_gross:.1f} kg")
        self.totals_20gp.setText(f"{total_20gp}个")

        # 超标警告样式
        if total_pallets > 20:
            self.totals_pallets.setStyleSheet("font-weight: bold; color: red;")
        else:
            self.totals_pallets.setStyleSheet("font-weight: bold; color: #333;")

        if total_cbm > 28:
            self.totals_cbm.setStyleSheet("font-weight: bold; color: red;")
        else:
            self.totals_cbm.setStyleSheet("font-weight: bold; color: #333;")

    def _clear_package_items(self):
        """清除包装计算表格"""
        self.package_table.setRowCount(0)
        self._update_package_totals()

    def _import_data_to_phase2(self):
        """将Phase 1数据导入到Phase 2"""
        if self.booking_widget and self.merged_data:
            self.booking_widget.set_phase1_data(self.merged_data)

            # 获取包装数据
            package_data = self._get_package_data()
            if package_data:
                self.booking_widget.set_package_data(package_data)

    def _get_package_data(self) -> dict:
        """获取包装计算结果"""
        package_data = {}
        row_count = self.package_table.rowCount()
        if row_count > 0:
            pkg_type = self.package_table.item(0, 1).text() if self.package_table.item(0, 1) else ""
            qty = self.package_table.item(0, 2).text() if self.package_table.item(0, 2) else "0"
            drums = self.package_table.item(0, 3).text() if self.package_table.item(0, 3) else "0"
            pallets = self.package_table.item(0, 4).text() if self.package_table.item(0, 4) else "0"
            volume = self.package_table.item(0, 5).text() if self.package_table.item(0, 5) else "0"
            weight = self.package_table.item(0, 6).text() if self.package_table.item(0, 6) else "0"

            package_data = {
                "package_type": pkg_type,
                "quantity_kg": qty,
                "drums": drums,
                "pallets": pallets.replace("无", "0"),
                "volume_cbm": volume.replace(" CBM", ""),
                "gross_weight_kg": weight.replace(" kg", ""),
            }
            pallet_type = self.pallet_combo.currentText()
            package_data["pallet_type"] = pallet_type
        return package_data

    def enter_phase2(self):
        """进入Phase 2（切换Tab）"""
        self.main_tabs.setCurrentIndex(1)