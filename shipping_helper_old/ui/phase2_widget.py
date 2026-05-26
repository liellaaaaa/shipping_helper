# -*- coding: utf-8 -*-
"""
Phase 2 UI - 包装计算器界面
支持两种模式：不打卡板 / 打卡板
"""

import os
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QTextEdit,
    QPushButton, QGroupBox, QGridLayout, QLineEdit,
    QFrame, QComboBox, QMessageBox, QApplication,
    QButtonGroup, QRadioButton, QTableWidget, QTableWidgetItem,
    QHeaderView, QAbstractItemView, QScrollArea
)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QFont

from core.package_calculator import PackageCalculator


class Phase2Widget(QWidget):
    """Phase 2 界面 - 包装计算器（向导式）"""

    def __init__(self, packaging_file: str = None, parent=None):
        super().__init__(parent)
        self.packaging_file = packaging_file
        self.calculator = PackageCalculator(packaging_file)
        self.current_mode = 'pallet'  # 默认打卡板模式
        self.init_ui()

    def init_ui(self):
        """初始化UI"""
        layout = QVBoxLayout()
        layout.setSpacing(10)

        # 标题
        title = QLabel("Phase 2：包装计算器")
        title.setFont(QFont("Microsoft YaHei", 14, QFont.Bold))
        layout.addWidget(title)

        # 订单要求显示
        self.order_group = self._create_order_group()
        layout.addWidget(self.order_group)

        # 模式选择
        mode_group = self._create_mode_group()
        layout.addWidget(mode_group)

        # 打卡板参考表
        self.ref_group = self._create_reference_group()
        layout.addWidget(self.ref_group)

        # 输入区域
        input_group = self._create_input_group()
        layout.addWidget(input_group)

        # 计算按钮
        self.btn_calculate = QPushButton("计算")
        self.btn_calculate.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                padding: 10px 20px;
                border: none;
                border-radius: 4px;
                font-weight: bold;
                font-size: 14px;
            }
            QPushButton:hover { background-color: #1976D2; }
        """)
        self.btn_calculate.clicked.connect(self.calculate)
        layout.addWidget(self.btn_calculate)

        # 结果区域
        self.result_group = self._create_result_group()
        layout.addWidget(self.result_group)

        # 按钮区
        btn_layout = QHBoxLayout()
        self.btn_copy = QPushButton("复制结果")
        self.btn_copy.setEnabled(False)
        self.btn_copy.clicked.connect(self.copy_result)
        self.btn_copy.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                padding: 8px 20px;
                border: none;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover { background-color: #45a049; }
            QPushButton:disabled { background-color: #ccc; }
        """)
        self.btn_clear = QPushButton("清空")
        self.btn_clear.clicked.connect(self.clear_all)
        btn_layout.addWidget(self.btn_copy)
        btn_layout.addStretch()
        btn_layout.addWidget(self.btn_clear)
        layout.addLayout(btn_layout)

        self.setLayout(layout)

    def _create_order_group(self) -> QGroupBox:
        """创建订单要求区域"""
        group = QGroupBox("订单要求")
        layout = QVBoxLayout()

        self.order_label = QLabel("等待从 Phase 1 流转数据...")
        self.order_label.setStyleSheet("color: #666; padding: 5px; background: #f9f9f9; border-radius: 3px;")
        self.order_label.setWordWrap(True)
        layout.addWidget(self.order_label)
        group.setLayout(layout)
        return group

    def _create_mode_group(self) -> QGroupBox:
        """创建模式选择区域"""
        group = QGroupBox("装柜方式")
        layout = QHBoxLayout()

        self.mode_group = QButtonGroup()
        self.rbtn_pallet = QRadioButton("打卡板")
        self.rbtn_pallet.setChecked(True)
        self.rbtn_pallet.toggled.connect(lambda: self.on_mode_changed('pallet'))
        self.rbtn_no_pallet = QRadioButton("不打卡板")
        self.rbtn_no_pallet.toggled.connect(lambda: self.on_mode_changed('no_pallet'))

        self.mode_group.addButton(self.rbtn_pallet)
        self.mode_group.addButton(self.rbtn_no_pallet)

        layout.addWidget(self.rbtn_pallet)
        layout.addWidget(self.rbtn_no_pallet)
        layout.addStretch()

        group.setLayout(layout)
        return group

    def on_mode_changed(self, mode: str):
        """模式切换"""
        self.current_mode = mode
        self.update_input_visibility()

    def update_input_visibility(self):
        """更新输入区域可见性"""
        if not hasattr(self, 'lbl_pallet'):
            return
        if self.current_mode == 'pallet':
            self.lbl_pallet.show()
            self.pallet_combo.show()
            self.ref_group.show()
        else:
            self.lbl_pallet.hide()
            self.pallet_combo.hide()
            self.ref_group.hide()

    def _create_reference_group(self) -> QGroupBox:
        """创建参考表"""
        group = QGroupBox("打卡板参考数量")
        layout = QVBoxLayout()

        self.ref_table = QTableWidget()
        self.ref_table.setColumnCount(3)
        self.ref_table.setHorizontalHeaderLabels(['桶类型', '卡板类型', '可放桶数'])
        self.ref_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.ref_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.ref_table.setMaximumHeight(150)
        self.ref_table.setEditTriggers(QAbstractItemView.NoEditTriggers)

        # 填充参考数据
        ref_data = [
            ('50kg蓝桶(细口/大口)', '1.1*1.1m卡板', '18桶'),
            ('125kg新款胶桶', '1.0*1.0m卡板', '4桶'),
            ('125kg新款胶桶', '1.1*1.1m卡板', '5桶'),
            ('25kg纸桶', '1.1*1.1m卡板', '18桶'),
            ('50kg纸桶', '1.1*1.1m卡板', '12桶'),
            ('30kg蓝桶', '1.1*1.1m卡板', '24桶'),
        ]
        self.ref_table.setRowCount(len(ref_data))
        for i, (drum, pallet, count) in enumerate(ref_data):
            self.ref_table.setItem(i, 0, QTableWidgetItem(drum))
            self.ref_table.setItem(i, 1, QTableWidgetItem(pallet))
            self.ref_table.setItem(i, 2, QTableWidgetItem(count))

        layout.addWidget(self.ref_table)
        group.setLayout(layout)
        return group

    def _create_input_group(self) -> QGroupBox:
        """创建输入区域"""
        group = QGroupBox("输入")
        layout = QGridLayout()
        layout.setSpacing(8)

        # 桶类型
        lbl_drum = QLabel("桶类型:")
        layout.addWidget(lbl_drum, 0, 0)
        self.drum_combo = QComboBox()
        self.drum_combo.addItems(self.calculator.get_package_options())
        layout.addWidget(self.drum_combo, 0, 1)

        # 卡板类型
        self.lbl_pallet = QLabel("卡板类型:")
        layout.addWidget(self.lbl_pallet, 0, 2)
        self.pallet_combo = QComboBox()
        self.pallet_combo.addItems(self.calculator.get_pallet_options())
        layout.addWidget(self.pallet_combo, 0, 3)

        # 产品数量
        lbl_qty = QLabel("产品数量(kg):")
        layout.addWidget(lbl_qty, 1, 0)
        self.qty_input = QLineEdit()
        self.qty_input.setPlaceholderText("输入产品总重量(kg)")
        layout.addWidget(self.qty_input, 1, 1, 1, 3)

        # 每板桶数（用户手动指定实际每板装的数量）
        lbl_per_pallet = QLabel("每板桶数:")
        layout.addWidget(lbl_per_pallet, 2, 0)
        self.per_pallet_input = QLineEdit()
        self.per_pallet_input.setPlaceholderText("手动填写每板实际桶数")
        self.per_pallet_input.setMaximumWidth(120)
        layout.addWidget(self.per_pallet_input, 2, 1)

        # 添加按钮
        self.btn_add = QPushButton("+ 添加")
        self.btn_add.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                padding: 5px 15px;
                border: none;
                border-radius: 3px;
            }
            QPushButton:hover { background-color: #45a049; }
        """)
        self.btn_add.clicked.connect(self.add_item)
        layout.addWidget(self.btn_add, 2, 2, 1, 2)

        # 物品列表
        self.items_table = QTableWidget()
        self.items_table.setColumnCount(9)
        self.items_table.setHorizontalHeaderLabels([
            '装柜方式', '桶类型', '卡板类型', '产品数量(kg)', '桶数', '卡板数',
            '桶体积(CBM)', '毛重(kg)', '备注'
        ])
        self.items_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.items_table.setMaximumHeight(150)
        self.items_table.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.EditKeyPressed)
        self.items_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        layout.addWidget(self.items_table, 3, 0, 1, 4)

        # 操作按钮
        btn_row = QHBoxLayout()
        self.btn_remove = QPushButton("- 移除选中")
        self.btn_remove.clicked.connect(self.remove_item)
        btn_row.addWidget(self.btn_remove)
        btn_row.addStretch()
        layout.addLayout(btn_row, 4, 0, 1, 4)

        group.setLayout(layout)
        return group

    def _create_result_group(self) -> QGroupBox:
        """创建结果区域"""
        group = QGroupBox("计算结果")
        layout = QGridLayout()
        layout.setSpacing(8)

        self.result_fields = {}

        fields = [
            ('总产品数量(kg)', 'total_product_kg'),
            ('总桶数', 'total_drums'),
            ('总卡板数', 'total_pallets'),
            ('桶体积(CBM)', 'drum_cbm'),
            ('卡板占用(CBM)', 'pallet_cbm'),
            ('总体积(CBM)', 'total_cbm'),
            ('总毛重(kg)', 'total_gross_kg'),
            ('20GP柜数', 'container_20gp'),
            ('能否装20GP', 'fits_20gp'),
        ]

        row = 0
        col = 0
        for label_text, key in fields:
            lbl = QLabel(f"{label_text}:")
            lbl.setFont(QFont("Microsoft YaHei", 9))
            layout.addWidget(lbl, row, col * 2)

            edit = QLineEdit()
            edit.setReadOnly(True)
            edit.setStyleSheet("""
                QLineEdit {
                    background-color: #f5f5f5;
                    border: 1px solid #ddd;
                    padding: 4px 8px;
                    border-radius: 3px;
                }
            """)
            layout.addWidget(edit, row, col * 2 + 1)
            self.result_fields[key] = edit

            col += 1
            if col >= 2:
                col = 0
                row += 1

        group.setLayout(layout)
        return group

    def set_order_text(self, order_text: str, order_qty: float = None):
        """从外部设置订单要求和订单量"""
        display = order_text
        if order_qty:
            display = f"订单要求：{order_text}　|　订单量：{order_qty}kg（可编辑）"
        else:
            display = f"订单要求：{order_text}"
        self.order_label.setText(display)
        self.order_label.setStyleSheet("color: #333; padding: 5px; background: #f9f9f9; border-radius: 3px;")
        # 预填订单量到输入框
        if order_qty:
            try:
                qty_value = float(order_qty)
                self.qty_input.setText(str(int(qty_value)))
            except (ValueError, TypeError):
                self.qty_input.setText(str(order_qty))

    def add_item(self):
        """添加一项"""
        drum_name = self.drum_combo.currentText()
        qty_text = self.qty_input.text().strip()
        per_pallet_text = self.per_pallet_input.text().strip()

        if not qty_text:
            QMessageBox.warning(self, "提示", "请输入产品数量")
            return

        try:
            qty_kg = float(qty_text)
        except ValueError:
            QMessageBox.warning(self, "错误", "产品数量必须是数字")
            return

        if qty_kg <= 0:
            QMessageBox.warning(self, "错误", "产品数量必须大于0")
            return

        mode = self.current_mode
        pallet_name = self.pallet_combo.currentText() if mode == 'pallet' else '不打卡板'

        # 计算
        if mode == 'pallet':
            result = self.calculator.calculate_with_pallet(drum_name, pallet_name, qty_kg)
        else:
            result = self.calculator.calculate_no_pallet(drum_name, qty_kg)

        if 'error' in result:
            QMessageBox.warning(self, "错误", result['error'])
            return

        r = result['result']
        c = result.get('container_fit', {})

        # 如果用户手动指定了每板桶数，使用用户指定的值
        custom_per_pallet = None
        if per_pallet_text and mode == 'pallet':
            try:
                custom_per_pallet = int(per_pallet_text)
            except ValueError:
                pass

        # 添加到表格
        row = self.items_table.rowCount()
        self.items_table.insertRow(row)

        self.items_table.setItem(row, 0, QTableWidgetItem('打卡板' if mode == 'pallet' else '不打卡板'))
        self.items_table.setItem(row, 1, QTableWidgetItem(drum_name))
        self.items_table.setItem(row, 2, QTableWidgetItem(pallet_name))
        self.items_table.setItem(row, 3, QTableWidgetItem(str(int(qty_kg))))

        if custom_per_pallet:
            # 使用用户指定的每板桶数重新计算
            total_drums = r['total_drums']
            pallets_needed = (total_drums + custom_per_pallet - 1) // custom_per_pallet  # 向上取整
            drums_per_last_pallet = total_drums % custom_per_pallet
            if drums_per_last_pallet == 0:
                drums_per_last_pallet = custom_per_pallet

            # 重新计算体积和毛重（按比例）
            drum_volume_each = r['drum_volume_cbm'] / r['total_drums'] if r['total_drums'] > 0 else 0
            gross_weight_each = r['gross_weight_kg'] / r['total_drums'] if r['total_drums'] > 0 else 0

            self.items_table.setItem(row, 4, QTableWidgetItem(str(total_drums)))
            self.items_table.setItem(row, 5, QTableWidgetItem(str(pallets_needed)))
            self.items_table.setItem(row, 6, QTableWidgetItem(str(round(r['drum_volume_cbm'], 3))))
            self.items_table.setItem(row, 7, QTableWidgetItem(str(round(r['gross_weight_kg'], 1))))
            note = f"每板{custom_per_pallet}桶(手动)"
            if drums_per_last_pallet != custom_per_pallet:
                note += f"，最后板{drums_per_last_pallet}桶"
            note += f"，卡板体积:{r.get('pallet_volume_cbm', 0):.3f}CBM"
            self.items_table.setItem(row, 8, QTableWidgetItem(note))
        else:
            self.items_table.setItem(row, 4, QTableWidgetItem(str(r['total_drums'])))
            self.items_table.setItem(row, 5, QTableWidgetItem(str(r.get('total_pallets', 0))))
            self.items_table.setItem(row, 6, QTableWidgetItem(str(r['drum_volume_cbm'])))
            self.items_table.setItem(row, 7, QTableWidgetItem(str(r['gross_weight_kg'])))
            if mode == 'pallet':
                note = f"卡板体积:{r.get('pallet_volume_cbm', 0):.3f}CBM"
            else:
                note = c.get('fits_20gp', False) and '可装20GP' or f"需{int(c.get('full_20gp_loads', 1))}个柜"
            self.items_table.setItem(row, 8, QTableWidgetItem(note))

        # 清空输入框
        self.qty_input.clear()
        self.per_pallet_input.clear()
        self.qty_input.setFocus()

    def remove_item(self):
        """移除选中的项"""
        current_row = self.items_table.currentRow()
        if current_row >= 0:
            self.items_table.removeRow(current_row)

    def calculate(self):
        """汇总计算"""
        if self.items_table.rowCount() == 0:
            QMessageBox.warning(self, "提示", "请先添加物品")
            return

        total_product_kg = 0.0
        total_drums = 0
        total_pallets = 0
        drum_cbm = 0.0
        pallet_cbm = 0.0
        total_gross_kg = 0.0

        for row in range(self.items_table.rowCount()):
            try:
                total_product_kg += float(self.items_table.item(row, 3).text())
                total_drums += int(self.items_table.item(row, 4).text())
                total_pallets += int(self.items_table.item(row, 5).text())
                drum_cbm += float(self.items_table.item(row, 6).text())
                total_gross_kg += float(self.items_table.item(row, 7).text())

                # 从备注中提取卡板体积
                note = self.items_table.item(row, 8).text()
                if '卡板体积' in note:
                    import re
                    match = re.search(r'卡板体积:([\d.]+)CBM', note)
                    if match:
                        pallet_cbm += float(match.group(1))
            except (ValueError, AttributeError):
                continue

        total_cbm = drum_cbm + pallet_cbm

        # 使用动态集装箱参数（20GP）
        container_20gp_specs = self.calculator.container_specs['20gp']
        max_cbm = container_20gp_specs['max_cbm']
        max_pallets = container_20gp_specs['max_pallets']

        fits_20gp = total_cbm <= max_cbm and total_pallets <= max_pallets
        container_20gp = max(
            int(total_cbm // max_cbm) + (1 if total_cbm % max_cbm > 0 else 0),
            int(total_pallets // max_pallets) + (1 if total_pallets % max_pallets > 0 else 0)
        )

        # 更新结果
        self.result_fields['total_product_kg'].setText(f"{total_product_kg:.1f}")
        self.result_fields['total_drums'].setText(str(total_drums))
        self.result_fields['total_pallets'].setText(str(total_pallets))
        self.result_fields['drum_cbm'].setText(f"{drum_cbm:.3f}")
        self.result_fields['pallet_cbm'].setText(f"{pallet_cbm:.3f}")
        self.result_fields['total_cbm'].setText(f"{total_cbm:.3f}")
        self.result_fields['total_gross_kg'].setText(f"{total_gross_kg:.1f}")
        self.result_fields['container_20gp'].setText(str(container_20gp))
        self.result_fields['fits_20gp'].setText('可以' if fits_20gp else f'超出，需{container_20gp}个柜')

        self.btn_copy.setEnabled(True)

    def copy_result(self):
        """复制结果"""
        lines = []
        for key, edit in self.result_fields.items():
            label = [k for k, v in [
                ('总产品数量(kg)', 'total_product_kg'),
                ('总桶数', 'total_drums'),
                ('总卡板数', 'total_pallets'),
                ('桶体积(CBM)', 'drum_cbm'),
                ('卡板占用(CBM)', 'pallet_cbm'),
                ('总体积(CBM)', 'total_cbm'),
                ('总毛重(kg)', 'total_gross_kg'),
                ('20GP柜数', 'container_20gp'),
                ('能否装20GP', 'fits_20gp'),
            ] if v == key][0]
            lines.append(f"{label}: {edit.text()}")

        text = '\n'.join(lines)
        QApplication.clipboard().setText(text)
        QMessageBox.information(self, "复制成功", "结果已复制到剪贴板")

    def clear_all(self):
        """清空所有"""
        self.order_label.setText("等待从 Phase 1 流转数据...")
        self.order_label.setStyleSheet("color: #666; padding: 5px;")
        self.items_table.setRowCount(0)
        for edit in self.result_fields.values():
            edit.clear()
        self.qty_input.clear()
        self.per_pallet_input.clear()
        self.btn_copy.setEnabled(False)