# -*- coding: utf-8 -*-
"""
Phase 1 UI - PI数据汇聚界面
"""

import os
import json
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QTextEdit,
    QPushButton, QFileDialog, QMessageBox, QGroupBox,
    QGridLayout, QLineEdit, QFrame, QScrollArea, QApplication,
    QSizePolicy
)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QFont

from core.order_parser import OrderParser
from core.pi_extractor import PIExtractor
from core.code_matcher import CodeMatcher
from core.merger import Merger
from core.project_manager import ProjectManager
from core.config_manager import get_config


class ClickableLineEdit(QLineEdit):
    """可编辑且点击可复制的文本框"""
    clicked = pyqtSignal()

    def __init__(self, text='', parent=None):
        super().__init__(text, parent)
        self.setStyleSheet("""
            QLineEdit {
                background-color: #fff;
                border: 1px solid #ddd;
                padding: 4px 8px;
                border-radius: 3px;
            }
            QLineEdit:hover {
                background-color: #e8f4fc;
                border: 1px solid #2196F3;
            }
        """)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.selectAll()
            QApplication.clipboard().setText(self.text())
            self.clicked.emit()
        super().mousePressEvent(event)


class Phase1Widget(QWidget):
    """Phase 1 界面 - PI数据汇聚"""

    # 保存成功后发送信号：(订单要求, 订单量kg)
    order_saved = pyqtSignal(tuple)

    # 汇聚结果字段列表
    RESULT_FIELDS = [
        ('订单号', '订单号'),
        ('客户编码', '客户编码'),
        ('业务员', '业务员'),
        ('内部编号', '内部编号'),
        ('产品中文名', '产品中文名'),
        ('报关名称', '报关名称'),
        ('规格kg', '规格kg'),
        ('订单量kg', '订单量kg'),
        ('数量', '数量'),
        ('单价', '单价'),
        ('金额', '金额'),
        ('产品颜色', '产品颜色'),
        ('海关编码', '海关编码'),
        ('报关成分', '报关成分'),
        ('收货人', '收货人'),
        ('地址', '地址'),
        ('品名英文', '品名英文'),
        ('卸货港', '卸货港'),
        ('H.S.Code', 'H.S.Code'),
        ('包装说明', '包装说明'),
        ('日期', '日期'),
        ('发货人', '发货人'),
    ]

    def __init__(self, codes_file: str, parent=None):
        super().__init__(parent)
        self.codes_file = codes_file
        self.code_matcher = CodeMatcher()
        self.project_manager = ProjectManager()
        self.merged_data = {}
        self.pi_filepath = None
        self.result_fields = {}  # 初始化result_fields字典

        self.init_ui()

        # 加载商品编码表
        self.load_codes()

        # 加载默认发货人
        self._load_default_shipper()

    def init_ui(self):
        """初始化UI"""
        layout = QVBoxLayout()
        layout.setSpacing(10)

        # 标题
        title = QLabel("Phase 1：PI数据汇聚")
        title.setFont(QFont("Microsoft YaHei", 14, QFont.Bold))
        layout.addWidget(title)

        # 编码表状态
        self.codes_status = QLabel("商品编码表：未加载")
        self.codes_status.setStyleSheet("color: #666;")
        layout.addWidget(self.codes_status)

        # 步骤1：粘贴外贸销售订单表
        group1 = self._create_paste_group()
        layout.addWidget(group1)

        # 步骤2：上传Proforma Invoice
        group2 = self._create_upload_group()
        layout.addWidget(group2)

        # 步骤3：解析预览
        group3 = self._create_preview_group()
        layout.addWidget(group3)

        # 执行汇聚按钮
        self.btn_merge = QPushButton("执行汇聚")
        self.btn_merge.setStyleSheet("""
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
        self.btn_merge.clicked.connect(self.do_merge_and_display)
        layout.addWidget(self.btn_merge)

        # 汇聚结果（步骤4和5合并）
        self.result_group = self._create_result_group()
        layout.addWidget(self.result_group)

        # 按钮区
        btn_layout = QHBoxLayout()
        self.btn_copy_all = QPushButton("复制全部字段")
        self.btn_copy_all.setEnabled(False)
        self.btn_copy_all.clicked.connect(self.copy_all_fields)

        self.btn_save = QPushButton("确认并保存")
        self.btn_save.setEnabled(False)
        self.btn_save.setStyleSheet("""
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
        self.btn_save.clicked.connect(self.save_data)

        self.btn_clear = QPushButton("清空重新输入")
        self.btn_clear.clicked.connect(self.clear_all)

        self.btn_refresh = QPushButton("刷新")
        self.btn_refresh.setStyleSheet("""
            QPushButton {
                background-color: #ff9800;
                color: white;
                padding: 8px 16px;
                border: none;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover { background-color: #e68a00; }
        """)
        self.btn_refresh.clicked.connect(self.refresh)

        btn_layout.addWidget(self.btn_copy_all)
        btn_layout.addStretch()
        btn_layout.addWidget(self.btn_save)
        btn_layout.addWidget(self.btn_clear)
        btn_layout.addWidget(self.btn_refresh)
        layout.addLayout(btn_layout)

        self.setLayout(layout)

    def _create_paste_group(self) -> QGroupBox:
        """创建粘贴区域"""
        group = QGroupBox("步骤2：粘贴外贸销售订单表的一行数据")
        layout = QVBoxLayout()

        self.paste_edit = QTextEdit()
        self.paste_edit.setPlaceholderText("从在线文档复制一行数据，Ctrl+V 粘贴到此处...")
        self.paste_edit.setMaximumHeight(120)
        self.paste_edit.textChanged.connect(self.on_paste_changed)
        layout.addWidget(self.paste_edit)

        self.paste_status = QLabel("等待粘贴数据...")
        self.paste_status.setStyleSheet("color: #666;")
        layout.addWidget(self.paste_status)

        group.setLayout(layout)
        return group

    def _create_upload_group(self) -> QGroupBox:
        """创建文件上传区域"""
        group = QGroupBox("步骤3：上传 Proforma Invoice (.xls)")
        layout = QHBoxLayout()

        self.pi_path_edit = QLineEdit()
        self.pi_path_edit.setReadOnly(True)
        self.pi_path_edit.setPlaceholderText("请选择 Proforma Invoice 文件...")

        btn_browse = QPushButton("选择文件")
        btn_browse.clicked.connect(self.browse_pi_file)

        layout.addWidget(self.pi_path_edit, 1)
        layout.addWidget(btn_browse)

        group.setLayout(layout)
        return group

    def _create_preview_group(self) -> QGroupBox:
        """创建解析预览区域"""
        group = QGroupBox("解析预览")
        layout = QVBoxLayout()

        self.preview_label = QLabel("等待数据输入...")
        self.preview_label.setWordWrap(True)
        self.preview_label.setStyleSheet("color: #666; padding: 5px; background: #f9f9f9;")
        layout.addWidget(self.preview_label)
        group.setLayout(layout)
        return group

    def _create_result_group(self) -> QGroupBox:
        """创建汇聚结果区域 - 两列响应式布局"""
        group = QGroupBox("步骤5：汇聚结果（点击可复制单个字段）")

        outer = QVBoxLayout()
        outer.setSpacing(4)
        outer.setContentsMargins(3, 5, 3, 5)

        # 左列10个字段，右列12个字段
        left = self.RESULT_FIELDS[:10]
        right = self.RESULT_FIELDS[10:]

        # 使用两个并排的GridLayout
        left_grid = QGridLayout()
        left_grid.setSpacing(3)
        left_grid.setContentsMargins(0, 0, 0, 0)

        right_grid = QGridLayout()
        right_grid.setSpacing(3)
        right_grid.setContentsMargins(0, 0, 0, 0)

        for row_i, (lbl_txt, key) in enumerate(left):
            lbl = QLabel(f"{lbl_txt}:")
            lbl.setFont(QFont("Microsoft YaHei", 9))
            lbl.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
            edit = ClickableLineEdit()
            edit.setFont(QFont("Microsoft YaHei", 9))
            edit.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            left_grid.addWidget(lbl, row_i, 0)
            left_grid.addWidget(edit, row_i, 1)
            self.result_fields[key] = edit

        for row_i, (lbl_txt, key) in enumerate(right):
            lbl = QLabel(f"{lbl_txt}:")
            lbl.setFont(QFont("Microsoft YaHei", 9))
            lbl.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
            edit = ClickableLineEdit()
            edit.setFont(QFont("Microsoft YaHei", 9))
            edit.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            right_grid.addWidget(lbl, row_i, 0)
            right_grid.addWidget(edit, row_i, 1)
            self.result_fields[key] = edit

        # 列宽比例：标签固定60，input占剩余空间
        left_grid.setColumnMinimumWidth(0, 60)
        left_grid.setColumnStretch(1, 1)
        right_grid.setColumnMinimumWidth(0, 60)
        right_grid.setColumnStretch(1, 1)

        row = QHBoxLayout()
        row.setSpacing(8)
        row.addLayout(left_grid, 1)
        row.addLayout(right_grid, 1)

        warning = QLabel("⚠️ 请仔细核对以下数据，确认无误后点击\"确认并保存\"")
        warning.setFont(QFont("Microsoft YaHei", 9))
        warning.setStyleSheet("color: #ff9800; font-weight: bold; padding: 4px; background-color: #fff3e0; border-radius: 3px;")

        outer.addLayout(row)
        outer.addWidget(warning)
        group.setLayout(outer)
        return group

    def load_codes(self):
        """加载商品编码表"""
        if self.code_matcher.load(self.codes_file):
            version = self.code_matcher.get_version()
            self.codes_status.setText(f"商品编码表：已加载 (版本: {version or '未知'})")
            self.codes_status.setStyleSheet("color: #4CAF50;")
        else:
            self.codes_status.setText("商品编码表：加载失败！")
            self.codes_status.setStyleSheet("color: #f44336;")

    def _load_default_shipper(self):
        """加载默认发货人"""
        config = get_config()
        default_shipper = config.get_shipper()
        if default_shipper:
            self.result_fields['发货人'].setText(default_shipper)

    def on_paste_changed(self):
        """粘贴内容变化时的处理"""
        text = self.paste_edit.toPlainText().strip()
        if text:
            self.paste_status.setText("正在解析...")
            self.paste_status.setStyleSheet("color: #2196F3;")
            # 延迟解析，避免频繁触发
            self.parse_order_data()
        else:
            self.paste_status.setText("等待粘贴数据...")
            self.paste_status.setStyleSheet("color: #666;")
            self.preview_label.setText("等待数据输入...")

    def parse_order_data(self):
        """解析外贸销售订单表数据"""
        text = self.paste_edit.toPlainText()
        if not text.strip():
            return

        parser = OrderParser()
        data = parser.parse(text)

        # 更新预览
        pi_no = data.get('订单号', '')
        internal_code = data.get('内部编号', '')

        preview = f"<b>订单号(PI号):</b> {pi_no or '<font color=red>未识别</font>'}<br>"
        preview += f"<b>客户编号:</b> {data.get('客户编号', '')}<br>"
        preview += f"<b>业务员:</b> {data.get('业务员', '')}<br>"
        preview += f"<b>内部编号:</b> {internal_code or '<font color=red>未识别</font>'}"

        self.preview_label.setText(preview)
        self.paste_status.setText("解析完成")
        self.paste_status.setStyleSheet("color: #4CAF50;")

    def browse_pi_file(self):
        """浏览并选择PI文件"""
        filepath, _ = QFileDialog.getOpenFileName(
            self,
            "选择 Proforma Invoice 文件",
            "",
            "Excel Files (*.xls *.xlsx);;All Files (*)"
        )
        if filepath:
            self.pi_filepath = filepath
            self.pi_path_edit.setText(filepath)
            self.parse_pi_file()

    def parse_pi_file(self):
        """解析PI文件并更新预览"""
        if not self.pi_filepath or not os.path.exists(self.pi_filepath):
            return

        try:
            extractor = PIExtractor()
            pi_data = extractor.parse_file(self.pi_filepath)

            # 更新预览区域显示PI解析结果
            preview = self.preview_label.text()
            if preview == "等待数据输入...":
                preview = ""

            preview += f"<br><b>PI文件解析结果:</b><br>"
            for key, value in pi_data.items():
                if value:
                    preview += f"{key}: {value}<br>"

            self.preview_label.setText(preview if preview else "等待数据输入...")
        except Exception as e:
            self.preview_label.setText(f"PI文件解析失败: {str(e)}")

    def on_field_clicked(self, edit: ClickableLineEdit, key: str):
        """字段被点击复制"""
        pass  # 已通过ClickableLineEdit自动处理

    def copy_all_fields(self):
        """复制全部字段"""
        lines = []
        for label_text, key in self.RESULT_FIELDS:
            value = self.result_fields[key].text()
            lines.append(f"{label_text}: {value}")

        text = '\t'.join([self.result_fields[key].text() for _, key in self.RESULT_FIELDS])
        QApplication.clipboard().setText(text)

        QMessageBox.information(self, "复制成功", "全部字段已复制到剪贴板（Tab分隔）")

    def merge_data(self):
        """执行数据汇聚"""
        # 获取外贸销售订单表数据
        order_text = self.paste_edit.toPlainText()
        if not order_text.strip():
            QMessageBox.warning(self, "提示", "请先粘贴外贸销售订单表数据")
            return False

        # 获取用户编辑的发货人
        shipper = self.result_fields['发货人'].text() or ''

        # 执行汇聚
        merger = Merger(shipper=shipper)
        self.merged_data = merger.merge(order_text, self.pi_filepath, self.code_matcher)

        return True

    def update_result_display(self):
        """更新结果显示"""
        for label_text, key in self.RESULT_FIELDS:
            value = self.merged_data.get(key, '')
            self.result_fields[key].setText(str(value))

    def save_data(self):
        """保存数据"""
        # 获取PI号
        pi_no = self.result_fields['订单号'].text() or self.result_fields['PI号'].text()
        if not pi_no:
            QMessageBox.warning(self, "错误", "PI号不能为空，请检查数据")
            return

        # 创建项目目录
        project_dir = self.project_manager.create_project(pi_no)

        # 保存PI文件到input
        if self.pi_filepath:
            self.project_manager.copy_input_file('p1', self.pi_filepath)

        # 保存粘贴数据缓存到input
        paste_text = self.paste_edit.toPlainText()
        paste_cache_path = os.path.join(
            self.project_manager.get_phase_dir('p1', 'input'),
            'order_data_cache.txt'
        )
        with open(paste_cache_path, 'w', encoding='utf-8') as f:
            f.write(paste_text)

        # 保存汇聚结果到output
        output_path = self.project_manager.save_json(
            'p1',
            'pi_row_data.json',
            self.merged_data
        )

        # 提示成功
        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Information)
        msg.setWindowTitle("保存成功")
        msg.setText(f"数据已保存至：\n{output_path}\n\n项目目录：\n{project_dir}")
        btn_close = msg.addButton("关闭", QMessageBox.RejectRole)
        btn_next = msg.addButton("进入下一步", QMessageBox.AcceptRole)
        msg.setDefaultButton(btn_next)
        msg.exec_()

        if msg.clickedButton() == btn_next:
            # 发送订单要求+订单量信号
            order_req = self.merged_data.get('订单要求', '')
            order_qty = self.merged_data.get('订单量kg', 0)
            self.order_saved.emit((order_req, order_qty))

        # 启用复制按钮
        self.btn_copy_all.setEnabled(True)
        self.btn_save.setEnabled(True)

    def clear_all(self):
        """清空所有输入"""
        self.paste_edit.clear()
        self.pi_filepath = None
        self.pi_path_edit.clear()
        self.merged_data = {}
        self.preview_label.setText("等待数据输入...")
        self.paste_status.setText("等待粘贴数据...")
        self.paste_status.setStyleSheet("color: #666;")

        for key in self.result_fields:
            self.result_fields[key].clear()

        self.btn_copy_all.setEnabled(False)
        self.btn_save.setEnabled(False)

    def refresh(self):
        """刷新/重置界面"""
        self.clear_all()

    def do_merge_and_display(self):
        """执行汇聚并显示结果（供外部调用）"""
        if self.merge_data():
            self.update_result_display()
            self.btn_save.setEnabled(True)