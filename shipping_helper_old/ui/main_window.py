# -*- coding: utf-8 -*-
"""
主窗口 - ShippingHelper
"""

import os
import sys
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QTabWidget,
    QMessageBox, QLabel, QPushButton, QHBoxLayout,
    QScrollArea, QFrame
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont, QIcon

from .phase1_widget import Phase1Widget
from .phase2_widget import Phase2Widget
from .phase3_widget import Phase3Widget
from core.config_manager import get_config


class MainWindow(QMainWindow):
    """主窗口"""

    def __init__(self):
        super().__init__()

        self.config = get_config()
        self.init_ui()

    def init_ui(self):
        """初始化UI"""
        self.setWindowTitle("ShippingHelper - 船务部效率工具")
        self.setMinimumSize(900, 700)

        # 滚动区域
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        self.setCentralWidget(scroll)

        # 中央控件
        central_widget = QWidget()
        layout = QVBoxLayout()

        # 标题栏
        header = QLabel("ShippingHelper - 船务部效率工具")
        header.setFont(QFont("Microsoft YaHei", 16, QFont.Bold))
        header.setAlignment(Qt.AlignCenter)
        header.setStyleSheet("""
            QLabel {
                padding: 15px;
                background-color: #2196F3;
                color: white;
                border-radius: 5px;
            }
        """)
        layout.addWidget(header)

        # 标签页
        tabs = QTabWidget()
        self.tabs = tabs

        # Phase 1 标签页
        codes_file = self.config.get_path('product_codes_json')
        self.phase1_widget = Phase1Widget(codes_file, self)
        self.phase1_widget.order_saved.connect(self.on_order_saved)
        tabs.addTab(self.phase1_widget, "Phase 1: PI数据汇聚")

        # Phase 2 标签页（包装计算器自动查找包装资料文件）
        self.phase2_widget = Phase2Widget(None, self)
        tabs.addTab(self.phase2_widget, "Phase 2: 包装计算器")

        # Phase 3 标签页
        self.phase3_widget = Phase3Widget(self)
        tabs.addTab(self.phase3_widget, "Phase 3: MSDS/LOI填写")

        layout.addWidget(tabs)

        central_widget.setLayout(layout)
        scroll.setWidget(central_widget)

        # 菜单栏
        self.create_menu_bar()

    def on_order_saved(self, order_data: tuple):
        """接收Phase 1保存的订单要求和订单量，自动跳转到Phase 2并填充"""
        order_req, order_qty = order_data
        # 切换到Phase 2标签页
        self.tabs.setCurrentIndex(1)
        # 填充订单要求和订单量
        self.phase2_widget.set_order_text(order_req, order_qty)

    def create_menu_bar(self):
        """创建菜单栏"""
        menubar = self.menuBar()

        # 文件菜单
        file_menu = menubar.addMenu('文件')

        exit_action = file_menu.addAction('退出')
        exit_action.triggered.connect(self.close)

        # 帮助菜单
        help_menu = menubar.addMenu('帮助')

        about_action = help_menu.addAction('关于')
        about_action.triggered.connect(self.show_about)

    def show_about(self):
        """显示关于对话框"""
        QMessageBox.about(
            self,
            "关于 ShippingHelper",
            "<h3>ShippingHelper</h3>"
            "<p>船务部效率工具 v1.0</p>"
            "<p>Phase 1: PI数据汇聚</p>"
            "<p>自动汇聚外贸销售订单表、Proforma Invoice、商品编码表数据</p>"
        )

    def closeEvent(self, event):
        """关闭窗口时的事件"""
        reply = QMessageBox.question(
            self,
            '确认退出',
            '确定要退出程序吗？',
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()
