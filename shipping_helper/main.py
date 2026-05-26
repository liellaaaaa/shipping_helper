# shipping_helper/main.py
# -*- coding: utf-8 -*-
"""
ShippingHelper - 宏昊船务助手
入口文件
"""

import sys
from PyQt5.QtWidgets import QApplication
from ui.main_window import MainWindow


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("ShippingHelper")

    window = MainWindow()
    window.show()

    sys.exit(app.exec_())


if __name__ == '__main__':
    main()