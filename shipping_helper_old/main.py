# -*- coding: utf-8 -*-
"""
ShippingHelper - 船务部效率工具
入口文件
"""

import sys
import os

# 设置 Qt 平台插件路径（在 import PyQt5 之前）
import PyQt5
plugin_path = os.path.join(os.path.dirname(PyQt5.__file__), 'Qt5', 'plugins')
os.environ['QT_QPA_PLATFORM_PLUGIN_PATH'] = plugin_path

# 添加项目根目录到路径
if getattr(sys, 'frozen', False):
    # 打包后的程序
    base_dir = os.path.dirname(sys.executable)
    sys.path.insert(0, base_dir)
else:
    # 开发环境
    base_dir = os.path.dirname(os.path.abspath(__file__))
    sys.path.insert(0, base_dir)

from PyQt5.QtWidgets import QApplication
from ui.main_window import MainWindow


def main():
    """主函数"""
    app = QApplication(sys.argv)

    # 设置应用程序样式
    app.setStyle('Fusion')

    # 创建并显示主窗口
    window = MainWindow()
    window.show()

    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
