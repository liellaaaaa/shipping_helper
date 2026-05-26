# -*- coding: utf-8 -*-
"""
Phase 3 UI - MSDS/LOI 模板填写界面
左侧显示从运输鉴定报告提取的字段，右侧显示Word编辑器
"""

import os
import shutil
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QGroupBox, QGridLayout, QLineEdit, QScrollArea,
    QFileDialog, QMessageBox, QFrame, QComboBox, QTableWidget,
    QTableWidgetItem, QHeaderView, QAbstractItemView, QApplication
)
from PyQt5.QtCore import Qt, pyqtSignal, QSize
from PyQt5.QtGui import QFont

from core.report_extractor import ReportExtractor
from core.product_codes import ProductCodesLoader
from core.config_manager import get_config

# 检查QAxContainer是否可用
try:
    from PyQt5.QtAxContainer import QAxWidget
    HAS_QAxContainer = True
except ImportError:
    HAS_QAxContainer = False


class FieldDisplayPanel(QWidget):
    """提取字段显示面板（左侧）"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.extracted_fields = {}
        self.init_ui()

    def init_ui(self):
        """初始化UI"""
        layout = QVBoxLayout()
        layout.setSpacing(8)

        # 标题
        title = QLabel("从运输鉴定报告提取的字段")
        title.setFont(QFont("Microsoft YaHei", 12, QFont.Bold))
        layout.addWidget(title)

        # 选择文件按钮
        self.btn_select_file = QPushButton("选择运输鉴定报告")
        self.btn_select_file.clicked.connect(self.select_report_file)
        layout.addWidget(self.btn_select_file)

        # 文件路径显示
        self.file_path_label = QLabel("未选择文件")
        self.file_path_label.setStyleSheet("color: #666; font-size: 11px;")
        self.file_path_label.setWordWrap(True)
        layout.addWidget(self.file_path_label)

        # 分隔线
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setStyleSheet("color: #ddd;")
        layout.addWidget(line)

        # 商品编码搜索
        codes_group = QGroupBox("商品编码搜索")
        codes_layout = QVBoxLayout()

        self.product_search_edit = QLineEdit()
        self.product_search_edit.setPlaceholderText("输入产品内编搜索...")
        self.product_search_edit.textChanged.connect(self.on_product_search_changed)

        self.products_table = QTableWidget()
        self.products_table.setColumnCount(4)
        self.products_table.setHorizontalHeaderLabels(['产品内编', '报关名称', '产品外观', '新商品编码'])
        self.products_table.horizontalHeader().setStretchLastSection(True)
        self.products_table.setMaximumHeight(150)
        self.products_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.products_table.setEditTriggers(QAbstractItemView.NoEditTriggers)

        codes_layout.addWidget(self.product_search_edit)
        codes_layout.addWidget(self.products_table)
        codes_group.setLayout(codes_layout)
        layout.addWidget(codes_group)

        # 字段显示区域（滚动）
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)

        fields_widget = QWidget()
        self.fields_layout = QVBoxLayout()
        self.fields_layout.setSpacing(6)

        # 字段标签
        self.field_labels = {}
        field_names = [
            ('样品名称中文', '样品名称(中文)'),
            ('样品名称英文', '样品名称(英文)'),
            ('报告编号', '报告编号'),
            ('样品编号', '样品编号'),
            ('外观性状中文', '外观/性状(中文)'),
            ('外观性状英文', '外观/性状(英文)'),
            ('委托单位', '委托单位'),
            ('生产单位', '生产单位'),
            ('鉴定结论', '鉴定结论'),
            ('运输类型', '运输类型'),
        ]

        for key, label_text in field_names:
            field_group = self._create_field_group(key, label_text)
            self.fields_layout.addWidget(field_group)

        self.fields_layout.addStretch()
        fields_widget.setLayout(self.fields_layout)

        scroll.setWidget(fields_widget)
        layout.addWidget(scroll, 1)

        self.setLayout(layout)

    def set_product_codes_loader(self, loader: ProductCodesLoader):
        """设置商品编码加载器"""
        self.product_codes_loader = loader
        self.update_products_table(loader.get_all_products())

    def update_products_table(self, products: list):
        """更新商品编码表格"""
        self.products_table.setRowCount(0)
        for product in products[:50]:  # 最多显示50条
            row = self.products_table.rowCount()
            self.products_table.insertRow(row)
            self.products_table.setItem(row, 0, QTableWidgetItem(product.get('产品内编', '')))
            self.products_table.setItem(row, 1, QTableWidgetItem(product.get('报关名称', '')))
            self.products_table.setItem(row, 2, QTableWidgetItem(product.get('产品外观', '')))
            self.products_table.setItem(row, 3, QTableWidgetItem(product.get('新商品编码', '')))

    def on_product_search_changed(self, text: str):
        """产品搜索"""
        if hasattr(self, 'product_codes_loader') and text:
            results = self.product_codes_loader.search_by_name(text)
            self.update_products_table(results)
        elif hasattr(self, 'product_codes_loader'):
            self.update_products_table(self.product_codes_loader.get_all_products())

    def _create_field_group(self, key: str, label_text: str) -> QWidget:
        """创建单个字段显示组"""
        group = QWidget()
        group_layout = QHBoxLayout()
        group_layout.setContentsMargins(0, 2, 0, 2)

        # 标签
        label = QLabel(f"{label_text}:")
        label.setMinimumWidth(110)
        label.setFont(QFont("Microsoft YaHei", 10))
        group_layout.addWidget(label)

        # 值显示（点击复制）
        value_edit = ClickableLineEdit()
        value_edit.setReadOnly(True)
        value_edit.setPlaceholderText("待提取...")
        value_edit.setFont(QFont("Microsoft YaHei", 10))
        group_layout.addWidget(value_edit, 1)

        # 复制按钮
        btn_copy = QPushButton("复制")
        btn_copy.setFixedWidth(50)
        btn_copy.clicked.connect(lambda: self.copy_field(key))
        group_layout.addWidget(btn_copy)

        group.setLayout(group_layout)
        self.field_labels[key] = value_edit

        return group

    def select_report_file(self):
        """选择运输鉴定报告文件"""
        filepath, _ = QFileDialog.getOpenFileName(
            self,
            "选择运输鉴定报告",
            "",
            "运输鉴定报告 (*.pdf *.docx *.doc);;PDF文件 (*.pdf);;Word文件 (*.docx *.doc);;所有文件 (*)"
        )

        if filepath:
            self.load_report(filepath)

    def load_report(self, filepath: str):
        """加载鉴定报告并提取字段"""
        try:
            extractor = ReportExtractor()
            fields = extractor.extract_from_file(filepath)

            self.extracted_fields = fields
            self.file_path_label.setText(os.path.basename(filepath))
            self.file_path_label.setToolTip(filepath)

            # 更新字段显示
            for key, label in self.field_labels.items():
                value = fields.get(key, '')
                label.setText(value)

            QMessageBox.information(self, "成功", f"已从报告中提取 {len(fields)} 个字段")

        except Exception as e:
            QMessageBox.warning(self, "错误", f"提取失败: {str(e)}")

    def copy_field(self, key: str):
        """复制指定字段到剪贴板"""
        value = self.extracted_fields.get(key, '')
        if value:
            QApplication.clipboard().setText(value)
            self.field_labels[key].setStyleSheet("background-color: #e8f5e9;")
            QApplication.processEvents()
            self.field_labels[key].setStyleSheet("")

    def get_extracted_fields(self) -> dict:
        """获取所有提取的字段"""
        return self.extracted_fields.copy()


class ClickableLineEdit(QLineEdit):
    """可点击复制的文本框"""
    clicked = pyqtSignal()

    def __init__(self, text='', parent=None):
        super().__init__(text, parent)
        self.setStyleSheet("""
            QLineEdit {
                background-color: #fafafa;
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
        if event.button() == Qt.LeftButton and self.text():
            QApplication.clipboard().setText(self.text())
            self.clicked.emit()
        super().mousePressEvent(event)


class EmbeddedWordWidget(QWidget):
    """嵌入的Word编辑控件（支持QAxContainer或回退到外部Word）"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.word_document = None
        self.current_filepath = None
        self.is_modified = False
        self.word_ax = None
        self.init_ui()

    def init_ui(self):
        """初始化UI"""
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)

        # 占位提示
        self.placeholder_label = QLabel("请加载模板文件\n\n点击上方'加载模板'按钮选择Word文档")
        self.placeholder_label.setAlignment(Qt.AlignCenter)
        self.placeholder_label.setStyleSheet("""
            QLabel {
                color: #999;
                font-size: 14px;
                padding: 50px;
                background-color: #f5f5f5;
                border: 2px dashed #ddd;
                border-radius: 8px;
            }
        """)
        layout.addWidget(self.placeholder_label)

        # Word控件占位符
        self.word_container = QWidget()
        self.word_container.setStyleSheet("background-color: white;")
        self.word_container_layout = QVBoxLayout()
        self.word_container_layout.setContentsMargins(0, 0, 0, 0)
        self.word_container.setLayout(self.word_container_layout)

        # 创建QAxWidget（如果可用）
        if HAS_QAxContainer:
            try:
                from PyQt5.QtAxContainer import QAxWidget
                self.word_ax = QAxWidget("{000209FF-0000-0000-C000-000000000046}")  # Word's CLSID
                self.word_container_layout.addWidget(self.word_ax)
                self.word_ax.setControl("")
                self.word_ax.setVisible(False)
            except Exception as e:
                self.word_ax = None
                print(f"Word控件创建失败: {e}")
        else:
            self.word_ax = None

        # 默认隐藏容器
        self.word_container.setVisible(False)
        layout.addWidget(self.word_container)

        self.setLayout(layout)

    def load_document(self, filepath: str) -> bool:
        """加载Word文档"""
        if not os.path.exists(filepath):
            return False

        try:
            self.current_filepath = filepath

            if self.word_ax and HAS_QAxContainer:
                # 使用QAxWidget嵌入Word
                self.placeholder_label.setVisible(False)
                self.word_container.setVisible(True)
                self.word_ax.setControl(filepath)
                self.word_ax.setVisible(True)
                self.word_document = self.word_ax.querySubObject("Documents")
                return True
            else:
                # 回退：使用外部Word
                self.placeholder_label.setVisible(False)
                self.word_container.setVisible(True)

                # 显示文档路径提示
                self.doc_path_label = QLabel(f"文档已加载:\n{os.path.basename(filepath)}")
                self.doc_path_label.setAlignment(Qt.AlignCenter)
                self.doc_path_label.setStyleSheet("""
                    QLabel {
                        color: #333;
                        font-size: 13px;
                        padding: 30px;
                        background-color: #e8f5e9;
                        border: 2px solid #4CAF50;
                        border-radius: 8px;
                    }
                """)
                self.word_container_layout.addWidget(self.doc_path_label)

                # 用外部Word打开
                os.startfile(filepath)
                return True

        except Exception as e:
            self.placeholder_label.setText(f"加载失败:\n{str(e)}")
            self.placeholder_label.setVisible(True)
            self.word_container.setVisible(False)
            return False

    def close_document(self):
        """关闭当前文档"""
        if self.word_ax and self.current_filepath:
            try:
                doc = self.word_ax.querySubObject("ActiveDocument")
                if doc:
                    doc.close()
            except:
                pass

            self.word_ax.setControl("")
            self.word_ax.setVisible(False)

        # 清理回退模式的控件
        if hasattr(self, 'doc_path_label'):
            self.doc_path_label.deleteLater()
            delattr(self, 'doc_path_label')

        self.word_container.setVisible(False)
        self.placeholder_label.setVisible(True)
        self.current_filepath = None
        self.word_document = None

    def is_word_available(self) -> bool:
        """检查Word是否可用（嵌入模式）"""
        return self.word_ax is not None

    def is_external_word_available(self) -> bool:
        """检查外部Word是否可用"""
        return True  # os.startfile always works on Windows if file association exists


class TemplateEditorPanel(QWidget):
    """模板编辑面板（右侧）"""

    template_loaded = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.current_template_path = None
        self.init_ui()

    def init_ui(self):
        """初始化UI"""
        layout = QVBoxLayout()
        layout.setSpacing(10)

        # 标题
        title = QLabel("模板编辑区")
        title.setFont(QFont("Microsoft YaHei", 12, QFont.Bold))
        layout.addWidget(title)

        # 模板类型选择
        type_group = QGroupBox("选择模板类型")
        type_layout = QVBoxLayout()

        # MSDS选项
        msds_layout = QHBoxLayout()
        msds_label = QLabel("MSDS:")
        msds_layout.addWidget(msds_label)

        self.msds_combo = QComboBox()
        self.msds_combo.addItems(['请选择...', 'MSDS中文模板', 'MSDS英文模板'])
        self.msds_combo.currentIndexChanged.connect(self.on_template_type_changed)
        msds_layout.addWidget(self.msds_combo, 1)
        msds_layout.addStretch()
        type_layout.addLayout(msds_layout)

        # LOI选项
        loi_layout = QHBoxLayout()
        loi_label = QLabel("LOI:")
        loi_layout.addWidget(loi_label)

        self.loi_combo = QComboBox()
        self.loi_combo.addItems(['请选择...', 'LOI非危险品保函', 'LOI液体保函'])
        self.loi_combo.currentIndexChanged.connect(self.on_template_type_changed)
        loi_layout.addWidget(self.loi_combo, 1)
        loi_layout.addStretch()
        type_layout.addLayout(loi_layout)

        type_group.setLayout(type_layout)
        layout.addWidget(type_group)

        # 模板文件选择
        template_group = QGroupBox("加载模板文件")
        template_layout = QHBoxLayout()

        self.template_path_edit = QLineEdit()
        self.template_path_edit.setReadOnly(True)
        self.template_path_edit.setPlaceholderText("选择模板文件...")

        self.btn_browse_template = QPushButton("浏览...")
        self.btn_browse_template.clicked.connect(self.browse_template)

        template_layout.addWidget(self.template_path_edit, 1)
        template_layout.addWidget(self.btn_browse_template)

        template_group.setLayout(template_layout)
        layout.addWidget(template_group)

        # 嵌入Word编辑区
        word_group = QGroupBox("Word编辑器")
        word_layout = QVBoxLayout()

        self.word_widget = EmbeddedWordWidget()
        word_layout.addWidget(self.word_widget)

        word_group.setLayout(word_layout)
        layout.addWidget(word_group, 1)

        # 按钮区域
        btn_layout = QHBoxLayout()

        self.btn_load_template = QPushButton("加载模板")
        self.btn_load_template.clicked.connect(self.load_template)

        self.btn_save_as = QPushButton("另存为...")
        self.btn_save_as.clicked.connect(self.save_as)
        self.btn_save_as.setEnabled(False)

        btn_layout.addWidget(self.btn_load_template)
        btn_layout.addWidget(self.btn_save_as)
        btn_layout.addStretch()

        layout.addLayout(btn_layout)

        self.setLayout(layout)

    def on_template_type_changed(self, index):
        """模板类型改变"""
        sender = self.sender()
        config = get_config()

        # 根据选择的类型自动填充路径
        if sender == self.msds_combo:
            self.loi_combo.setCurrentIndex(0)
            if index == 1:  # MSDS中文
                self.template_path_edit.setText(config.get_path('msds_template_cn'))
            elif index == 2:  # MSDS英文
                self.template_path_edit.setText(config.get_path('msds_template_en'))
        elif sender == self.loi_combo:
            self.msds_combo.setCurrentIndex(0)
            if index == 1:  # LOI非危险品
                self.template_path_edit.setText(config.get_path('loi_template_non_hazard'))
            elif index == 2:  # LOI液体
                self.template_path_edit.setText(config.get_path('loi_template_liquid'))

        self.current_template_path = self.template_path_edit.text()

    def browse_template(self):
        """浏览选择模板文件"""
        config = get_config()
        default_dir = os.path.dirname(config.get_path('msds_template_cn'))
        if not os.path.exists(default_dir):
            default_dir = os.path.dirname(config.get_base_dir())

        filepath, _ = QFileDialog.getOpenFileName(
            self,
            "选择模板文件",
            default_dir,
            "Word文档 (*.doc *.docx);;所有文件 (*)"
        )

        if filepath:
            self.template_path_edit.setText(filepath)
            self.current_template_path = filepath
            self.msds_combo.setCurrentIndex(0)
            self.loi_combo.setCurrentIndex(0)

    def load_template(self):
        """加载模板文件"""
        if not self.current_template_path:
            QMessageBox.warning(self, "提示", "请先选择模板文件")
            return

        if not os.path.exists(self.current_template_path):
            QMessageBox.warning(self, "错误", f"模板文件不存在:\n{self.current_template_path}")
            return

        # 加载模板（内部自动处理嵌入或外部Word）
        success = self.word_widget.load_document(self.current_template_path)

        if success:
            self.btn_save_as.setEnabled(True)
            self.template_loaded.emit(self.current_template_path)

            if not self.word_widget.is_word_available():
                QMessageBox.information(
                    self,
                    "提示",
                    "Word模板已用外部Word打开。\n"
                    "编辑完成后关闭Word，文件会自动保存。\n\n"
                    "如果需要另存为，请使用'另存为'按钮。"
                )

    def save_as(self):
        """另存为"""
        if not self.current_template_path:
            return

        ext = os.path.splitext(self.current_template_path)[1]
        default_name = f"填写好的{os.path.basename(self.current_template_path)}"

        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "另存为",
            default_name,
            f"Word文档 (*{ext});;所有文件 (*)"
        )

        if output_path:
            try:
                # 直接复制原文件到目标位置
                shutil.copy2(self.current_template_path, output_path)

                # 如果Word控件可用，尝试保存
                if self.word_widget.word_ax:
                    try:
                        doc = self.word_widget.word_ax.querySubObject("ActiveDocument")
                        if doc:
                            doc.saveAs(output_path)
                    except:
                        pass

                QMessageBox.information(self, "成功", f"文件已保存到:\n{output_path}")
            except Exception as e:
                QMessageBox.warning(self, "错误", f"保存失败: {str(e)}")

    def get_template_type(self) -> str:
        """获取当前模板类型"""
        if self.msds_combo.currentIndex() > 0:
            return 'msds'
        elif self.loi_combo.currentIndex() > 0:
            return 'loi'
        return 'unknown'


class Phase3Widget(QWidget):
    """Phase 3 主界面 - MSDS/LOI模板填写"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.project_manager = None
        self.product_codes_loader = ProductCodesLoader()
        self.current_pi_no = None
        self.init_ui()
        self.load_product_codes()

    def init_ui(self):
        """初始化UI"""
        layout = QHBoxLayout()
        layout.setSpacing(15)

        # 左侧：提取的字段面板
        self.field_panel = FieldDisplayPanel()
        self.field_panel.setMinimumWidth(420)
        layout.addWidget(self.field_panel, 1)

        # 右侧：模板编辑面板
        self.template_panel = TemplateEditorPanel()
        self.template_panel.setMinimumWidth(600)
        layout.addWidget(self.template_panel, 2)

        self.setLayout(layout)

    def load_product_codes(self):
        """加载商品编码表"""
        config = get_config()
        codes_path = config.get_path('product_codes_excel')

        if os.path.exists(codes_path):
            if self.product_codes_loader.load(codes_path):
                self.field_panel.set_product_codes_loader(self.product_codes_loader)
                print(f"商品编码表已加载: {self.product_codes_loader.get_count()} 条产品")
            else:
                print(f"商品编码表加载失败: {codes_path}")
        else:
            print(f"商品编码表文件不存在: {codes_path}")

    def set_pi_no(self, pi_no: str):
        """设置当前PI号"""
        self.current_pi_no = pi_no

    def set_project_manager(self, pm):
        """设置项目管理器"""
        self.project_manager = pm

    def get_extracted_fields(self) -> dict:
        """获取提取的字段"""
        return self.field_panel.get_extracted_fields()

    def load_report_and_display(self, filepath: str):
        """加载鉴定报告并显示"""
        self.field_panel.load_report(filepath)