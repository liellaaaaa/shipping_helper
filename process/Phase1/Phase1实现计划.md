# ShippingHelper Phase 1 实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task.

**Goal:** 实现Phase 1订单数据提取与包装计算功能，包括订单表解析、PI文件提取、数据汇聚、包装计算、左右分栏UI。

**Architecture:** 复用shipping_helper_old中的成熟模块（order_parser, pi_extractor, merger, package_calculator, code_matcher），迁移到新项目结构，添加PyQt5界面。

**Tech Stack:** Python 3.x, PyQt5, xlrd, openpyxl, JSON

---

## 文件结构

```
shipping_helper/
├── core/
│   ├── __init__.py
│   ├── order_parser.py        # 从old迁移 + 适配
│   ├── pi_extractor.py        # 从old迁移 + 适配
│   ├── code_matcher.py        # 从old迁移
│   ├── package_calculator.py  # 从old迁移
│   └── merger.py              # 从old迁移
├── ui/
│   ├── __init__.py
│   └── main_window.py         # 新建：PyQt5左右分栏窗口
├── knowledge/
│   ├── products_knowledge.json  # 新建：商品编码知识库
│   └── packaging_data.json      # 新建：包装资料
├── config/
│   └── settings.json          # 新建：配置文件
├── main.py                   # 新建：入口文件
└── requirements.txt          # 新建：依赖
```

---

## 任务分解

### Task 1: 创建项目目录结构

**Files:**
- Create: `shipping_helper/core/__init__.py`
- Create: `shipping_helper/ui/__init__.py`
- Create: `shipping_helper/knowledge/.gitkeep`
- Create: `shipping_helper/config/.gitkeep`

- [ ] **Step 1: Create directories**

```bash
mkdir -p shipping_helper/core shipping_helper/ui shipping_helper/knowledge shipping_helper/config
```

- [ ] **Step 2: Create __init__.py files**

```bash
touch shipping_helper/core/__init__.py shipping_helper/ui/__init__.py
```

- [ ] **Step 3: Commit**

```bash
git add shipping_helper/core/__init__.py shipping_helper/ui/__init__.py
git commit -m "chore: initialize project structure"
```

---

### Task 2: 迁移 order_parser.py

**Files:**
- Create: `shipping_helper/core/order_parser.py`
- Source: `shipping_helper_old/core/order_parser.py`
- Test: `shipping_helper/tests/core/test_order_parser.py`

- [ ] **Step 1: Copy from old project**

```python
# -*- coding: utf-8 -*-
"""
外贸销售订单表粘贴解析模块
从在线文档复制的一行数据（Tab分隔）进行解析
"""

import re


class OrderParser:
    """解析外贸销售订单表粘贴的数据"""

    COLUMNS = [
        '业务员', '客户编号', '内部编号', '产品中文名', '报关名称', '规格kg',
        '订单量kg', '是否调价', '有无样品', '订单要求', '交货日期', '审核',
        '销售区域', '订单号', '出货抬头', '单据类型', '跟单员', '下单日期',
        '确认下单', '生产交期', '出货渠道', '出货方式', '规格异常'
    ]

    def __init__(self):
        self.data = {}

    def parse(self, text: str) -> dict:
        """解析粘贴的文本"""
        text = text.strip()
        if not text:
            return {}

        has_tabs = '\t' in text
        newline_count = text.count('\n')

        if has_tabs:
            lines = text.split('\n')
            if lines and not lines[0].strip():
                lines = lines[1:]
            if lines:
                first_line = lines[0]
                parts = first_line.split('\t')
            else:
                parts = []
        elif newline_count >= 10:
            lines = text.split('\n')
            parts = [line for line in lines if line.strip()]
        else:
            parts = text.split()

        parts = [p.strip() for p in parts]

        if parts and parts[0].isdigit():
            parts = parts[1:]

        parts = self._fix_column_offset(parts)

        if parts and parts[0] == '业务员':
            parts = parts[1:]

        if len(parts) > len(self.COLUMNS):
            parts = self._collapse_extra_fields(parts)

        while len(parts) < len(self.COLUMNS):
            parts.append('')

        parts = parts[:len(self.COLUMNS)]

        for i, col in enumerate(self.COLUMNS):
            self.data[col] = parts[i] if i < len(parts) else ''

        return self.data

    def _fix_column_offset(self, parts: list) -> list:
        """智能修复列偏移"""
        if len(parts) <= 4:
            return parts

        col_e = parts[4] if len(parts) > 4 else ''

        if col_e.isdigit():
            col_f = parts[5] if len(parts) > 5 else ''
            if col_f.isdigit():
                return parts
            return parts[:4] + [''] + parts[4:]

        return parts

    def _collapse_extra_fields(self, parts: list) -> list:
        """合并嵌入换行导致的额外字段"""
        if len(parts) <= len(self.COLUMNS):
            return parts

        result = []
        i = 0
        while i < len(parts):
            if i < 9:
                result.append(parts[i])
                i += 1
            elif i == 9:
                merged = parts[i]
                j = i + 1
                while j < len(parts):
                    next_part = parts[j].strip()
                    if j >= 10 and self._looks_like_real_field_start(next_part):
                        break
                    if not next_part and j + 1 < len(parts) and self._looks_like_real_field_start(parts[j + 1].strip()):
                        break
                    merged += '\n' + next_part
                    j += 1
                result.append(merged.strip())
                i = j
            else:
                result.append(parts[i])
                i += 1

        return result

    def _looks_like_real_field_start(self, text: str) -> bool:
        """判断是否是真正的新字段开始"""
        import re
        text = text.strip()
        if not text:
            return False

        real_markers = [
            '已审核', '待审核', '正常单', '外贸部', 'HT20', '宏昊抬头', '刘洁婷', '下单', '出货渠道',
        ]

        for marker in real_markers:
            if text == marker:
                return True

        if re.match(r'^\d{4}\s*年\s*\d+月', text) and '：' not in text:
            return True

        if re.match(r'^\d{1,2}\s*月\s*\d{1,2}\s*日', text):
            return True

        if re.match(r'^\d+$', text) and len(text) <= 6:
            return True

        return False

    def get_pi号(self) -> str:
        return self.data.get('订单号', '')

    def get_internal_code(self) -> str:
        return self.data.get('内部编号', '')

    def get_all_data(self) -> dict:
        return self.data.copy()

    def validate(self) -> tuple:
        if not self.get_pi号():
            return False, "订单号（PI号）不能为空"
        if not self.get_internal_code():
            return False, "内部编号不能为空"
        return True, ""
```

- [ ] **Step 2: Write test file**

```python
# shipping_helper/tests/core/test_order_parser.py
import pytest
from core.order_parser import OrderParser


def test_parse_tab_separated():
    parser = OrderParser()
    text = "张三\tWA524\tK70-EB\t皂洗剂\t\t50kg\t1000\t否\t无\t液体：50kg蓝桶细口，共10桶\n粉：25kg纸桶，共5袋\t2026-06-01\t已审核\t华南\tHT260301\t宏昊\tPI\t李四\t2026-03-01\t是\t2026-06-15\t快递\t件\t正常"
    result = parser.parse(text)
    assert result['业务员'] == '张三'
    assert result['客户编号'] == 'WA524'
    assert result['内部编号'] == 'K70-EB'


def test_fix_column_offset():
    parser = OrderParser()
    # 报关名称位置是纯数字的情况
    text = "张三\tWA524\tK70-EB\t皂洗剂\t1000\t50kg\t1000\t否\t无\t备注\t2026-06-01"
    result = parser.parse(text)
    assert result['报关名称'] == ''


def test_empty_text():
    parser = OrderParser()
    result = parser.parse("")
    assert result == {}


def test_validate_empty_pi():
    parser = OrderParser()
    parser.parse("张三\tWA524\t\t皂洗剂")
    is_valid, msg = parser.validate()
    assert is_valid == False
    assert "订单号" in msg
```

- [ ] **Step 3: Run test to verify it fails**

```bash
pytest shipping_helper/tests/core/test_order_parser.py -v
```

- [ ] **Step 4: Commit**

```bash
git add shipping_helper/core/order_parser.py shipping_helper/tests/core/test_order_parser.py
git commit -m "feat: migrate order_parser.py from old project"
```

---

### Task 3: 迁移 pi_extractor.py

**Files:**
- Create: `shipping_helper/core/pi_extractor.py`
- Source: `shipping_helper_old/core/pi_extractor.py`

- [ ] **Step 1: Copy and adapt from old project**

```python
# shipping_helper/core/pi_extractor.py
# -*- coding: utf-8 -*-
"""
Proforma Invoice 解析模块
解析自由文本格式的PI文件（.xls格式）
"""

import xlrd
import re
from datetime import datetime


class PIExtractor:
    """解析 Proforma Invoice 文件"""

    def __init__(self):
        self.data = {}

    def parse_file(self, filepath: str) -> dict:
        """解析Proforma Invoice文件"""
        wb = xlrd.open_workbook(filepath, encoding_override='utf-8')
        ws = wb.sheet_by_index(0)

        full_text = []
        for i in range(ws.nrows):
            row_text = []
            for j in range(ws.ncols):
                val = ws.cell_value(i, j)
                if val:
                    row_text.append(str(val).strip())
            full_text.append(row_text)

        self._extract_consignee(full_text)
        self._extract_date(full_text)
        self._extract_pi_no(full_text)
        self._extract_goods(full_text)
        self._extract_quantity(full_text)
        self._extract_hs_code(full_text)
        self._extract_destination(full_text)
        self._extract_packing(full_text)

        return self.data

    def _find_text(self, full_text, keywords):
        """在文本中查找包含关键词的行"""
        for row in full_text:
            for cell in row:
                if any(kw.lower() in cell.lower() for kw in keywords):
                    return row
        return None

    def _extract_consignee(self, full_text):
        """提取收货人信息"""
        for i, row in enumerate(full_text):
            for cell in row:
                if 'Name :' in cell:
                    self.data['收货人'] = cell.replace('Name :', '').strip()
                    break
            if '收货人' in self.data:
                break

        if '收货人' not in self.data:
            for row in full_text:
                for i, cell in enumerate(row):
                    if 'Name :' in str(cell):
                        if i + 1 < len(row) and row[i + 1]:
                            self.data['收货人'] = str(row[i + 1]).strip()
                            break
                if '收货人' in self.data:
                    break

        for row in full_text:
            row_str = ' '.join([str(c) for c in row])
            if 'Add.:' in row_str or 'Add:' in row_str:
                for cell in row:
                    if str(cell).strip().lower().startswith('add.:'):
                        addr = str(cell).replace('Add.:', '').replace('Add:', '').strip()
                        for j in range(row.index(cell) + 1, len(row)):
                            if row[j]:
                                addr += ' ' + str(row[j]).strip()
                        self.data['地址'] = addr.strip()
                        break
            if '地址' in self.data:
                break

    def _extract_date(self, full_text):
        """提取日期"""
        for row in full_text:
            for cell in row:
                cell_str = str(cell).strip()
                if 'Date :' in cell_str or cell_str.startswith('Date:'):
                    date_str = cell_str.replace('Date :', '').replace('Date:', '').strip()
                    if date_str and not date_str.replace('.', '').isdigit():
                        self.data['日期'] = date_str
                        break
                    idx = row.index(cell)
                    for j in range(idx + 1, len(row)):
                        next_val = str(row[j]).strip()
                        if next_val and '20' in next_val:
                            self.data['日期'] = next_val
                            break
            if '日期' in self.data:
                break

        if '日期' not in self.data:
            for row in full_text:
                for cell in row:
                    cell_str = str(cell).strip()
                    if '20' in cell_str and ('-' in cell_str or '/' in cell_str):
                        match = re.search(r'20\d[/-]\d[/-]\d', cell_str)
                        if match:
                            self.data['日期'] = match.group()
                            break
                if '日期' in self.data:
                    break

    def _extract_pi_no(self, full_text):
        """提取PI号"""
        for row in full_text:
            row_str = ' '.join(row)
            if 'PI No.:' in row_str or 'PI No.:' in row_str:
                for cell in row:
                    if 'HT' in cell and ('PI' in cell or len(cell) > 5):
                        match = re.search(r'HT\d+[A-Z0-9]+', cell)
                        if match:
                            self.data['PI号'] = match.group()
                            break
                    match = re.search(r'HT\d+[A-Z0-9]+', cell)
                    if match:
                        self.data['PI号'] = match.group()
                        break
            if 'PI号' in self.data:
                break

        if 'PI号' not in self.data:
            all_text = ' '.join([' '.join(row) for row in full_text])
            match = re.search(r'HT\d+[A-Z0-9]+', all_text)
            if match:
                self.data['PI号'] = match.group()

    def _extract_goods(self, full_text):
        """提取商品信息"""
        in_goods_section = False
        for row in full_text:
            row_str = ' '.join(row)
            if 'DESCRIPTION OF GOODS' in row_str:
                in_goods_section = True
                continue
            if in_goods_section and row:
                if len(row) >= 2:
                    has_number = any(
                        isinstance(x, (int, float)) or (isinstance(x, str) and x.replace('.', '').isdigit())
                        for x in row
                    )
                    if has_number and len(row[0]) > 3:
                        self.data['品名英文'] = row[0].strip()
                        break

    def _extract_quantity(self, full_text):
        """提取数量、单价、金额"""
        for row in full_text:
            if len(row) >= 4:
                try:
                    qty = row[1]
                    cif = row[2]
                    amount = row[3]
                    if isinstance(qty, (int, float)) and isinstance(amount, (int, float)):
                        self.data['数量'] = qty
                        self.data['单价'] = cif
                        self.data['金额'] = amount
                        break
                except (IndexError, ValueError):
                    continue

        for row in full_text:
            if 'TOTAL' in ' '.join(row) and len(row) >= 3:
                try:
                    qty = row[1]
                    amount = row[2]
                    if isinstance(qty, (int, float)) and isinstance(amount, (int, float)):
                        if '数量' not in self.data:
                            self.data['数量'] = qty
                        if '金额' not in self.data:
                            self.data['金额'] = amount
                        break
                except:
                    continue

    def _extract_hs_code(self, full_text):
        """提取H.S. Code"""
        for row in full_text:
            for cell in row:
                if 'H.S. Code' in cell or 'H.S. Code' in cell:
                    match = re.search(r'(\d{10})', cell)
                    if match:
                        self.data['H.S.Code'] = match.group(1)
                        break
            if 'H.S.Code' in self.data:
                break

    def _extract_destination(self, full_text):
        """提取卸货港"""
        for row in full_text:
            for cell in row:
                if 'Destination' in cell and ':' in cell:
                    dest = cell.split(':', 1)[1].strip()
                    if dest:
                        self.data['卸货港'] = dest
                        break
            if '卸货港' in self.data:
                break

    def _extract_packing(self, full_text):
        """提取包装说明"""
        for row in full_text:
            for cell in row:
                if 'Packing' in cell and ':' in cell:
                    packing = cell.split(':', 1)[1].strip()
                    if packing:
                        self.data['包装说明'] = packing
                        break
            if '包装说明' in self.data:
                break

        if '包装说明' not in self.data:
            for row in full_text:
                row_str = ' '.join(row)
                if 'Packing:' in row_str or 'export standard' in row_str.lower():
                    for cell in row:
                        if 'export standard' in cell.lower() or 'kg' in cell.lower():
                            self.data['包装说明'] = cell.strip()
                            break
                    if '包装说明' not in self.data:
                        self.data['包装说明'] = row_str.strip()
                    break

    def get_all_data(self) -> dict:
        """获取所有解析数据"""
        return self.data.copy()
```

- [ ] **Step 2: Commit**

```bash
git add shipping_helper/core/pi_extractor.py
git commit -m "feat: migrate pi_extractor.py from old project"
```

---

### Task 4: 迁移其他核心模块

**Files:**
- Create: `shipping_helper/core/code_matcher.py`
- Create: `shipping_helper/core/package_calculator.py`
- Create: `shipping_helper/core/merger.py`

- [ ] **Step 1: Copy code_matcher.py from old project**

从 `shipping_helper_old/core/code_matcher.py` 复制完整内容

- [ ] **Step 2: Copy package_calculator.py from old project**

从 `shipping_helper_old/core/package_calculator.py` 复制完整内容

- [ ] **Step 3: Copy merger.py from old project**

从 `shipping_helper_old/core/merger.py` 复制完整内容

- [ ] **Step 4: Commit**

```bash
git add shipping_helper/core/code_matcher.py shipping_helper/core/package_calculator.py shipping_helper/core/merger.py
git commit -m "feat: migrate code_matcher, package_calculator, merger from old project"
```

---

### Task 5: 创建知识库JSON

**Files:**
- Create: `shipping_helper/knowledge/products_knowledge.json`
- Create: `shipping_helper/knowledge/packaging_data.json`

- [ ] **Step 1: Create products_knowledge.json**

```json
{
  "version": "2024.12",
  "products": {
    "K70-EB": {
      "报关名称": "皂洗剂",
      "海关编码": "3402420000",
      "成分": "表面活性剂混合物",
      "产品外观": "黄色粉末",
      "桶类型": ["50kg蓝桶", "125kg新款胶桶"],
      "卡板容量": {
        "50kg蓝桶(细口)": {"1.1*1.1": 18},
        "50kg蓝桶(大口)": {"1.1*1.1": 18},
        "125kg新款胶桶": {"1.1*1.1": 5}
      }
    },
    "K80-EB": {
      "报关名称": "匀染剂",
      "海关编码": "3402420000",
      "成分": "表面活性剂混合物",
      "产品外观": "棕黑色液体",
      "桶类型": ["50kg蓝桶", "125kg新款胶桶"],
      "卡板容量": {
        "50kg蓝桶(细口)": {"1.1*1.1": 18},
        "125kg新款胶桶": {"1.1*1.1": 5}
      }
    }
  },
  "customers": {
    "WA524": {
      "名称": "TOA-DOVECHEM INDUSTRIES CO., LTD"
    }
  }
}
```

- [ ] **Step 2: Create packaging_data.json**

```json
{
  "version": "2020-6-10",
  "packages": [
    {"name": "30kg蓝桶", "dims": "250*250*430mm", "cbm": 0.026, "tare_kg": 2.0, "gross_kg": 32.0, "net_kg": 30.0},
    {"name": "25kg/包", "dims": "200*500*800mm", "cbm": 0.028, "tare_kg": 0.5, "gross_kg": 25.5, "net_kg": 25.0},
    {"name": "50kg蓝桶(细口)", "dims": "395*395*585mm", "cbm": 0.09, "tare_kg": 2.5, "gross_kg": 52.5, "net_kg": 50.0},
    {"name": "50kg蓝桶(大口)", "dims": "330*390*590mm", "cbm": 0.08, "tare_kg": 2.5, "gross_kg": 52.5, "net_kg": 50.0},
    {"name": "125kg新款胶桶", "dims": "510*510*810mm", "cbm": 0.21, "tare_kg": 6.0, "gross_kg": 131.0, "net_kg": 125.0},
    {"name": "200kg双环闭口桶", "dims": "590*590*930mm", "cbm": 0.31, "tare_kg": 10.0, "gross_kg": 210.0, "net_kg": 200.0},
    {"name": "1吨桶", "dims": "1200*1000*1150mm", "cbm": 1.38, "tare_kg": 58.0, "gross_kg": 1058.0, "net_kg": 1000.0}
  ],
  "pallets": [
    {"name": "1.0*1.0m卡板", "size_m": "1.0*1.0", "tare_kg": 17.0, "cbm": 0.15},
    {"name": "1.1*1.1m卡板", "size_m": "1.1*1.1", "tare_kg": 18.5, "cbm": 0.2}
  ]
}
```

- [ ] **Step 3: Commit**

```bash
git add shipping_helper/knowledge/products_knowledge.json shipping_helper/knowledge/packaging_data.json
git commit -m "feat: add knowledge base JSON files"
```

---

### Task 6: 创建配置和依赖文件

**Files:**
- Create: `shipping_helper/config/settings.json`
- Create: `shipping_helper/requirements.txt`

- [ ] **Step 1: Create settings.json**

```json
{
  "app": {
    "name": "ShippingHelper",
    "version": "1.0.0"
  },
  "shipper": "广东宏昊化工有限公司",
  "knowledge_base": {
    "products": "knowledge/products_knowledge.json",
    "packaging": "knowledge/packaging_data.json"
  },
  "defaults": {
    "pallet_type": "1.1*1.1m卡板"
  }
}
```

- [ ] **Step 2: Create requirements.txt**

```
PyQt5>=5.15.0
xlrd>=2.0.1
openpyxl>=3.0.0
```

- [ ] **Step 3: Commit**

```bash
git add shipping_helper/config/settings.json shipping_helper/requirements.txt
git commit -m "chore: add settings.json and requirements.txt"
```

---

### Task 7: 创建主窗口UI

**Files:**
- Create: `shipping_helper/ui/main_window.py`

- [ ] **Step 1: Create main_window.py with left-right layout**

```python
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
```

- [ ] **Step 2: Commit**

```bash
git add shipping_helper/ui/main_window.py
git commit -m "feat: add PyQt5 main window with left-right layout"
```

---

### Task 8: 创建入口文件

**Files:**
- Create: `shipping_helper/main.py`

- [ ] **Step 1: Create main.py**

```python
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
```

- [ ] **Step 2: Commit**

```bash
git add shipping_helper/main.py
git commit -m "feat: add main entry point"
```

---

### Task 9: 测试验证

- [ ] **Step 1: 验证所有模块可导入**

```bash
cd shipping_helper
python -c "from core.order_parser import OrderParser; from core.pi_extractor import PIExtractor; from core.code_matcher import CodeMatcher; from core.package_calculator import PackageCalculator; print('All imports OK')"
```

- [ ] **Step 2: 启动GUI测试**

```bash
cd shipping_helper
python main.py
```

- [ ] **Step 3: Commit final**

```bash
git add -A
git commit -m "feat: complete Phase 1 basic implementation"
```

---

## 执行选项

**计划已保存至 `docs/superpowers/plans/2026-05-25-shipping-helper-phase1-plan.md`**

**两种执行方式：**

**1. Subagent-Driven（推荐）** - 每个任务分配一个subagent，任务间审查，快速迭代

**2. Inline Execution** - 在当前session执行任务，使用executing-plans，带检查点

**选择哪种方式？**