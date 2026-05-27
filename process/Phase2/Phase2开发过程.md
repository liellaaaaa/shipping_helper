# Phase 2 开发过程文档

## 概述

本文档记录了 Phase 2「订舱出货」模块的开发过程，包括设计决策、技术挑战和实现细节。

---

## 一、数据模型设计 (data_models.py)

### 1.1 设计背景

Phase 1 使用简单的字典存储合并数据，但 Phase 2 需要处理多种不同来源的结构化数据：
- 运输鉴定报告 → CargoInfo
- MSDS文档 → MSDSInfo + Component + PhysicalChemical
- 出口商品编码表 → ExportCode
- 包装计算结果 → PackageInfo

### 1.2 数据类设计

| 类名 | 职责 | 关键字段 |
|------|------|---------|
| `Component` | MSDS成分项 | name, cas_no, concentration |
| `PhysicalChemical` | 理化特性 | appearance, ph, boiling_point, density, flash_point |
| `CargoInfo` | 货物信息 | product_name_cn/en, report_no, appearance_cn/en, transport_type, is_dangerous |
| `MSDSInfo` | MSDS完整信息 | product_name, components[], physical_chemical |
| `ExportCode` | 出口商品编码 | internal_code, hs_code, customs_name, composition |
| `PackageInfo` | 包装信息 | package_type, quantity_kg, drums, pallets, volume_cbm, gross_weight_kg |
| `ShipmentData` | 整合数据容器 | 继承Phase1字段 + cargo + msds + export_code + package_info |

### 1.3 关键设计决策

**决策1**: 使用 `@dataclass` 而非普通类
- 原因: 减少样板代码，自动生成 `__init__`, `__repr__` 等方法
- 收益: 代码简洁，易于维护

**决策2**: ShipmentData 作为数据整合容器
- 设计: `ShipmentData.from_phase1(pi_data)` 类方法从 Phase 1 数据创建
- 设计: `set_shipper()` 方法设置发货人
- 设计: `is_liquid()` 方法判断货物类型

---

## 二、出口商品编码加载器 (export_codes_loader.py)

### 2.1 功能需求

从 Excel 文件加载出口商品编码数据，供后续自动匹配使用。

### 2.2 文件格式

```
文件路径: 02.订舱出货/2024.12.5 最新出口商品编码及报关成分.xlsx
```

### 2.3 关键方法

| 方法 | 说明 |
|------|------|
| `load(filepath)` | 加载Excel文件 |
| `find_by_internal_code(code)` | 通过内部编号精确查找 |
| `find_by_name(name)` | 通过品名模糊搜索 |
| `is_loaded()` | 检查是否已加载 |
| `get_count()` | 获取数据条数 |

### 2.4 列名智能映射

由于 Excel 列名可能不统一，实现了智能映射：
- 产品内编: `["产品内编", "内编", "内部编号", "产品编号"]`
- 商品编码: `["新商品编码", "商品编码", "海关编码", "HS编码"]`
- 报关名称: `["报关名称", "品名", "名称"]`

---

## 三、运输鉴定报告解析器 (report_parser.py)

### 3.1 功能需求

解析运输鉴定报告（PDF和DOCX格式），提取货物信息。

### 3.2 支持格式

| 格式 | 解析方式 |
|------|---------|
| PDF | PyMuPDF (fitz) |
| DOCX | python-docx |

### 3.3 提取字段

- 样品名称（中文/英文）
- 鉴定书编号 (No.XXXXXXXX)
- 样品编号 (WXST2513478-036)
- 外观描述（中文/英文）
- 运输类型（海运/空运/陆运）
- 鉴定结论
- 委托单位
- 生产单位

### 3.4 技术实现

```python
# PDF解析核心逻辑
doc = fitz.open(filepath)
full_text = ""
for page in doc:
    full_text += page.get_text()
# 通过正则表达式提取字段
```

### 3.5 运输类型判断

根据文本内容判断：
- `imo imdg` / `imdg code` → 海运
- `iata dgr` / `空运` → 空运
- `jt/t` / `道路运输` → 陆运

---

## 四、MSDS解析器 (msds_parser.py)

### 4.1 功能需求

解析DOC格式的MSDS文档，提取第1/3/9部分的信息。

### 4.2 技术挑战

**挑战**: DOC格式（MSDS文件大多是老旧的.doc格式）不支持直接文本解析

**解决方案**: 使用 pywin32 Word COM 自动化

```python
import win32com.client
word_app = win32com.client.Dispatch("Word.Application")
document = word_app.Documents.Open(filepath)
```

### 4.3 第3部分成分表解析

**挑战**: 成分表存在不规则换行和合并单元格

**解决方案**:
1. 遍历文档中的表格
2. 查找包含"组分"/"CAS"/"含量"关键词的表格
3. 识别表头行确定列索引
4. 逐行解析数据，处理跨行合并

### 4.4 提取字段

| 部分 | 字段 |
|------|------|
| 第1部分 | 产品名称 |
| 第3部分 | 成分表（组分/CAS号/含量） |
| 第9部分 | 外观、pH、熔点、沸点、密度、闪点、溶解性 |

---

## 五、数据整合器 (data_merger.py)

### 5.1 功能需求

整合 Phase 1 数据和 Phase 2 解析结果，形成完整的 ShipmentData。

### 5.2 核心方法

| 方法 | 说明 |
|------|------|
| `load_export_codes(filepath)` | 加载出口商品编码 |
| `create_shipment_from_phase1(pi_data, order_data)` | 从Phase1创建ShipmentData |
| `assign_cargo(shipment, cargo_data)` | 分配货物信息 |
| `assign_msds(shipment, msds_data)` | 分配MSDS信息 |
| `assign_export_code(shipment, internal_code)` | 匹配出口商品编码 |
| `assign_package_info(shipment, package_data)` | 分配包装信息 |
| `merge_all(...)` | 一键整合所有数据源 |

### 5.3 自动匹配逻辑

```python
# 根据内部编号自动匹配出口商品编码
export_code = export_codes_loader.find_by_internal_code(internal_code)
if export_code:
    shipment.export_code = export_code
    # 同步海关编码
    if export_code.hs_code:
        shipment.hs_code = export_code.hs_code
```

---

## 六、LOI保函生成器 (loi_generator.py)

### 6.1 功能需求

生成非危险品保函和液体保函，基于模板填充。

### 6.2 模板文件

| 模板 | 文件路径 |
|------|---------|
| 非危险品保函 | `LOI-op-非危险品保函模板.docx` |
| 液体保函 | `LOI-op-液体保函模板.docx` |

### 6.3 模板选择逻辑

```python
def select_template(self, shipment):
    if shipment.is_liquid():
        return "液体保函模板.docx"
    else:
        return "非危险品保函模板.docx"
```

### 6.4 占位符替换

使用 `{{key}}` 格式的占位符，通过 python-docx 替换：

```python
for para in doc.paragraphs:
    text = para.text
    for match in PLACEHOLDER_PATTERN.finditer(text):
        key = match.group(1)
        value = data.get(key, '')
        text = text.replace(match.group(0), value)
```

---

## 七、MSDS生成器 (msds_generator.py)

### 7.1 功能需求

生成中文版和英文版MSDS文档。

### 7.2 数据来源优先级

| 优先级 | 来源 | 字段 |
|--------|------|------|
| 1 | MSDSInfo | 产品名称、成分、理化特性 |
| 2 | ExportCode | 报关成分 |
| 3 | CargoInfo | 外观描述 |

### 7.3 生成内容

| 版本 | 内容 |
|------|------|
| 中文MSDS | 第1/3/9部分完整中文格式 |
| 英文MSDS | 翻译后内容，保持英文格式 |

### 7.4 英文翻译

实现了简单的中文到英文翻译映射：
```python
translations = {
    "液体": "liquid",
    "固体": "solid",
    "粉末": "powder",
    "白色": "white",
    "黄色": "yellow",
    "透明": "clear",
    # ...
}
```

---

## 八、订舱单生成器 (booking_generator.py)

### 8.1 功能需求

生成完整的订舱单文档，包含所有收发货信息和货物信息。

### 8.2 文档结构

```
1. 标题: 订舱单
2. 副标题: PI号
3. 发货人信息 (Shipper)
4. 收货人信息 (Consignee)
5. 通知人信息 (Notify Party)
6. 货物信息 (Cargo Details)
7. 包装信息 (Packing Details)
8. 目的地 (Destination)
9. 理化特性 (如有MSDS)
```

### 8.3 关键实现

使用 python-docx 创建文档，设置页面边距，添加标题和字段：

```python
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

doc = Document()
title = doc.add_heading("订舱单", 0)
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
```

---

## 九、UI界面 (booking_widget.py)

### 9.1 布局设计

采用左右分栏布局：
- **左侧**: 数据看板（只读展示）
- **右侧**: 模板编辑区（Tab切换）

### 9.2 左侧数据看板

| Group | 内容 |
|-------|------|
| 继承自Phase 1 | 发货人/收货人/通知人/PI号等 |
| 运输鉴定报告 | 报告编号/运输类型/外观/结论 |
| MSDS信息 | 产品名称/成分表/理化特性 |
| 出口商品编码 | 海关编码/报关名称/成分 |
| 包装信息 | 包装类型/数量/桶数/体积/毛重 |

### 9.3 右侧模板编辑区

Tab选项卡：
1. **订舱单**: 订舱单预览
2. **MSDS**: 中英文MSDS生成
3. **保函**: LOI模板选择

### 9.4 Word嵌入方案

使用 `QAxWidget` 嵌入 Word ActiveX，实现 WYSIWYG 编辑：
- 用户可以直接在嵌入的 Word 中编辑
- 原生 Word 功能保证格式保留
- 无需自行实现富文本编辑器

### 9.5 Phase 1 集成

在 `main_window.py` 中，点击"进入 Phase 2"按钮时：
1. 提取 Phase 1 的合并数据
2. 提取包装计算结果
3. 创建 BookingWidget 并传递数据

---

## 十、遇到的问题及解决

### 问题1: 相对导入错误

**错误**: `ImportError: attempted relative import beyond top-level package`

**原因**: `booking_widget.py` 使用 `from ..core.phase2 import ...` 相对导入，但该文件在 `ui/` 目录下，无法使用相对导入访问 `core/`。

**解决**: 使用 `sys.path.insert(0, ...)` 动态添加路径，然后使用绝对导入。

### 问题2: __init__.py 内容错误

**问题**: `phase2/__init__.py` 包含了数据模型的代码（应该是空模块或只做导入）

**原因**: 最初开发时将数据模型代码直接写在 `__init__.py` 中

**解决**: 创建独立的 `data_models.py`，`__init__.py` 只做导入导出

### 问题3: python-docx 缺失

**错误**: `ModuleNotFoundError: No module named 'docx'`

**解决**: `pip install python-docx`

### 问题4: Pt 未定义

**错误**: `NameError: name 'Pt' is not defined`

**原因**: `Pt` 在方法内部导入，但应该放在文件顶部

**解决**: 将导入语句移到文件顶部

---

## 十一、文件结构

```
shipping_helper/
├── core/
│   ├── phase2/
│   │   ├── __init__.py          # 模块导出
│   │   ├── data_models.py       # 数据类定义
│   │   ├── export_codes_loader.py  # 出口商品编码加载
│   │   ├── report_parser.py     # 运输鉴定报告解析
│   │   ├── msds_parser.py       # MSDS解析
│   │   ├── data_merger.py       # 数据整合
│   │   ├── booking_generator.py  # 订舱单生成
│   │   ├── msds_generator.py    # MSDS生成
│   │   └── loi_generator.py      # 保函生成
│   └── ...
├── ui/
│   ├── booking_widget.py         # Phase 2 UI
│   └── main_window.py            # Phase 1 UI
└── process/                      # 开发过程文档
    ├── Phase2测试报告.md
    └── Phase2开发过程.md
```