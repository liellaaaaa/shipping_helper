# Phase 1 设计规格

## 概述

Phase 1 是 ShippingHelper 船务助手的第一阶段，专注于**订单数据提取与包装计算**。目标是解决外贸订单处理中的数据重复录入和包装计算繁琐的问题。

---

## 一、核心功能

### 1.1 订单表粘贴解析

**功能描述**: 用户从在线表格复制一行订单数据，粘贴到系统，系统自动解析23个字段。

**解析字段列表**:

| 序号 | 字段名 | 说明 | 示例 |
|------|--------|------|------|
| 1 | 业务员 | 负责该订单的业务员 | 王小明 |
| 2 | 客户编号 | 客户唯一标识 | 客户001 |
| 3 | 内部编号 | 产品内部编码 | 1011 |
| 4 | 产品中文名 | 产品中文名称 | 固色剂 |
| 5 | 报关名称 | 出口报关用名称 | 固色剂 |
| 6 | 规格kg | 每桶/包重量 | 25kg/桶 |
| 7 | 订单量kg | 总订单数量 | 1000 |
| 8 | 是否调价 | 价格调整标识 | 否 |
| 9 | 有无样品 | 样品标识 | 无 |
| 10 | 订单要求 | 特殊要求说明 | 液体包装+纸桶 |
| 11 | 交货日期 | 预计交货日期 | 2024-12-15 |
| 12 | 审核 | 审核状态 | 已审核 |
| 13 | 销售区域 | 销售地区 | 华南 |
| 14 | 订单号 | 内部订单号 | HT2024001 |
| 15 | 出货抬头 | 出货显示名称 | 宏昊 |
| 16 | 单据类型 | 文档类型 | PI |
| 17 | 跟单员 | 跟单负责人 | 李娜 |
| 18 | 下单日期 | 下单时间 | 2024-12-01 |
| 19 | 确认下单 | 确认状态 | 确认 |
| 20 | 生产交期 | 生产完成日期 | 2024-12-20 |
| 21 | 出货渠道 | 出货方式 | 出货 |
| 22 | 出货方式 | 运输方式 | 海运 |
| 23 | 规格异常 | 特殊规格说明 | 粉包装异常 |

**输入格式**: Tab分隔或换行分隔的文本

**技术实现**: OrderParser 类，自动识别分隔符

### 1.2 PI文件提取

**功能描述**: 读取PI合同Excel文件，提取收发货人信息。

**提取字段**:

| 字段名 | 说明 | 示例 |
|--------|------|------|
| 收货人 | 买方公司名称 | TOA-DOVECHEM INDUSTRIES CO., LTD |
| 收货人地址 | 买方地址 | 泰国曼谷 |
| 日期 | PI日期 | 2024-12-01 |
| PI号 | 合同编号 | HT2024001 |
| 品名英文 | 产品英文名称 | Fixing Agent |
| 数量 | 订单数量 | 1000 |
| 单价 | CIF单价 | USD 2.50/KG |
| 金额 | 总金额 | USD 2500 |
| H.S.Code | 海关编码 | 3808930000 |
| 卸货港 | 目的港 | BANGKOK PORT |
| 包装说明 | 包装要求 | 纸箱包装 |

**文件格式**: .xls (xlrd库读取)

**技术实现**: PIExtractor 类

### 1.3 数据汇聚

**功能描述**: 将订单数据、PI数据、商品编码知识库三合一。

**合并逻辑**:

```
merged_data = order_data ∪ pi_data ∪ code_info
```

**优先级规则**:
- 报关名称: 订单数据 > 商品编码知识库
- H.S.Code: PI数据 > 商品编码知识库

**技术实现**: Merger 类

### 1.4 包装计算

**功能描述**: 根据订单数量和包装类型，计算所需桶数、卡板数，体积，毛重。

**支持包装类型**:

| 包装类型 | 说明 |
|----------|------|
| 30kg蓝桶 | 蓝色塑料桶，30kg/桶 |
| 25kg/包 | 纸袋包装，25kg/包 |
| 25kg 阀口袋(灰色粒) | 阀口袋包装 |
| 25kg 纸桶 | 纸桶包装 |
| 50kg蓝桶(细口) | 50kg细口塑料桶 |
| 50kg蓝桶(大口) | 50kg大口塑料桶 |
| 125kg新款胶桶 | 大容量胶桶 |
| 200kg双环闭口桶 | 200L大桶 |
| 1吨桶 | IBC吨桶 |

**卡板规格**:

| 卡板规格 | 说明 |
|----------|------|
| 1.0*1.0m | 标准卡板 |
| 1.1*1.1m | 大尺寸卡板 |

**计算输出**:

| 字段 | 说明 |
|------|------|
| 桶数 | 所需桶数 |
| 卡板数 | 所需卡板数 |
| 体积(CBM) | 总体积 |
| 毛重(kg) | 总毛重 |
| 20GP判定 | 是否适合20GP集装箱 |

**技术实现**: PackageCalculator 类

---

## 二、数据模型

### 2.1 知识库

**产品知识库** (products_knowledge.json):

```json
{
  "version": "2024.12",
  "products": {
    "1011": {
      "报关名称": "固色剂",
      "海关编码": "3808930000",
      "报关成分": "具体成分..."
    }
  }
}
```

**包装数据** (packaging_data.json):

```json
{
  "version": "2024.12",
  "packages": [...],
  "pallets": [...],
  "pallet_capacity": {
    "30kg蓝桶": {
      "1.0*1.0": null,
      "1.1*1.1": 24
    }
  }
}
```

---

## 三、UI设计

### 3.1 整体布局

采用**左右分栏布局**:

```
┌─────────────────────────────────────────────────────────────────┐
│  ShippingHelper - Phase 1: 订单数据提取与包装计算              │
├──────────────────────────┬──────────────────────────────────────┤
│                          │                                      │
│      左侧面板 (40%)       │           右侧面板 (60%)             │
│                          │                                      │
│  ┌──────────────────┐    │  ┌────────────────────────────────┐│
│  │ 数据输入         │    │  │ 订单要求                        ││
│  │ ├─ 订单表粘贴    │    │  │ (QTextEdit, 可编辑)             ││
│  │ └─ PI文件选择    │    │  └────────────────────────────────┘│
│  │ [开始解析]       │    │                                      │
│  └──────────────────┘    │  ┌────────────────────────────────┐│
│                          │  │ 包装计算                        ││
│  ┌──────────────────┐    │  │ ├─ 包装类型选择                 ││
│  │ 合并字段结果     │    │  │ ├─ 数量(kg)输入                ││
│  │ ├─ 外贸订单表    │    │  │ ├─ 卡板选择                    ││
│  │ │   (23字段表格) │    │  │ └─ [计算] [清除]              ││
│  │ └─ PI文件       │    │  │                                ││
│  │   (12字段表格)  │    │  │ 包装表格 (可多行)              ││
│  │                  │    │  │ 序号|包装|数量|桶数|卡板|...  ││
│  │ [上移] [下移]   │    │  │ ─────────────────────────────  ││
│  └──────────────────┘    │  │ [合计行]                       ││
│                          │  └────────────────────────────────┘│
├──────────────────────────┴──────────────────────────────────────┤
│                    [进入 Phase 2]                               │
└─────────────────────────────────────────────────────────────────┘
```

### 3.2 组件说明

| 组件 | 类型 | 说明 |
|------|------|------|
| order_text_edit | QTextEdit | 订单表粘贴区域 |
| pi_path_edit | QLineEdit | PI文件路径显示 |
| btn_select_pi | QPushButton | 选择PI文件 |
| btn_calculate | QPushButton | 执行解析计算 |
| order_fields_table | QTableWidget | 订单字段结果(23行) |
| pi_fields_table | QTableWidget | PI字段结果(12行) |
| order_req_edit | QTextEdit | 订单要求(可编辑) |
| pkg_type_combo | QComboBox | 包装类型选择 |
| pallet_combo | QComboBox | 卡板规格选择 |
| per_pallet_edit | QLineEdit | 每板数量(可选) |
| package_table | QTableWidget | 包装计算结果表格 |
| totals_widget | QWidget | 底部固定合计行 |
| btn_phase2 | QPushButton | 进入Phase 2 |

---

## 四、关键类设计

### 4.1 OrderParser

```python
class OrderParser:
    COLUMNS = 23  # 固定23个字段

    def parse(self, text: str) -> dict:
        """解析订单文本，返回字段字典"""

    def validate(self) -> tuple[bool, str]:
        """验证数据有效性"""

    def get_internal_code(self) -> str:
        """提取内部编号"""

    def get_pi号(self) -> str:
        """提取PI号"""
```

### 4.2 PIExtractor

```python
class PIExtractor:
    def parse_file(self, filepath: str) -> dict:
        """解析PI Excel文件"""

    def _extract_consignee(self, ws) -> str:
        """提取收货人"""

    def _extract_pi_no(self, ws) -> str:
        """提取PI号"""

    def _extract_hs_code(self, ws) -> str:
        """提取H.S.Code"""
```

### 4.3 CodeMatcher

```python
class CodeMatcher:
    def load(self, filepath: str) -> bool:
        """加载知识库"""

    def get_all_info(self, internal_code: str) -> dict:
        """获取产品完整信息"""

    def match(self, name: str) -> list:
        """模糊匹配产品名称"""
```

### 4.4 PackageCalculator

```python
class PackageCalculator:
    def calculate_with_pallet(
        self, pkg_type: str, pallet_type: str,
        qty: float, drums_per_pallet: int = None,
        pallet_version: str = 'new'
    ) -> dict:
        """计算带卡板的包装"""

    def calculate_no_pallet(self, pkg_type: str, qty: float) -> dict:
        """计算不带卡板的包装"""

    def parse_order_requirements(self, req: str) -> dict:
        """解析订单要求中的包装信息"""
```

---

## 五、数据流

### 5.1 完整数据流

```
1. 用户粘贴订单数据
   ↓
2. OrderParser.parse() 解析23字段
   ↓
3. OrderParser.validate() 验证
   ↓
4. 用户选择PI文件
   ↓
5. PIExtractor.parse_file() 提取PI数据
   ↓
6. CodeMatcher.get_all_info() 查询商品编码
   ↓
7. Merger.merge() 合并数据
   ↓
8. PackageCalculator 计算包装
   ↓
9. UI更新显示结果
   ↓
10. 用户可手动调整
   ↓
11. 点击"进入Phase 2"
```

### 5.2 数据优先级

| 字段 | 来源 | 优先级 |
|------|------|--------|
| 发货人 | 固定值(宏昊/民浩) | - |
| 收货人 | PI文件 | 最高 |
| 报关名称 | 订单表 > 知识库 | 订单优先 |
| H.S.Code | PI文件 > 知识库 | PI优先 |
| 包装信息 | 用户选择 | 实时计算 |

---

## 六、技术选型

| 组件 | 技术 | 原因 |
|------|------|------|
| UI框架 | PyQt5 | 成熟的桌面GUI框架 |
| 数据存储 | JSON | 知识库简单易用 |
| Excel读取(.xls) | xlrd | Phase 1 PI文件是.xls格式 |
| Excel读取(.xlsx) | openpyxl | 新版Excel格式 |
| 文档生成 | python-docx | Word文档生成 |

---

## 七、目录结构

```
shipping_helper/
├── main.py                    # 入口文件
├── requirements.txt           # 依赖列表
├── CLAUDE.md                 # 开发指南
│
├── core/                     # 核心业务逻辑
│   ├── order_parser.py       # 订单解析器
│   ├── pi_extractor.py       # PI提取器
│   ├── code_matcher.py       # 编码匹配器
│   ├── package_calculator.py # 包装计算器
│   └── merger.py             # 数据汇聚器
│
├── ui/
│   └── main_window.py        # 主窗口UI
│
├── knowledge/                # 知识库
│   ├── products_knowledge.json  # 产品知识库
│   └── packaging_data.json      # 包装数据
│
└── tests/
    └── core/
        └── test_phase1.py    # Phase 1测试
```