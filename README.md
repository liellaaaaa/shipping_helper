# ShippingHelper - 宏昊船务助手

船务部效率工具，用于处理下单生产阶段的订单合并、包装计算等工作。

## 快速启动

### 1. 创建虚拟环境

```bash
cd c:/Users/windows/Desktop/shipping_helper
.venv\Scripts\python.exe -m venv .venv
```

### 2. 安装依赖

```bash
.venv\Scripts\pip.exe install -r shipping_helper/requirements.txt
```

### 3. 运行程序

```bash
.venv\Scripts\python.exe shipping_helper/main.py
```

## Phase 1 功能 ✅ 已完成

- [x] 订单表粘贴解析（23个字段）
- [x] PI文件提取（.xls格式）
- [x] 数据汇聚（订单表 + PI + 商品编码）
- [x] 包装计算（桶数、卡板、体积、毛重、毛重明细）
- [x] 表格字段上移/下移调整
- [x] 行高自适应换行
- [x] 卡板组功能（同一卡板放置多个产品）
- [x] 卡板版本选择（1.0*1.0m / 1.1*1.1m）
- [x] 手动调整卡板数量
- [x] 合并结果表格展示（支持点击复制）

## 目录结构

```
shipping_helper/
├── core/                      # 核心业务逻辑
│   ├── order_parser.py        # 订单解析器
│   ├── pi_extractor.py        # PI提取器
│   ├── code_matcher.py        # 商品编码匹配器
│   ├── package_calculator.py  # 包装计算器
│   └── merger.py              # 数据汇聚器
├── ui/                        # 用户界面
│   └── main_window.py         # 主窗口
├── knowledge/                 # 知识库
│   ├── products_knowledge.json
│   └── packaging_data.json
├── config/                    # 配置文件
│   └── settings.json
├── docs/                      # 设计文档
│   └── superpowers/
│       ├── specs/
│       └── plans/
├── tests/                     # 测试
│   └── core/
├── main.py                    # 入口
└── requirements.txt           # 依赖
```

## 技术栈

- Python 3.x
- PyQt5（桌面客户端界面）
- xlrd（读取.xls文件）
- openpyxl（读取.xlsx文件）
- JSON（知识库）

## 后续计划

- Phase 2: 模板填充（LOI、MSDS等）
- Phase 3: 订舱出货

## 注意事项

- 公司内部文件（01.下单生产、02.订舱出货）不纳入版本管理
- 虚拟环境依赖安装在 `.venv` 目录下