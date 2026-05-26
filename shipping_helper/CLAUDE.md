# ShippingHelper 开发指南

## 项目概述
船务部效率工具，用于处理下单生产阶段的订单合并、包装计算等工作。

## 开发流程
1. 每次开发前调用 brainstorming 技能进行需求确认
2. 根据任务性质调用对应技能（using-superpowers）
3. 如本地没有对应技能，使用 find-skills 查找

## Phase 1 核心任务
- 订单表粘贴解析（23个字段）
- PI文件提取（.xls格式）
- 数据汇聚（订单表 + PI + 商品编码）
- 包装计算（桶数、卡板、体积、毛重）

## 技术栈
- Python 3.x
- PyQt5（桌面客户端界面）
- xlrd（读取.xls文件）
- openpyxl（读取.xlsx文件）
- JSON（知识库）

## 关键注意事项
- PI文件是.xls格式，需用xlrd读取
- 订单要求字段含嵌入换行（液体包装+粉包装分开写）
- 报关名称可能为空，需从知识库补全

## 目录结构
```
shipping_helper/
├── core/           # 核心业务逻辑
├── ui/             # 用户界面
├── knowledge/      # 知识库JSON
├── config/         # 配置文件
└── main.py         # 入口文件
```

## brainstorming 技能使用指南

### 何时使用
- 每个新功能开发前必须调用 brainstorming 技能
- 需求不明确或需要澄清时

### 使用流程
1. 调用 brainstorming 技能
2. 探索项目上下文
3. 提出问题（一次一个）
4. 提出方案建议
5. 获得用户确认后，编写设计文档
6. 调用 writing-plans 技能编写实现计划

### 问题语言要求
- 向用户提问时使用**简体中文**
- 使用非技术语言，让用户容易理解
- 使用选择题形式（推荐）优于开放性问题

## 设计文档位置
- 设计规格：`docs/superpowers/specs/YYYY-MM-DD-<feature>-design.md`
- 实现计划：`docs/superpowers/plans/YYYY-MM-DD-<feature>-plan.md`

## 复用代码来源
`shipping_helper_old/` 目录包含可复用的模块：
- `core/order_parser.py` - 订单解析
- `core/pi_extractor.py` - PI文件提取
- `core/merger.py` - 数据汇聚
- `core/package_calculator.py` - 包装计算
- `core/code_matcher.py` - 商品编码匹配