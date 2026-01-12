# CLAUDE.md

此文件为 Claude Code (claude.ai/code) 在本仓库中工作时提供指导。

## 项目概述

会计助手是一个 Python 桌面应用程序，用于自动化制造企业的会计操作。通过 Tkinter GUI 提供两个主要功能：

1. **生成凭证**：从发票 Excel 文件创建会计凭证
2. **调整税负率**：计算并调整成本系数以达到目标税负率

## 常用命令

```bash
# 安装依赖
pip install -r requirements.txt

# 运行应用 (Windows)
pythonw accounting_assistant.pyw

# 运行应用 (跨平台)
python accounting_assistant.pyw

# 运行测试
pytest __tests__/

# 运行特定模块测试
pytest __tests__/voucher/
pytest __tests__/tax_adjuster/
```

## 架构

### 项目结构

```
accounting-assistant/
├── accounting_assistant.pyw   # 主入口文件
├── requirements.txt           # Python 依赖
├── CLAUDE.md                  # 项目文档
├── data/                      # 示例/测试数据文件
│   └── *.xlsx                 # Excel 数据文件
│
├── modules/                   # 核心业务模块
│   ├── __init__.py
│   ├── voucher/               # 凭证生成模块
│   │   ├── __init__.py
│   │   ├── voucher_tab.py         # UI 标签页，支持拖放选择文件
│   │   ├── handle_outbound_data.py # 解析出库发票 Excel 文件
│   │   ├── handle_inbound_data.py  # 解析入库发票 Excel 文件
│   │   ├── create_outbound.py      # 生成出库凭证工作表
│   │   ├── create_inbound.py       # 生成入库凭证工作表
│   │   ├── create_issuing.py       # 生成领料单
│   │   ├── create_receiving.py     # 生成收料单
│   │   ├── helpers.py              # 工具函数（边框、随机数生成）
│   │   └── config.py               # 单位类型和无效产品过滤配置
│   │
│   └── tax_adjuster/          # 税负率调整模块
│       ├── __init__.py
│       ├── tax_tab.py              # 税负调整 UI 标签页
│       └── adjust_tax.py           # TaxAdjuster 类，包含计算引擎
│
└── __tests__/                 # 测试目录
    ├── __init__.py
    ├── conftest.py                 # 全局 pytest fixtures
    ├── voucher/                    # 凭证模块测试
    │   ├── conftest.py             # 凭证模块 fixtures
    │   ├── test_handle_outbound_data.py
    │   ├── test_handle_inbound_data.py
    │   ├── test_create_outbound.py
    │   └── test_helpers.py
    │
    └── tax_adjuster/               # 税负调整模块测试
        ├── conftest.py             # 税负模块 fixtures
        └── test_adjust_tax.py
```

### 关键数据流

**凭证生成流程：**
- 出库发票 → `handle_outbound_data` → `create_outbound` → 输出工作表
- 测算表 → `create_inbound` + `create_issuing` → 输出工作表
- 入库发票 → `handle_inbound_data` → `create_receiving` → 输出工作表

**税负调整流程：**
- 用户输入目标税负率 → `TaxAdjuster.calculate_adjustment()` 使用二分查找寻找最优 G25（成本系数）→ 应用修改到 Excel

### 重要 Excel 单元格引用（测算表）

- **E17**: 年收入
- **E18**: 年利润总额
- **E21**: 年应纳税额
- **G25**: 成本系数 - 主要调整参数
- **B46**: 当月利润

### 数据处理模式

凭证模块在 `create_outbound.py` 中使用多阶段处理管道：
1. 按公司合并
2. 按日期拆分
3. 合并数量
4. 每7行分组
5. 按日期排序

### 配置说明 (config.py)

- `INT_UNITS`: 需要整数数量的单位（如 '个'）
- `FLOAT_UNITS`: 允许小数数量的单位（如 '吨', 'kg'）
- `INVALID_PRODUCT_TYPES`: 需要排除的产品类型（'机动车', '劳务'）

## 技术栈

- **GUI**: Tkinter + tkinterdnd2（拖放支持）
- **Excel**: openpyxl 读写 .xlsx 文件
- **平台**: 主要针对 Windows（.pyw 扩展名），支持跨平台运行
