# 库存毛利率调整优化设计

## 概述

优化 `modules/tax_adjuster/adjust_tax.py` 中的库存毛利率调整功能，提升算法性能并支持用户自定义参数。

## 问题分析

### 当前问题

1. **性能问题**：现有算法使用三阶段搜索（粗略扫描 121 点 + 精细搜索 ~100 点 + 梯度下降 ~20 次），总计约 240 次 Excel 公式计算，速度较慢
2. **精度问题**：网格搜索容易错过最优解
3. **灵活性不足**：H11、F20 目标值和毛利率范围硬编码，用户无法自定义

### 优化目标

- 算法性能提升 5-8 倍（从 ~240 次计算降至 ~40 次）
- 支持用户自定义 H11、F20、毛利率的范围约束

## 设计方案

### 1. 算法优化

**核心思路：利用 F20-B11 线性关系 + 二分法搜索 margin**

#### 原理

- 对于任意固定的 margin 值，F20 与 B11 是完美线性关系
- 只需 2 次计算即可确定直线方程，直接求解使 F20 落在目标范围内的 B11
- 对 margin 使用二分法搜索，找到使 H11 落在目标范围内的值

#### 新算法伪代码

```python
def find_optimal_margin_v2(h11_range, f20_range, margin_range):
    margin_min, margin_max = margin_range
    h11_min, h11_max = h11_range
    f20_min, f20_max = f20_range

    for margin in binary_search(margin_min, margin_max):
        # 利用线性关系直接求解 B11
        # 采样两点确定 F20 = k * B11 + b
        f20_0 = calculate(margin, B11=0)
        f20_1 = calculate(margin, B11=100000)
        k = (f20_1 - f20_0) / 100000
        b = f20_0

        # 求解使 F20 在目标范围中点的 B11
        target_f20 = (f20_min + f20_max) / 2
        B11 = (target_f20 - b) / k

        # 验证结果
        H11, F20 = calculate(margin, B11)

        if h11_min <= H11 <= h11_max and f20_min <= F20 <= f20_max:
            return (margin, B11, H11, F20)  # 找到满足约束的解

    return best_approximate_solution
```

#### 预期性能

- 二分法迭代次数：约 15-20 次
- 每次迭代计算：2 次（线性采样）+ 1 次（验证）= 3 次
- 总计算量：约 45-60 次，相比原来 240 次提升约 5 倍

### 2. 参数对话框 UI

#### 对话框布局

```
┌─────────────────────────────────────────────┐
│        调整库存毛利率 - 参数设置             │
├─────────────────────────────────────────────┤
│                                             │
│  H11 范围      [  -10  ] ~ [   10  ]        │
│                                             │
│  F20 范围      [-40000 ] ~ [ 40000 ]        │
│                                             │
│  毛利率范围     [  0.70 ] ~ [  0.90 ]        │
│                                             │
│            [取消]      [开始计算]            │
└─────────────────────────────────────────────┘
```

#### 默认值

| 参数 | 最小值 | 最大值 |
|------|--------|--------|
| H11 | -10 | 10 |
| F20 | -40000 | 40000 |
| 毛利率 | 0.70 | 0.90 |

#### 交互流程

1. 用户点击"调整库存毛利率"按钮
2. 弹出参数设置对话框，显示默认值
3. 用户修改参数（可选）
4. 点击"开始计算"关闭对话框，开始计算并显示进度条
5. 计算完成后显示结果

### 3. 代码结构变更

#### 修改文件

| 文件 | 变更内容 |
|------|----------|
| `adjust_tax.py` | 新增 `find_optimal_margin_v2()` 方法；修改 `calculate_inventory_margin_adjustment()` 接受范围参数 |
| `tax_tab.py` | 新增 `MarginParamsDialog` 对话框类；修改 `adjust_inventory_margin()` 方法 |

#### adjust_tax.py 变更

```python
class TaxAdjuster:
    # 新增方法
    def find_optimal_margin_v2(self, h11_range, f20_range, margin_range):
        """
        优化版搜索算法：利用 F20-B11 线性关系 + 二分法

        Args:
            h11_range: (min, max) H11 目标范围
            f20_range: (min, max) F20 目标范围
            margin_range: (min, max) 毛利率搜索范围

        Returns:
            dict: 包含 margin, B11, H11, F20, converged
        """
        pass

    # 修改现有方法签名
    def calculate_inventory_margin_adjustment(
        self,
        h11_range=(-10, 10),
        f20_range=(-40000, 40000),
        margin_range=(0.70, 0.90),
        max_solutions=5
    ):
        pass
```

#### tax_tab.py 变更

```python
class MarginParamsDialog(wx.Dialog):
    """库存毛利率参数设置对话框"""

    def __init__(self, parent):
        super().__init__(parent, title="调整库存毛利率 - 参数设置")
        # H11、F20、毛利率范围输入框
        # 取消、开始计算按钮

    def get_params(self):
        """返回用户输入的参数"""
        return {
            'h11_range': (h11_min, h11_max),
            'f20_range': (f20_min, f20_max),
            'margin_range': (margin_min, margin_max),
        }
```

## 实现步骤

1. 在 `tax_tab.py` 新增 `MarginParamsDialog` 对话框类
2. 修改 `adjust_inventory_margin()` 弹出对话框获取参数
3. 在 `adjust_tax.py` 新增 `find_optimal_margin_v2()` 方法
4. 修改 `calculate_inventory_margin_adjustment()` 接受范围参数并调用新算法
5. 测试验证

## 兼容性

- 默认参数值与当前行为一致
- 现有的 `find_optimal_margin_fast()` 保留作为备用
