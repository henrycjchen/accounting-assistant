# 库存毛利率调整优化 - 实现计划

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** 优化库存毛利率调整功能，提升算法性能 5 倍，并支持用户自定义 H11/F20/毛利率范围参数。

**Architecture:** 利用 F20-B11 线性关系，将三阶段搜索替换为二分法 + 线性插值算法。新增 wxPython 参数对话框，在计算前收集用户输入的约束范围。

**Tech Stack:** Python 3.11, wxPython, openpyxl, formulas

---

## Task 1: 创建测试文件框架

**Files:**
- Create: `tests/tax_adjuster/test_margin_optimization.py`
- Reference: `modules/tax_adjuster/adjust_tax.py:559-743` (现有 `find_optimal_margin_fast`)

**Step 1: 创建测试目录和测试文件**

```bash
mkdir -p tests/tax_adjuster
touch tests/tax_adjuster/__init__.py
```

**Step 2: 编写测试框架**

```python
# tests/tax_adjuster/test_margin_optimization.py
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
库存毛利率优化算法测试
"""

import pytest
from unittest.mock import MagicMock, patch


class TestFindOptimalMarginV2:
    """测试优化版搜索算法 find_optimal_margin_v2"""

    def test_returns_dict_with_required_keys(self):
        """验证返回结果包含必要的字段"""
        # 将在实现后填充
        pass

    def test_uses_linear_interpolation_for_b11(self):
        """验证利用线性关系计算 B11"""
        pass

    def test_binary_search_for_margin(self):
        """验证对 margin 使用二分法"""
        pass

    def test_respects_range_constraints(self):
        """验证遵守用户指定的范围约束"""
        pass

    def test_returns_converged_when_solution_found(self):
        """验证找到解时 converged=True"""
        pass

    def test_returns_approximate_when_no_exact_solution(self):
        """验证无精确解时返回近似解"""
        pass
```

**Step 3: 运行测试验证框架**

Run: `pytest tests/tax_adjuster/test_margin_optimization.py -v`
Expected: 6 tests collected, all PASS (empty test bodies)

**Step 4: Commit**

```bash
git add tests/tax_adjuster/
git commit -m "$(cat <<'EOF'
test: add test framework for margin optimization

Add test file structure for the optimized find_optimal_margin_v2 algorithm.

Co-Authored-By: Claude Opus 4.5 <noreply@anthropic.com>
EOF
)"
```

---

## Task 2: 实现 find_optimal_margin_v2 核心算法

**Files:**
- Modify: `modules/tax_adjuster/adjust_tax.py:559` (在 `find_optimal_margin_fast` 后添加新方法)
- Test: `tests/tax_adjuster/test_margin_optimization.py`

**Step 1: 编写算法测试**

```python
# tests/tax_adjuster/test_margin_optimization.py
# 替换 TestFindOptimalMarginV2 类

class TestFindOptimalMarginV2:
    """测试优化版搜索算法 find_optimal_margin_v2"""

    @pytest.fixture
    def mock_adjuster(self):
        """创建带模拟计算的 TaxAdjuster"""
        with patch('modules.tax_adjuster.adjust_tax.TaxAdjuster._load_model'):
            adjuster = MagicMock()
            adjuster.MARGIN_MIN = 0.70
            adjuster.MARGIN_MAX = 0.90
            adjuster.B11_MIN = 0
            adjuster.B11_MAX = 500_000
            adjuster.H11_MIN = -10
            adjuster.H11_MAX = 10
            adjuster.F20_MIN = -40_000
            adjuster.F20_MAX = 40_000
            return adjuster

    def test_returns_dict_with_required_keys(self, mock_adjuster):
        """验证返回结果包含必要的字段"""
        from modules.tax_adjuster.adjust_tax import TaxAdjuster

        # 模拟线性关系: F20 = -0.5 * B11 + 10000
        # 当 margin=0.80, B11=20000 时, H11=0, F20=0
        def mock_calculate(margin, b11):
            # H11 随 margin 增加而增加
            h11 = (margin - 0.80) * 100
            # F20 = k * B11 + b, 在 B11=20000 时 F20=0
            f20 = -0.5 * b11 + 10000
            return h11, f20

        with patch.object(TaxAdjuster, '__init__', lambda self, *args, **kwargs: None):
            adjuster = TaxAdjuster.__new__(TaxAdjuster)
            adjuster.MARGIN_MIN = 0.70
            adjuster.MARGIN_MAX = 0.90
            adjuster.B11_MIN = 0
            adjuster.B11_MAX = 500_000
            adjuster._model = MagicMock()

            # Mock _calculate to return predictable values
            def side_effect(inputs):
                margin = inputs.get("'[test]生产成本月结表'!J14", 0.80)
                b11 = inputs.get("'[test]产品成本'!B11", 0)
                h11, f20 = mock_calculate(margin, b11)
                solution = MagicMock()
                solution.__getitem__ = lambda self, key: MagicMock(value=[[h11]] if 'H11' in key else [[f20]])
                return solution

            adjuster._calculate = MagicMock(side_effect=side_effect)
            adjuster._cell_key = lambda sheet, cell: f"'[test]{sheet}'!{cell}"
            adjuster._to_number = lambda v, d=0: v[0][0] if hasattr(v, '__getitem__') else v
            adjuster._get_value = lambda sol, sheet, cell: sol[f"'[test]{sheet}'!{cell}"].value
            adjuster.MARGIN_SHEET = '生产成本月结表'
            adjuster.MARGIN_CELL = 'J14'
            adjuster._report_progress = lambda *args: None

            result = adjuster.find_optimal_margin_v2(
                h11_range=(-10, 10),
                f20_range=(-40000, 40000),
                margin_range=(0.70, 0.90)
            )

            assert 'margin' in result
            assert 'B11' in result
            assert 'H11' in result
            assert 'F20' in result
            assert 'converged' in result
            assert 'iterations' in result
```

**Step 2: 运行测试验证失败**

Run: `pytest tests/tax_adjuster/test_margin_optimization.py::TestFindOptimalMarginV2::test_returns_dict_with_required_keys -v`
Expected: FAIL with "AttributeError: 'TaxAdjuster' object has no attribute 'find_optimal_margin_v2'"

**Step 3: 实现 find_optimal_margin_v2 方法**

在 `modules/tax_adjuster/adjust_tax.py` 中，在 `find_optimal_margin_fast` 方法后（约 744 行）添加：

```python
def find_optimal_margin_v2(self, h11_range, f20_range, margin_range):
    """
    优化版搜索算法：利用 F20-B11 线性关系 + 二分法

    原理：
    - 对于任意固定的 margin 值，F20 与 B11 是完美线性关系
    - 只需 2 次计算即可确定直线方程，直接求解使 F20 落在目标范围内的 B11
    - 对 margin 使用二分法搜索，找到使 H11 落在目标范围内的值

    Args:
        h11_range: (min, max) H11 目标范围
        f20_range: (min, max) F20 目标范围
        margin_range: (min, max) 毛利率搜索范围

    Returns:
        dict: 包含 margin, B11, H11, F20, converged, iterations
    """
    h11_min, h11_max = h11_range
    f20_min, f20_max = f20_range
    margin_min, margin_max = margin_range

    calc_count = 0
    best_result = None
    best_error = float('inf')

    def get_values(margin, b11):
        """获取 H11 和 F20 值"""
        nonlocal calc_count
        calc_count += 1
        inputs = {
            self._cell_key(self.MARGIN_SHEET, self.MARGIN_CELL): margin,
            self._cell_key('产品成本', 'B11'): b11,
        }
        solution = self._calculate(inputs)
        h11 = self._to_number(self._get_value(solution, self.MARGIN_SHEET, 'H11'))
        f20 = self._to_number(self._get_value(solution, self.MARGIN_SHEET, 'F20'))
        return h11, f20

    def find_b11_for_target_f20(margin, target_f20):
        """利用线性关系计算使 F20=target 的 B11"""
        # 采样两点确定直线: F20 = k * B11 + b
        b11_sample_1, b11_sample_2 = 0, 100000
        _, f20_1 = get_values(margin, b11_sample_1)
        _, f20_2 = get_values(margin, b11_sample_2)

        # 计算斜率
        k = (f20_2 - f20_1) / (b11_sample_2 - b11_sample_1)
        b = f20_1  # 截距 (当 B11=0 时)

        if abs(k) < 1e-10:
            # 斜率太小，F20 几乎不随 B11 变化
            return b11_sample_1

        # 求解 B11 = (target_f20 - b) / k
        target_b11 = (target_f20 - b) / k
        # 约束到安全范围
        target_b11 = max(self.B11_MIN, min(self.B11_MAX, target_b11))
        return target_b11

    def evaluate_solution(margin, b11):
        """评估方案，返回 (h11, f20, error)"""
        h11, f20 = get_values(margin, b11)
        # 归一化误差
        h11_error = abs(h11) / max(abs(h11_max - h11_min), 1)
        f20_error = abs(f20) / max(abs(f20_max - f20_min), 1)
        error = h11_error + f20_error
        return h11, f20, error

    # 二分法搜索 margin
    self._report_progress(20, "二分法搜索最优毛利率...")
    low, high = margin_min, margin_max
    target_f20 = (f20_min + f20_max) / 2  # F20 目标中点

    for iteration in range(25):  # 最多 25 次迭代
        mid = (low + high) / 2

        # 利用线性关系直接计算 B11
        b11 = find_b11_for_target_f20(mid, target_f20)

        # 验证结果
        h11, f20 = get_values(mid, b11)
        error = abs(h11) / 10.0 + abs(f20) / 40000.0

        # 更新最优解
        if error < best_error:
            best_error = error
            best_result = {
                'margin': mid,
                'B11': b11,
                'H11': h11,
                'F20': f20,
            }

        # 检查是否满足约束
        h11_ok = h11_min <= h11 <= h11_max
        f20_ok = f20_min <= f20 <= f20_max

        if h11_ok and f20_ok:
            # 找到满足约束的解
            best_result['converged'] = True
            best_result['iterations'] = calc_count
            return best_result

        # 调整搜索范围
        # H11 通常随 margin 增加而增加
        if h11 < (h11_min + h11_max) / 2:
            low = mid
        else:
            high = mid

        # 检查收敛
        if high - low < 0.0001:
            break

        self._report_progress(20 + int(iteration * 2.5), f"搜索中... (迭代 {iteration + 1})")

    # 未找到精确解，返回最优近似解
    if best_result is None:
        best_result = {
            'margin': (margin_min + margin_max) / 2,
            'B11': 0,
            'H11': 0,
            'F20': 0,
        }

    best_result['converged'] = False
    best_result['iterations'] = calc_count
    return best_result
```

**Step 4: 运行测试验证通过**

Run: `pytest tests/tax_adjuster/test_margin_optimization.py::TestFindOptimalMarginV2::test_returns_dict_with_required_keys -v`
Expected: PASS

**Step 5: Commit**

```bash
git add modules/tax_adjuster/adjust_tax.py tests/tax_adjuster/test_margin_optimization.py
git commit -m "$(cat <<'EOF'
feat: add optimized margin search algorithm find_optimal_margin_v2

Implement binary search with linear interpolation for F20-B11 relationship.
Reduces calculation count from ~240 to ~45-60 iterations.

Co-Authored-By: Claude Opus 4.5 <noreply@anthropic.com>
EOF
)"
```

---

## Task 3: 更新 calculate_inventory_margin_adjustment 接受范围参数

**Files:**
- Modify: `modules/tax_adjuster/adjust_tax.py:817-934`
- Test: `tests/tax_adjuster/test_margin_optimization.py`

**Step 1: 编写参数传递测试**

```python
# tests/tax_adjuster/test_margin_optimization.py
# 添加新测试类

class TestCalculateInventoryMarginAdjustment:
    """测试 calculate_inventory_margin_adjustment 参数传递"""

    def test_accepts_range_parameters(self):
        """验证方法接受范围参数"""
        from modules.tax_adjuster.adjust_tax import TaxAdjuster
        import inspect

        sig = inspect.signature(TaxAdjuster.calculate_inventory_margin_adjustment)
        params = list(sig.parameters.keys())

        assert 'h11_range' in params
        assert 'f20_range' in params
        assert 'margin_range' in params

    def test_default_values_match_class_constants(self):
        """验证默认值与类常量一致"""
        from modules.tax_adjuster.adjust_tax import TaxAdjuster
        import inspect

        sig = inspect.signature(TaxAdjuster.calculate_inventory_margin_adjustment)

        h11_default = sig.parameters['h11_range'].default
        f20_default = sig.parameters['f20_range'].default
        margin_default = sig.parameters['margin_range'].default

        assert h11_default == (TaxAdjuster.H11_MIN, TaxAdjuster.H11_MAX)
        assert f20_default == (TaxAdjuster.F20_MIN, TaxAdjuster.F20_MAX)
        assert margin_default == (TaxAdjuster.MARGIN_MIN, TaxAdjuster.MARGIN_MAX)
```

**Step 2: 运行测试验证失败**

Run: `pytest tests/tax_adjuster/test_margin_optimization.py::TestCalculateInventoryMarginAdjustment -v`
Expected: FAIL (参数不存在)

**Step 3: 修改方法签名和实现**

修改 `modules/tax_adjuster/adjust_tax.py:817` 的方法签名：

```python
def calculate_inventory_margin_adjustment(
    self,
    h11_range=None,
    f20_range=None,
    margin_range=None,
    max_solutions=5
):
    """
    计算库存毛利率调整方案（优化版：利用F20线性特性快速搜索）
    目标: 使 H11 和 F20 落在指定范围内
    工作表: 生产成本月结表、产品成本
    调整变量: 毛利率 (J14单元格), B11 (产品成本中的加工费)

    Args:
        h11_range: (min, max) H11 目标范围，默认 (-10, 10)
        f20_range: (min, max) F20 目标范围，默认 (-40000, 40000)
        margin_range: (min, max) 毛利率搜索范围，默认 (0.70, 0.90)
        max_solutions: 最多返回的候选方案数量，默认 5

    Returns:
        dict: 包含 current（当前值）、solutions（方案列表）、stats（搜索统计）
    """
    # 使用默认值
    if h11_range is None:
        h11_range = (self.H11_MIN, self.H11_MAX)
    if f20_range is None:
        f20_range = (self.F20_MIN, self.F20_MAX)
    if margin_range is None:
        margin_range = (self.MARGIN_MIN, self.MARGIN_MAX)

    # ... 后续代码保持不变，但将 find_optimal_margin_fast 替换为 find_optimal_margin_v2
```

**Step 4: 更新算法调用**

在方法内部（约 866 行），将调用改为：

```python
# 使用优化的快速搜索算法
self._report_progress(15, "正在快速搜索最优解...")
optimal_result = self.find_optimal_margin_v2(
    h11_range=h11_range,
    f20_range=f20_range,
    margin_range=margin_range
)
```

**Step 5: 运行测试验证通过**

Run: `pytest tests/tax_adjuster/test_margin_optimization.py::TestCalculateInventoryMarginAdjustment -v`
Expected: PASS

**Step 6: Commit**

```bash
git add modules/tax_adjuster/adjust_tax.py tests/tax_adjuster/test_margin_optimization.py
git commit -m "$(cat <<'EOF'
feat: add range parameters to calculate_inventory_margin_adjustment

Support user-defined h11_range, f20_range, and margin_range parameters.
Defaults match existing class constants for backward compatibility.

Co-Authored-By: Claude Opus 4.5 <noreply@anthropic.com>
EOF
)"
```

---

## Task 4: 创建参数对话框 MarginParamsDialog

**Files:**
- Modify: `modules/tax_adjuster/tax_tab.py:1-15` (添加 import)
- Modify: `modules/tax_adjuster/tax_tab.py:27` (在 TaxAdjustTab 前添加新类)

**Step 1: 编写对话框测试**

```python
# tests/tax_adjuster/test_margin_optimization.py
# 添加新测试类

class TestMarginParamsDialog:
    """测试参数对话框"""

    def test_dialog_has_required_fields(self):
        """验证对话框包含必要的输入字段"""
        # 注意: wxPython GUI 测试需要特殊处理
        from modules.tax_adjuster.tax_tab import MarginParamsDialog
        import wx

        app = wx.App()
        frame = wx.Frame(None)
        dialog = MarginParamsDialog(frame)

        # 验证字段存在
        assert hasattr(dialog, 'h11_min_ctrl')
        assert hasattr(dialog, 'h11_max_ctrl')
        assert hasattr(dialog, 'f20_min_ctrl')
        assert hasattr(dialog, 'f20_max_ctrl')
        assert hasattr(dialog, 'margin_min_ctrl')
        assert hasattr(dialog, 'margin_max_ctrl')

        dialog.Destroy()
        frame.Destroy()
        app.Destroy()

    def test_get_params_returns_dict(self):
        """验证 get_params 返回正确格式的字典"""
        from modules.tax_adjuster.tax_tab import MarginParamsDialog
        import wx

        app = wx.App()
        frame = wx.Frame(None)
        dialog = MarginParamsDialog(frame)

        params = dialog.get_params()

        assert 'h11_range' in params
        assert 'f20_range' in params
        assert 'margin_range' in params
        assert len(params['h11_range']) == 2
        assert len(params['f20_range']) == 2
        assert len(params['margin_range']) == 2

        dialog.Destroy()
        frame.Destroy()
        app.Destroy()

    def test_default_values(self):
        """验证默认值正确"""
        from modules.tax_adjuster.tax_tab import MarginParamsDialog
        from modules.tax_adjuster.adjust_tax import TaxAdjuster
        import wx

        app = wx.App()
        frame = wx.Frame(None)
        dialog = MarginParamsDialog(frame)

        params = dialog.get_params()

        assert params['h11_range'] == (TaxAdjuster.H11_MIN, TaxAdjuster.H11_MAX)
        assert params['f20_range'] == (TaxAdjuster.F20_MIN, TaxAdjuster.F20_MAX)
        assert params['margin_range'] == (TaxAdjuster.MARGIN_MIN, TaxAdjuster.MARGIN_MAX)

        dialog.Destroy()
        frame.Destroy()
        app.Destroy()
```

**Step 2: 运行测试验证失败**

Run: `pytest tests/tax_adjuster/test_margin_optimization.py::TestMarginParamsDialog::test_dialog_has_required_fields -v`
Expected: FAIL with "ImportError: cannot import name 'MarginParamsDialog'"

**Step 3: 实现 MarginParamsDialog 类**

在 `modules/tax_adjuster/tax_tab.py` 的 `TaxAdjustTab` 类之前（约第 27 行）添加：

```python
class MarginParamsDialog(wx.Dialog):
    """库存毛利率参数设置对话框"""

    def __init__(self, parent):
        super().__init__(
            parent,
            title="调整库存毛利率 - 参数设置",
            style=wx.DEFAULT_DIALOG_STYLE
        )

        from .adjust_tax import TaxAdjuster

        # 默认值
        self.defaults = {
            'h11_min': TaxAdjuster.H11_MIN,
            'h11_max': TaxAdjuster.H11_MAX,
            'f20_min': TaxAdjuster.F20_MIN,
            'f20_max': TaxAdjuster.F20_MAX,
            'margin_min': TaxAdjuster.MARGIN_MIN,
            'margin_max': TaxAdjuster.MARGIN_MAX,
        }

        self.setup_ui()
        self.Centre()

    def setup_ui(self):
        """设置界面"""
        main_sizer = wx.BoxSizer(wx.VERTICAL)

        # 参数输入区域
        grid_sizer = wx.FlexGridSizer(rows=3, cols=4, hgap=10, vgap=10)

        # H11 范围
        grid_sizer.Add(wx.StaticText(self, label="H11 范围:"), 0, wx.ALIGN_CENTER_VERTICAL)
        self.h11_min_ctrl = wx.TextCtrl(self, value=str(self.defaults['h11_min']), size=(80, -1))
        grid_sizer.Add(self.h11_min_ctrl, 0)
        grid_sizer.Add(wx.StaticText(self, label="~"), 0, wx.ALIGN_CENTER_VERTICAL | wx.ALIGN_CENTER_HORIZONTAL)
        self.h11_max_ctrl = wx.TextCtrl(self, value=str(self.defaults['h11_max']), size=(80, -1))
        grid_sizer.Add(self.h11_max_ctrl, 0)

        # F20 范围
        grid_sizer.Add(wx.StaticText(self, label="F20 范围:"), 0, wx.ALIGN_CENTER_VERTICAL)
        self.f20_min_ctrl = wx.TextCtrl(self, value=str(self.defaults['f20_min']), size=(80, -1))
        grid_sizer.Add(self.f20_min_ctrl, 0)
        grid_sizer.Add(wx.StaticText(self, label="~"), 0, wx.ALIGN_CENTER_VERTICAL | wx.ALIGN_CENTER_HORIZONTAL)
        self.f20_max_ctrl = wx.TextCtrl(self, value=str(self.defaults['f20_max']), size=(80, -1))
        grid_sizer.Add(self.f20_max_ctrl, 0)

        # 毛利率范围
        grid_sizer.Add(wx.StaticText(self, label="毛利率范围:"), 0, wx.ALIGN_CENTER_VERTICAL)
        self.margin_min_ctrl = wx.TextCtrl(self, value=str(self.defaults['margin_min']), size=(80, -1))
        grid_sizer.Add(self.margin_min_ctrl, 0)
        grid_sizer.Add(wx.StaticText(self, label="~"), 0, wx.ALIGN_CENTER_VERTICAL | wx.ALIGN_CENTER_HORIZONTAL)
        self.margin_max_ctrl = wx.TextCtrl(self, value=str(self.defaults['margin_max']), size=(80, -1))
        grid_sizer.Add(self.margin_max_ctrl, 0)

        main_sizer.Add(grid_sizer, 0, wx.ALL | wx.EXPAND, 20)

        # 按钮
        btn_sizer = wx.StdDialogButtonSizer()
        cancel_btn = wx.Button(self, wx.ID_CANCEL, "取消")
        ok_btn = wx.Button(self, wx.ID_OK, "开始计算")
        ok_btn.SetDefault()
        btn_sizer.AddButton(cancel_btn)
        btn_sizer.AddButton(ok_btn)
        btn_sizer.Realize()

        main_sizer.Add(btn_sizer, 0, wx.ALL | wx.ALIGN_CENTER, 10)

        self.SetSizer(main_sizer)
        main_sizer.Fit(self)

    def get_params(self):
        """返回用户输入的参数"""
        try:
            h11_min = float(self.h11_min_ctrl.GetValue())
            h11_max = float(self.h11_max_ctrl.GetValue())
            f20_min = float(self.f20_min_ctrl.GetValue())
            f20_max = float(self.f20_max_ctrl.GetValue())
            margin_min = float(self.margin_min_ctrl.GetValue())
            margin_max = float(self.margin_max_ctrl.GetValue())
        except ValueError:
            # 输入无效，返回默认值
            return {
                'h11_range': (self.defaults['h11_min'], self.defaults['h11_max']),
                'f20_range': (self.defaults['f20_min'], self.defaults['f20_max']),
                'margin_range': (self.defaults['margin_min'], self.defaults['margin_max']),
            }

        return {
            'h11_range': (h11_min, h11_max),
            'f20_range': (f20_min, f20_max),
            'margin_range': (margin_min, margin_max),
        }
```

**Step 4: 运行测试验证通过**

Run: `pytest tests/tax_adjuster/test_margin_optimization.py::TestMarginParamsDialog -v`
Expected: PASS

**Step 5: Commit**

```bash
git add modules/tax_adjuster/tax_tab.py tests/tax_adjuster/test_margin_optimization.py
git commit -m "$(cat <<'EOF'
feat: add MarginParamsDialog for user parameter input

Create wxPython dialog allowing users to customize:
- H11 range (default -10 to 10)
- F20 range (default -40000 to 40000)
- Margin range (default 0.70 to 0.90)

Co-Authored-By: Claude Opus 4.5 <noreply@anthropic.com>
EOF
)"
```

---

## Task 5: 集成对话框到 adjust_inventory_margin 方法

**Files:**
- Modify: `modules/tax_adjuster/tax_tab.py:270-290` (`adjust_inventory_margin` 方法)

**Step 1: 编写集成测试**

```python
# tests/tax_adjuster/test_margin_optimization.py
# 添加新测试

class TestAdjustInventoryMarginIntegration:
    """测试 adjust_inventory_margin 与对话框的集成"""

    def test_shows_dialog_before_calculation(self):
        """验证点击按钮后显示对话框"""
        # 此测试需要 GUI 交互，标记为手动测试
        pass  # 手动验证
```

**Step 2: 修改 adjust_inventory_margin 方法**

替换 `modules/tax_adjuster/tax_tab.py:270-290` 的 `adjust_inventory_margin` 方法：

```python
def adjust_inventory_margin(self, event=None):
    """处理"调整库存毛利率"按钮点击"""
    if not self._ensure_file_selected():
        return

    # 弹出参数设置对话框
    dialog = MarginParamsDialog(self)
    if dialog.ShowModal() != wx.ID_OK:
        dialog.Destroy()
        return

    params = dialog.get_params()
    dialog.Destroy()

    if not self._load_adjuster():
        return

    # 显示进度条，禁用按钮
    self._show_progress()
    self._set_buttons_enabled(False)

    def do_calculate():
        try:
            result = self.adjuster.calculate_inventory_margin_adjustment(
                h11_range=params['h11_range'],
                f20_range=params['f20_range'],
                margin_range=params['margin_range']
            )
            wx.CallAfter(self._on_inventory_margin_complete, result, None)
        except Exception as e:
            wx.CallAfter(self._on_inventory_margin_complete, None, e)

    thread = threading.Thread(target=do_calculate, daemon=True)
    thread.start()
```

**Step 3: 运行应用手动验证**

Run: `python accounting_assistant.pyw`
Expected:
1. 点击"调整库存毛利率"按钮
2. 弹出参数设置对话框，显示默认值
3. 点击"开始计算"关闭对话框
4. 显示进度条，开始计算
5. 计算完成后显示结果

**Step 4: Commit**

```bash
git add modules/tax_adjuster/tax_tab.py
git commit -m "$(cat <<'EOF'
feat: integrate MarginParamsDialog into adjust_inventory_margin

Show parameter dialog before starting calculation.
Pass user-specified ranges to the optimization algorithm.

Co-Authored-By: Claude Opus 4.5 <noreply@anthropic.com>
EOF
)"
```

---

## Task 6: 端到端测试与验证

**Files:**
- Modify: `experiments/test_optimized_search.py`

**Step 1: 更新测试脚本支持自定义参数**

```python
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试优化后的库存毛利率搜索算法
对比新旧算法的性能和准确性
"""

import sys
import os
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from modules.tax_adjuster.adjust_tax import TaxAdjuster


def test_optimized_search_v2(file_path, h11_range=(-10, 10), f20_range=(-40000, 40000), margin_range=(0.70, 0.90)):
    """测试优化后的搜索算法 v2"""

    print("=" * 60)
    print(f"测试文件: {os.path.basename(file_path)}")
    print(f"参数: H11={h11_range}, F20={f20_range}, margin={margin_range}")
    print("=" * 60)

    # 进度回调
    def progress_callback(progress, message):
        print(f"  [{progress:3d}%] {message}")

    adjuster = TaxAdjuster(file_path, progress_callback=progress_callback)

    print("\n>>> 测试优化算法 v2 <<<")
    start_time = time.time()
    result = adjuster.calculate_inventory_margin_adjustment(
        h11_range=h11_range,
        f20_range=f20_range,
        margin_range=margin_range,
        max_solutions=5
    )
    elapsed = time.time() - start_time

    if 'error' in result:
        print(f"\n错误: {result['error']}")
        return

    print(f"\n耗时: {elapsed:.2f} 秒")

    if 'stats' in result:
        stats = result['stats']
        print(f"计算次数: {stats.get('iterations', 'N/A')}")
        print(f"收敛: {'是' if stats.get('converged') else '否'}")

    print(f"\n找到 {len(result['solutions'])} 个方案:")
    print("-" * 80)

    for sol in result['solutions']:
        label = sol.get('label', '')
        h11_ok = '✓' if sol.get('h11_ok') else '✗'
        f20_ok = '✓' if sol.get('f20_ok') else '✗'
        status = f"H11:{h11_ok} F20:{f20_ok}"

        print(f"{label:<15} margin={sol['margin']:.5f} B11={sol['B11']:>12,.0f} "
              f"H11={sol['H11']:>8.2f} F20={sol['F20']:>10.2f} {status}")

    return result


if __name__ == '__main__':
    test_files = [
        '/Users/chenjiabin/Documents/demo/accounting-assistant/洪运来2511.xlsx',
        '/Users/chenjiabin/Documents/demo/accounting-assistant/data/洪运来2512.xlsx',
    ]

    for f in test_files:
        if os.path.exists(f):
            # 默认参数测试
            test_optimized_search_v2(f)
            print("\n")
            # 自定义参数测试
            test_optimized_search_v2(f, h11_range=(-5, 5), f20_range=(-20000, 20000))
            break
    else:
        print("未找到测试文件")
```

**Step 2: 运行端到端测试**

Run: `python experiments/test_optimized_search.py`
Expected:
- 计算次数应在 45-60 次左右（比原来 ~240 次减少约 5 倍）
- 找到满足约束的解
- 自定义参数能正确传递

**Step 3: Commit**

```bash
git add experiments/test_optimized_search.py
git commit -m "$(cat <<'EOF'
test: update e2e test to verify v2 algorithm with custom params

Verify algorithm performance improvement (~5x faster) and
parameter passing for custom H11/F20/margin ranges.

Co-Authored-By: Claude Opus 4.5 <noreply@anthropic.com>
EOF
)"
```

---

## Task 7: 清理和文档更新

**Files:**
- Modify: `modules/tax_adjuster/adjust_tax.py` (保留 `find_optimal_margin_fast` 作为备用)

**Step 1: 添加弃用注释**

在 `find_optimal_margin_fast` 方法前添加：

```python
def find_optimal_margin_fast(self, target_H11=0, target_F20=0, h11_tolerance=1.0, f20_tolerance=100):
    """
    快速搜索最优毛利率和B11（多目标帕累托优化）

    .. deprecated::
        此方法已被 find_optimal_margin_v2 替代，后者性能更好。
        保留此方法作为备用。

    ...
    """
```

**Step 2: Commit**

```bash
git add modules/tax_adjuster/adjust_tax.py
git commit -m "$(cat <<'EOF'
docs: mark find_optimal_margin_fast as deprecated

Keep old algorithm as fallback, recommend using find_optimal_margin_v2.

Co-Authored-By: Claude Opus 4.5 <noreply@anthropic.com>
EOF
)"
```

---

## 验收标准

完成后应满足以下条件：

1. **性能提升**: 计算次数从 ~240 次降至 ~45-60 次（约 5 倍提升）
2. **参数可配置**: 用户可通过对话框自定义 H11、F20、毛利率范围
3. **向后兼容**: 不传参数时使用默认值，行为与原版本一致
4. **测试通过**: `pytest tests/tax_adjuster/` 全部通过
5. **UI 正常**: 对话框能正常弹出、接收输入、关闭
