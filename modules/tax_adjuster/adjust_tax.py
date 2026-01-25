#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
税负调整脚本
使用 formulas + openpyxl 实现纯 Python Excel 公式计算
"""

import os
import shutil
from datetime import datetime
import formulas
import openpyxl


class TaxAdjuster:
    """税负调整器 - 使用 formulas + openpyxl 实现"""

    # G25 安全范围
    G25_MIN = 0.85
    G25_MAX = 1.00

    # E18 安全范围（年利润总额）
    E18_MIN = 0
    E18_MAX = 10_000_000

    # 毛利率安全范围
    MARGIN_MIN = 0.70
    MARGIN_MAX = 0.90

    # B11 安全范围（产品成本加工费）
    B11_MIN = 0
    B11_MAX = 500_000

    # F20 安全范围（生产成本月结表）
    F20_MIN = -20_000
    F20_MAX = 20_000

    # H11 安全范围（生产成本月结表）
    H11_MIN = -5
    H11_MAX = 5

    # 毛利率单元格配置
    MARGIN_CELL = 'J14'
    MARGIN_SHEET = '生产成本月结表'

    def __init__(self, file_path, progress_callback=None):
        """初始化，保存文件路径

        Args:
            file_path: Excel 文件路径
            progress_callback: 进度回调函数，签名为 callback(progress, message)
                              progress: 0-100 的进度值
                              message: 进度描述文字
        """
        self.file_path = os.path.abspath(file_path)
        self._filename = os.path.basename(self.file_path)
        self.temp_file_path = None  # 临时副本文件路径
        self._model = None
        self._progress_callback = progress_callback

    def _report_progress(self, progress, message=""):
        """报告进度"""
        if self._progress_callback:
            self._progress_callback(progress, message)

    def _create_temp_copy(self):
        """创建原文件的临时副本"""
        if self.temp_file_path is None:
            base, ext = os.path.splitext(self.file_path)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            self.temp_file_path = f"{base}_temp_{timestamp}{ext}"
            shutil.copy2(self.file_path, self.temp_file_path)
        return self.temp_file_path

    def _cleanup_temp_file(self):
        """清理临时副本文件"""
        if self.temp_file_path and os.path.exists(self.temp_file_path):
            try:
                os.remove(self.temp_file_path)
            except Exception:
                pass  # 忽略清理失败
            self.temp_file_path = None

    def _cell_key(self, sheet_name, cell):
        """生成 formulas 单元格引用键"""
        return f"'[{self._filename}]{sheet_name}'!{cell}"

    def _load_model(self):
        """加载 formulas ExcelModel"""
        if self._model is None:
            temp_path = self._create_temp_copy()
            # 更新 _filename 为临时文件名，因为 formulas 使用文件名作为键的一部分
            self._filename = os.path.basename(temp_path)
            self._model = formulas.ExcelModel().loads(temp_path).finish()

    def _unload_model(self, save_to_original=False):
        """卸载模型并清理

        Args:
            save_to_original: 如果为 True，将副本内容复制回原文件
        """
        self._model = None

        # 如果需要保存到原文件，先复制再清理
        if save_to_original and self.temp_file_path:
            shutil.copy2(self.temp_file_path, self.file_path)

        # 清理临时文件
        self._cleanup_temp_file()

    def _get_value(self, solution, sheet_name, cell):
        """从 formulas solution 获取单元格值"""
        key = self._cell_key(sheet_name, cell)
        if key in solution:
            val = solution[key].value
            # formulas 返回 numpy 数组，需要提取标量值
            if hasattr(val, '__iter__') and not isinstance(val, str):
                try:
                    return val[0][0] if len(val) > 0 and hasattr(val[0], '__len__') else val[0]
                except (IndexError, TypeError):
                    return val
            return val
        return None

    def _calculate(self, inputs=None):
        """使用 formulas 计算，返回 solution"""
        if inputs is None:
            inputs = {}
        return self._model.calculate(inputs=inputs)

    def _to_number(self, value, default=0):
        """将值转换为数字"""
        if value is None:
            return default
        if isinstance(value, (int, float)):
            return value
        if isinstance(value, str):
            try:
                return float(value)
            except ValueError:
                return default
        return default

    def _check_range(self, value, min_val, max_val, name):
        """检查值是否在安全范围内，返回 (is_safe, message)"""
        if value < min_val:
            return False, f"{name} ({value:.4f}) 低于安全下限 ({min_val})"
        if value > max_val:
            return False, f"{name} ({value:.4f}) 超过安全上限 ({max_val})"
        return True, None

    def _check_margin_cell(self):
        """检查 J14 单元格是否有有效毛利率值

        Returns:
            (is_valid, margin_value, error_message)
        """
        try:
            wb = openpyxl.load_workbook(self.temp_file_path or self.file_path, data_only=True)
            ws = wb[self.MARGIN_SHEET]
            value = ws[self.MARGIN_CELL].value
            wb.close()

            if value is None:
                return False, None, f"请在 Excel 的 '{self.MARGIN_SHEET}' 工作表 {self.MARGIN_CELL} 单元格中设置毛利率值"

            margin = self._to_number(value, None)
            if margin is None or margin <= 0 or margin > 2:
                return False, None, f"'{self.MARGIN_SHEET}' 工作表 {self.MARGIN_CELL} 单元格的毛利率值无效 ({value})，请设置有效的毛利率（如 0.8411）"

            return True, margin, None

        except KeyError:
            return False, None, f"找不到工作表 '{self.MARGIN_SHEET}'"
        except Exception as e:
            return False, None, f"读取毛利率单元格时出错: {e}"

    def _get_margin(self):
        """获取 J14 单元格的毛利率值"""
        is_valid, margin, error_msg = self._check_margin_cell()
        if not is_valid:
            raise ValueError(error_msg)
        return margin

    def get_current_data(self):
        """获取当前数据（使用 formulas 读取实际计算值）"""
        self._load_model()
        try:
            solution = self._calculate()

            return {
                'E17': self._to_number(self._get_value(solution, '测算表', 'E17')),
                'E18': self._to_number(self._get_value(solution, '测算表', 'E18')),
                'E21': self._to_number(self._get_value(solution, '测算表', 'E21')),
                'E22': self._to_number(self._get_value(solution, '测算表', 'E22')),
                'E29': self._to_number(self._get_value(solution, '测算表', 'E29')),
                'E30': self._to_number(self._get_value(solution, '测算表', 'E30')),
                'E31': self._to_number(self._get_value(solution, '测算表', 'E31')),
                'G22': self._to_number(self._get_value(solution, '测算表', 'G22')),
                'B47': self._to_number(self._get_value(solution, '测算表', 'B47')),
                'G25': self._to_number(self._get_value(solution, '测算表', 'G25'), 1),
                'J12': self._to_number(self._get_value(solution, '销售成本', 'J12')),
                'B2': self._to_number(self._get_value(solution, '测算表', 'B2')),
            }
        finally:
            self._unload_model()

    def calculate_tax(self, income):
        """累进税率计算（个体工商户经营所得税率表）"""
        if income <= 30000:
            return income * 0.05
        elif income <= 90000:
            return income * 0.1 - 1500
        elif income <= 300000:
            return income * 0.2 - 10500
        elif income <= 500000:
            return income * 0.3 - 40500
        else:
            return income * 0.35 - 65500

    def reverse_calculate_income(self, tax):
        """根据税额反推应纳税所得额"""
        if tax <= 1500:
            return tax / 0.05
        elif tax <= 7500:
            return (tax + 1500) / 0.1
        elif tax <= 49500:
            return (tax + 10500) / 0.2
        elif tax <= 109500:
            return (tax + 40500) / 0.3
        else:
            return (tax + 65500) / 0.35

    def _get_G22_at_E18(self, E18):
        """设置 E18 并获取计算后的 G22 值"""
        inputs = {self._cell_key('测算表', 'E18'): E18}
        solution = self._calculate(inputs)
        return self._to_number(self._get_value(solution, '测算表', 'G22'))

    def _get_E31_at_G25(self, G25):
        """设置 G25 并获取计算后的 E31 值"""
        inputs = {self._cell_key('测算表', 'G25'): G25}
        solution = self._calculate(inputs)
        return self._to_number(self._get_value(solution, '测算表', 'E31'))

    def _get_B47_at_G25(self, G25):
        """设置 G25 并获取计算后的 B47 值"""
        inputs = {self._cell_key('测算表', 'G25'): G25}
        solution = self._calculate(inputs)
        return self._to_number(self._get_value(solution, '测算表', 'B47'))

    def find_E18_for_target_G22(self, target_G22=0, tolerance=0.009):
        """
        二分法查找 E18，使 G22 接近目标值
        返回 (E18, G22, is_in_range, boundary_info)
        """
        low, high = self.E18_MIN, self.E18_MAX

        # 先计算边界值
        G22_at_low = self._get_G22_at_E18(low)
        G22_at_high = self._get_G22_at_E18(high)

        # 确定 G22 随 E18 变化的方向
        # E18 增加 -> E21(税额) 增加 -> E22 增加 -> G22 减小
        # 所以 G22 随 E18 增加而减小

        # 检查目标是否在可达范围内
        min_G22 = min(G22_at_low, G22_at_high)
        max_G22 = max(G22_at_low, G22_at_high)

        if target_G22 < min_G22:
            # 目标太小，使用上界
            final_E18 = high if G22_at_high < G22_at_low else low
            final_G22 = self._get_G22_at_E18(final_E18)
            return final_E18, final_G22, False, {
                'reason': 'target_too_low',
                'target_G22': target_G22,
                'min_G22': min_G22,
            }
        if target_G22 > max_G22:
            # 目标太大，使用下界
            final_E18 = low if G22_at_low > G22_at_high else high
            final_G22 = self._get_G22_at_E18(final_E18)
            return final_E18, final_G22, False, {
                'reason': 'target_too_high',
                'target_G22': target_G22,
                'max_G22': max_G22,
            }

        # 二分法查找
        mid = (low + high) / 2
        for _ in range(50):  # 最多迭代50次
            if high - low < 1:  # E18 精度到 1 元
                break
            mid = (low + high) / 2
            G22 = self._get_G22_at_E18(mid)

            if abs(G22 - target_G22) < tolerance:
                return mid, G22, True, None

            # G22 随 E18 增加而减小
            if G22 > target_G22:
                low = mid
            else:
                high = mid

        final_G22 = self._get_G22_at_E18(mid)
        return mid, final_G22, abs(final_G22 - target_G22) < tolerance, None

    def find_G25_for_target_E31(self, target_E31=0, tolerance=0.009):
        """
        二分法查找 G25，使 E31 接近目标值
        返回 (G25, E31, B47, J12, is_in_range, boundary_info)
        """
        low, high = self.G25_MIN, self.G25_MAX

        def get_values_at_G25(g25):
            inputs = {self._cell_key('测算表', 'G25'): g25}
            solution = self._calculate(inputs)
            return (
                self._to_number(self._get_value(solution, '测算表', 'E31')),
                self._to_number(self._get_value(solution, '测算表', 'B47')),
                self._to_number(self._get_value(solution, '销售成本', 'J12')),
            )

        # 先计算边界值
        E31_at_low, B47_at_low, J12_at_low = get_values_at_G25(low)
        E31_at_high, B47_at_high, J12_at_high = get_values_at_G25(high)

        # E31 = E30 - E29, E30 = prev_profit + B47
        # G25 增加 -> 销售成本 J12 增加 -> B47 减小 -> E30 减小 -> E31 减小

        min_E31 = min(E31_at_low, E31_at_high)
        max_E31 = max(E31_at_low, E31_at_high)

        if target_E31 < min_E31:
            final_G25 = high if E31_at_high < E31_at_low else low
            E31, B47, J12 = get_values_at_G25(final_G25)
            return final_G25, E31, B47, J12, False, {
                'reason': 'target_too_low',
                'target_E31': target_E31,
                'min_E31': min_E31,
                'boundary_G25': final_G25,
            }
        if target_E31 > max_E31:
            final_G25 = low if E31_at_low > E31_at_high else high
            E31, B47, J12 = get_values_at_G25(final_G25)
            return final_G25, E31, B47, J12, False, {
                'reason': 'target_too_high',
                'target_E31': target_E31,
                'max_E31': max_E31,
                'boundary_G25': final_G25,
            }

        # 二分法查找
        mid = (low + high) / 2
        for _ in range(50):
            if high - low < 1e-9:
                break
            mid = (low + high) / 2
            E31, B47, J12 = get_values_at_G25(mid)

            if abs(E31 - target_E31) < tolerance:
                return mid, E31, B47, J12, True, None

            # E31 随 G25 增加而减小
            if E31 > target_E31:
                low = mid
            else:
                high = mid

        E31, B47, J12 = get_values_at_G25(mid)
        return mid, E31, B47, J12, abs(E31 - target_E31) < tolerance, None

    def calculate_combined_adjustment(self):
        """
        整合计算年利润和月毛利调整方案
        同时调整 E18 和 G25，使 G22 = 0 且 E31 = 0

        利用 formulas 公式计算，不在代码中硬编码公式
        """
        self._report_progress(0, "正在加载 Excel 模型...")
        self._load_model()
        try:
            self._report_progress(10, "正在读取当前数据...")
            # 获取当前数据（无输入修改）
            solution = self._calculate()

            original_E18 = self._to_number(self._get_value(solution, '测算表', 'E18'))
            original_G25 = self._to_number(self._get_value(solution, '测算表', 'G25'), 1)

            current = {
                'E17': self._to_number(self._get_value(solution, '测算表', 'E17')),
                'E18': original_E18,
                'E21': self._to_number(self._get_value(solution, '测算表', 'E21')),
                'E22': self._to_number(self._get_value(solution, '测算表', 'E22')),
                'G22': self._to_number(self._get_value(solution, '测算表', 'G22')),
                'G25': original_G25,
                'B47': self._to_number(self._get_value(solution, '测算表', 'B47')),
                'E29': self._to_number(self._get_value(solution, '测算表', 'E29')),
                'E30': self._to_number(self._get_value(solution, '测算表', 'E30')),
                'E31': self._to_number(self._get_value(solution, '测算表', 'E31')),
                'J12': self._to_number(self._get_value(solution, '销售成本', 'J12')),
            }

            # === 第一步: 查找 E18 使 G22 = 0 ===
            self._report_progress(20, "正在搜索最优 E18...")
            target_E18, verify_G22, e18_in_range, e18_boundary = self.find_E18_for_target_G22(
                target_G22=0, tolerance=0.009
            )

            # 检查 E18 是否在安全范围
            e18_safe, e18_msg = self._check_range(
                target_E18, self.E18_MIN, self.E18_MAX, 'E18'
            )

            # 设置新的 E18，计算新的 E29
            inputs_e18 = {self._cell_key('测算表', 'E18'): target_E18}
            solution_e18 = self._calculate(inputs_e18)
            new_E29 = self._to_number(self._get_value(solution_e18, '测算表', 'E29'))
            verify_E21 = self._to_number(self._get_value(solution_e18, '测算表', 'E21'))
            verify_E22 = self._to_number(self._get_value(solution_e18, '测算表', 'E22'))

            # === 第二步: 查找 G25 使 E31 = 0（在新 E18 基础上）===
            self._report_progress(50, "正在搜索最优 G25...")
            # 需要同时设置 E18 和 G25
            def get_values_at_G25_with_E18(g25):
                inputs = {
                    self._cell_key('测算表', 'E18'): target_E18,
                    self._cell_key('测算表', 'G25'): g25,
                }
                sol = self._calculate(inputs)
                return (
                    self._to_number(self._get_value(sol, '测算表', 'E31')),
                    self._to_number(self._get_value(sol, '测算表', 'B47')),
                    self._to_number(self._get_value(sol, '销售成本', 'J12')),
                )

            low, high = self.G25_MIN, self.G25_MAX

            E31_at_low, B47_at_low, J12_at_low = get_values_at_G25_with_E18(low)
            E31_at_high, B47_at_high, J12_at_high = get_values_at_G25_with_E18(high)

            min_E31 = min(E31_at_low, E31_at_high)
            max_E31 = max(E31_at_low, E31_at_high)
            target_E31 = 0
            tolerance = 0.009

            g25_in_range = True
            g25_boundary = None

            if target_E31 < min_E31:
                target_G25 = high if E31_at_high < E31_at_low else low
                verify_E31, verify_B47, verify_J12 = get_values_at_G25_with_E18(target_G25)
                g25_in_range = False
                g25_boundary = {
                    'reason': 'target_too_low',
                    'target_E31': target_E31,
                    'min_E31': min_E31,
                    'boundary_G25': target_G25,
                }
            elif target_E31 > max_E31:
                target_G25 = low if E31_at_low > E31_at_high else high
                verify_E31, verify_B47, verify_J12 = get_values_at_G25_with_E18(target_G25)
                g25_in_range = False
                g25_boundary = {
                    'reason': 'target_too_high',
                    'target_E31': target_E31,
                    'max_E31': max_E31,
                    'boundary_G25': target_G25,
                }
            else:
                # 二分法查找
                mid = (low + high) / 2
                for _ in range(50):
                    if high - low < 1e-9:
                        break
                    mid = (low + high) / 2
                    E31, B47, J12 = get_values_at_G25_with_E18(mid)

                    if abs(E31 - target_E31) < tolerance:
                        break

                    if E31 > target_E31:
                        low = mid
                    else:
                        high = mid

                target_G25 = mid
                verify_E31, verify_B47, verify_J12 = get_values_at_G25_with_E18(target_G25)
                g25_in_range = abs(verify_E31 - target_E31) < tolerance

            # 检查 G25 是否在安全范围
            g25_safe, g25_msg = self._check_range(
                target_G25, self.G25_MIN, self.G25_MAX, 'G25'
            )

            # 读取最终的 E30
            self._report_progress(90, "正在验证结果...")
            final_inputs = {
                self._cell_key('测算表', 'E18'): target_E18,
                self._cell_key('测算表', 'G25'): target_G25,
            }
            final_solution = self._calculate(final_inputs)
            verify_E30 = self._to_number(self._get_value(final_solution, '测算表', 'E30'))

            self._report_progress(100, "计算完成")
            # 构建结果
            result = {
                'current': current,
                'target': {
                    'E18': target_E18,
                    'G25': target_G25,
                    'B47': verify_B47,
                },
                'verify': {
                    'E21': verify_E21,
                    'E22': verify_E22,
                    'G22': verify_G22,
                    'B47': verify_B47,
                    'E29': new_E29,
                    'E30': verify_E30,
                    'E31': verify_E31,
                    'J12': verify_J12,
                },
                'in_range': e18_in_range and g25_in_range,
                'safety_check': {
                    'E18_safe': e18_safe,
                    'E18_msg': e18_msg,
                    'G25_safe': g25_safe,
                    'G25_msg': g25_msg,
                },
            }

            if not e18_in_range and e18_boundary:
                result['e18_boundary_info'] = e18_boundary
            if not g25_in_range and g25_boundary:
                result['boundary_info'] = g25_boundary

            return result

        finally:
            self._unload_model()

    def find_margin_and_B11_for_targets(self, target_H11=0, target_F20=0):
        """
        同时调整毛利率和 B11，使 H11 和 F20 都接近目标值

        优化策略（基于数据规律实验结果）：
        - H11 随 margin 单调递增，随 B11 单调递增
        - F20 随 margin 单调递增，随 B11 单调递减
        - 使用分层采样：先稀疏定位有效区间，再精细搜索
        - B11 用二分法搜索，利用线性插值加速收敛
        """
        # 结果缓存
        cache = {}

        def set_and_get_values(margin, b11):
            """设置毛利率和 B11，返回 H11 和 F20（带缓存）"""
            cache_key = (round(margin, 5), round(b11, 0))
            if cache_key in cache:
                return cache[cache_key]

            inputs = {
                self._cell_key(self.MARGIN_SHEET, self.MARGIN_CELL): margin,
                self._cell_key('产品成本', 'B11'): b11,
            }
            solution = self._calculate(inputs)

            result = (
                self._to_number(self._get_value(solution, self.MARGIN_SHEET, 'H11')),
                self._to_number(self._get_value(solution, self.MARGIN_SHEET, 'F20')),
            )
            cache[cache_key] = result
            return result

        def find_B11_for_H11(target, margin):
            """固定 margin，二分法找 B11 使 H11 接近目标"""
            low, high = self.B11_MIN, self.B11_MAX

            H11_low, F20_low = set_and_get_values(margin, low)
            H11_high, F20_high = set_and_get_values(margin, high)

            if target <= H11_low:
                return low, H11_low, F20_low
            if target >= H11_high:
                return high, H11_high, F20_high

            # 二分法搜索，精度提高到 10
            mid = (low + high) / 2
            for _ in range(25):
                if high - low < 10:  # 提高精度
                    break
                mid = (low + high) / 2
                h11, f20 = set_and_get_values(margin, mid)

                if abs(h11 - target) < 5:
                    return mid, h11, f20

                if h11 < target:
                    low = mid
                else:
                    high = mid

            mid = (low + high) / 2
            h11, f20 = set_and_get_values(margin, mid)
            return mid, h11, f20

        def evaluate_margin(margin):
            """评估给定 margin 时的 |F20| 值"""
            b11, h11, f20 = find_B11_for_H11(target_H11, margin)
            # H11 容差使用安全范围
            if abs(h11 - target_H11) <= self.H11_MAX and b11 >= 0:
                return abs(f20 - target_F20), b11, h11, f20
            return float('inf'), b11, h11, f20

        # 初始化最佳解
        best_score = float('inf')
        best_margin, best_b11, best_h11, best_f20 = 0.9, 0, 0, 0

        GOOD_ENOUGH = 2000  # 提前退出阈值

        # 第一阶段：稀疏采样定位有效区间（0.70-0.90）
        # 关键：需要覆盖 F20 从负变正的过渡区（约 0.82-0.83）
        self._report_progress(30, "第一阶段: 稀疏采样...")
        valid_samples = []
        sample_margins = [0.72, 0.76, 0.80, 0.83, 0.86, 0.89]
        for i, margin in enumerate(sample_margins):
            self._report_progress(30 + int(20 * (i + 1) / len(sample_margins)),
                                  f"采样毛利率 {margin:.2f}...")
            score, b11, h11, f20 = evaluate_margin(margin)
            if score < float('inf'):
                valid_samples.append((margin, score, b11, h11, f20))
                if score < best_score:
                    best_score, best_margin = score, margin
                    best_b11, best_h11, best_f20 = b11, h11, f20

        if not valid_samples:
            score, b11, h11, f20 = evaluate_margin(self.MARGIN_MAX)
            return self.MARGIN_MAX, b11, h11, f20

        # 第二阶段：在最佳点附近用二分法精确定位
        # 利用 F20 从负到正的单调性，找到 F20=0 的点
        self._report_progress(55, "第二阶段: 精细搜索...")
        if best_score >= GOOD_ENOUGH:
            # 找到 F20 符号变化的区间
            sorted_samples = sorted(valid_samples, key=lambda x: x[0])
            left_margin, right_margin = None, None

            for i in range(len(sorted_samples) - 1):
                f20_i = sorted_samples[i][4]
                f20_next = sorted_samples[i + 1][4]
                if f20_i * f20_next < 0:  # 符号变化
                    left_margin = sorted_samples[i][0]
                    right_margin = sorted_samples[i + 1][0]
                    break

            if left_margin is not None:
                # 二分法找 F20=0 的 margin
                for _ in range(10):
                    if right_margin - left_margin < 0.002:
                        break
                    mid_margin = (left_margin + right_margin) / 2
                    score, b11, h11, f20 = evaluate_margin(mid_margin)

                    if score < best_score:
                        best_score, best_margin = score, mid_margin
                        best_b11, best_h11, best_f20 = b11, h11, f20

                    if f20 < 0:
                        left_margin = mid_margin
                    else:
                        right_margin = mid_margin
            else:
                # 没有符号变化，在最佳点附近线性搜索
                search_start = max(self.MARGIN_MIN, best_margin - 0.03)
                search_end = min(self.MARGIN_MAX, best_margin + 0.03)
                margin = search_start
                while margin <= search_end:
                    score, b11, h11, f20 = evaluate_margin(margin)
                    if score < best_score:
                        best_score, best_margin = score, margin
                        best_b11, best_h11, best_f20 = b11, h11, f20
                    margin += 0.01

        self._report_progress(85, "搜索完成")
        return best_margin, best_b11, best_h11, best_f20

    def calculate_inventory_margin_adjustment(self):
        """
        计算库存毛利率调整方案
        目标: 使 H11 = 0.00, F20 = 0.00
        工作表: 生产成本月结表、产品成本
        调整变量: 毛利率 (J14单元格), B11 (产品成本中的加工费)
        """
        self._report_progress(0, "正在加载 Excel 模型...")
        self._load_model()
        try:
            self._report_progress(5, "正在检查毛利率单元格...")
            # 检查 J14 单元格是否有效
            is_valid, current_margin, error_msg = self._check_margin_cell()
            if not is_valid:
                return {
                    'error': error_msg,
                    'current': {},
                    'target': {},
                    'verify': {},
                }

            self._report_progress(10, "正在读取当前数据...")
            # 获取当前值
            solution = self._calculate()
            current_H11 = self._to_number(self._get_value(solution, self.MARGIN_SHEET, 'H11'))
            current_F20 = self._to_number(self._get_value(solution, self.MARGIN_SHEET, 'F20'))
            current_B11 = self._to_number(self._get_value(solution, '产品成本', 'B11'))

            # 调整
            self._report_progress(20, "正在搜索最优毛利率和 B11...")
            target_margin, target_B11, verify_H11, verify_F20 = \
                self.find_margin_and_B11_for_targets(target_H11=0, target_F20=0)

            self._report_progress(90, "正在验证结果...")
            # 检查安全范围
            margin_safe, margin_msg = self._check_range(
                target_margin, self.MARGIN_MIN, self.MARGIN_MAX, '毛利率'
            )
            b11_safe, b11_msg = self._check_range(
                target_B11, self.B11_MIN, self.B11_MAX, 'B11'
            )
            f20_safe, f20_msg = self._check_range(
                verify_F20, self.F20_MIN, self.F20_MAX, 'F20'
            )
            h11_safe, h11_msg = self._check_range(
                verify_H11, self.H11_MIN, self.H11_MAX, 'H11'
            )

            self._report_progress(100, "计算完成")
            return {
                'current': {
                    'margin': current_margin,
                    'H11': current_H11,
                    'F20': current_F20,
                    'B11': current_B11,
                },
                'target': {
                    'margin': target_margin,
                    'B11': target_B11,
                },
                'verify': {
                    'H11': verify_H11,
                    'F20': verify_F20,
                },
                'safety_check': {
                    'margin_safe': margin_safe,
                    'margin_msg': margin_msg,
                    'B11_safe': b11_safe,
                    'B11_msg': b11_msg,
                    'F20_safe': f20_safe,
                    'F20_msg': f20_msg,
                    'H11_safe': h11_safe,
                    'H11_msg': h11_msg,
                },
            }

        except KeyError as e:
            return {
                'error': f'找不到工作表: {e}',
                'current': {},
                'target': {},
                'verify': {},
            }
        finally:
            self._unload_model()

    def _save_to_file(self, values):
        """用 openpyxl 保存值到文件

        Args:
            values: dict of {sheet_name: {cell: value}}
        """
        wb = openpyxl.load_workbook(self.temp_file_path)
        for sheet_name, cells in values.items():
            ws = wb[sheet_name]
            for cell, value in cells.items():
                ws[cell] = value
        wb.save(self.temp_file_path)
        wb.close()
