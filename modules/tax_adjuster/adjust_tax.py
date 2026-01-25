#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
税负调整脚本
自动从Excel读取数据，根据目标计算调整方案
"""

import os
import openpyxl
from openpyxl import load_workbook
import re


class TaxAdjuster:
    """税负调整器"""

    def __init__(self, file_path):
        self.file_path = file_path
        self.wb_val = load_workbook(file_path, data_only=True)
        self.wb_formula = load_workbook(file_path, data_only=False)

        # 自动检测参数
        self.T5 = self.wb_val['销售成本']['T5'].value or 0
        self.prev_profit = self._extract_prev_profit()

    def _extract_prev_profit(self):
        """从E30公式中提取上期累计利润"""
        formula = self.wb_formula['测算表']['E30'].value
        if formula and isinstance(formula, str):
            match = re.match(r'=([0-9.]+)\+', formula)
            if match:
                return float(match.group(1))
        return 0

    def get_current_data(self):
        """获取当前数据"""
        ws = self.wb_val['测算表']
        ws_sale = self.wb_val['销售成本']
        E18 = ws['E18'].value or 0
        # E29 = E18/12*12 = E18，直接用 E18 值而不是读取缓存
        # 因为修改 E18 后，E29 的缓存值可能是 None 或旧值
        E29 = E18
        return {
            'E17': ws['E17'].value or 0,
            'E18': E18,
            'E21': ws['E21'].value or 0,
            'E22': ws['E22'].value or 0,
            'E29': E29,
            'E30': ws['E30'].value or 0,
            'E31': ws['E31'].value or 0,
            'G22': ws['G22'].value or 0,
            'B47': ws['B47'].value or 0,  # 利润总额
            'G25': ws['G25'].value or 1,
            'J12': ws_sale['J12'].value or 0,
            'B2': ws['B2'].value or 0,
        }

    def calculate_tax(self, income):
        """累进税率计算"""
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

    def calculate_B47_from_G25(self, G25):
        """根据G25计算B47(利润总额)"""
        ws = self.wb_val['测算表']
        ws_sale = self.wb_val['销售成本']

        # 从Excel读取数据
        B5 = ws_sale['B5'].value or 0
        C5 = ws_sale['C5'].value or 0
        F5 = ws_sale['F5'].value or 0
        H5 = ws_sale['H5'].value or 0
        D6 = ws_sale['D6'].value or 0
        E6 = ws_sale['E6'].value or 0
        F6 = ws_sale['F6'].value or 0
        D7 = ws_sale['D7'].value or 0
        E7 = ws_sale['E7'].value or 0
        F7 = ws_sale['F7'].value or 0
        D8 = ws_sale['D8'].value or 0
        E8 = ws_sale['E8'].value or 0
        F8 = ws_sale['F8'].value or 0

        # 计算销售成本
        # J12 = SUM(J5:J11) = J5 + J6 + J7 + J8
        # J5 = H5 * I5, I5 = (C5 + V5) / (B5 + F5), V5 = T5 * G25
        # J6 = G6 (当B6=C6=0时), G6 = E6/D6*F6*G25
        # J7 = G7 (当B7=C7=0时), G7 = E7/D7*F7*G25
        # J8 = G8 (当B8=C8=0时), G8 = E8/D8*F8*G25
        V5 = self.T5 * G25
        I5 = (C5 + V5) / (B5 + F5) if (B5 + F5) else 0
        J5 = H5 * I5
        G6 = E6 / D6 * F6 * G25 if D6 else 0
        G7 = E7 / D7 * F7 * G25 if D7 else 0
        G8 = E8 / D8 * F8 * G25 if D8 else 0
        J12 = J5 + G6 + G7 + G8

        # 计算利润
        # B23 = 营业利润 = B2 - B3 - B13 - B18 (B3 = J12 + K7, K7通常为0)
        # B47 = 利润总额 = B23 - B25 - B37 + B44 - B45
        B2 = ws['B2'].value or 0
        B13 = ws['B13'].value or 0
        B18 = ws['B18'].value or 0
        B25 = ws['B25'].value or 0  # 管理费用
        B37 = ws['B37'].value or 0  # 财务费用
        B44 = ws['B44'].value or 0  # 营业外收入
        B45 = ws['B45'].value or 0  # 营业外支出

        B23 = B2 - J12 - B13 - B18  # 营业利润
        B47 = B23 - B25 - B37 + B44 - B45  # 利润总额

        return B47, J12

    def find_G25_for_target_B47(self, target_B47):
        """二分法查找G25"""
        low, high = 0.85, 1.00
        while high - low > 1e-10:
            mid = (low + high) / 2
            B47_calc, _ = self.calculate_B47_from_G25(mid)
            if B47_calc > target_B47:
                low = mid
            else:
                high = mid
        return mid

    def _get_G22_constant(self):
        """从G22公式中提取常量（税负率常量）"""
        formula = self.wb_formula['测算表']['G22'].value
        if formula and isinstance(formula, str):
            # 公式格式: =E17*常量-E22
            match = re.search(r'E17\*([0-9.]+)', formula)
            if match:
                return float(match.group(1))
        return 0.00414  # 默认值

    def calculate_annual_profit_adjustment(self):
        """
        计算年利润调整方案
        目标: 使 G22 = 0.00 (±0.009)
        调整: E18 (年利润总额)
        公式分析:
          G22 = E17 * 常量 - E22
          E22 = E21 / 2
          E21 = calculate_tax(E18)
        算法: 直接计算
          target_E21 = 2 * E17 * 常量
          target_E18 = reverse_calculate_income(target_E21)
        """
        current = self.get_current_data()
        E17 = current['E17']

        # 获取G22公式中的常量
        constant = self._get_G22_constant()

        # 计算目标值
        # G22 = E17 * constant - E22 = 0
        # E22 = E21 / 2
        # 所以 E21 = 2 * E17 * constant
        target_E21 = 2 * E17 * constant
        target_E18 = self.reverse_calculate_income(target_E21)

        # 验证
        verify_E21 = self.calculate_tax(target_E18)
        verify_E22 = verify_E21 / 2
        verify_G22 = E17 * constant - verify_E22

        return {
            'current': {
                'E17': E17,
                'E18': current['E18'],
                'E21': current['E21'],
                'E22': current['E22'],
                'G22': current['G22'],
            },
            'target': {
                'E18': target_E18,
            },
            'verify': {
                'E21': verify_E21,
                'E22': verify_E22,
                'G22': verify_G22,
            },
            'constant': constant,
        }

    def calculate_monthly_profit_adjustment(self):
        """
        计算月利润调整方案
        目标: 使 E31 = 0.00 (±0.009)
        调整: G25 (成本系数/毛利率)
        公式分析:
          E31 = E30 - E29
          E30 = prev_profit + B47
          E29 = E18 (年利润总额)
        算法: 二分法查找G25
          target_B47 = E29 - prev_profit
          使用 find_G25_for_target_B47() 查找G25
        """
        current = self.get_current_data()

        E29 = current['E29']

        # 计算目标B47
        # E31 = E30 - E29 = 0
        # E30 = prev_profit + B47
        # 所以 B47 = E29 - prev_profit
        target_B47 = E29 - self.prev_profit

        # 二分法查找G25
        target_G25 = self.find_G25_for_target_B47(target_B47)

        # 验证
        verify_B47, verify_J12 = self.calculate_B47_from_G25(target_G25)
        verify_E30 = self.prev_profit + verify_B47
        verify_E31 = verify_E30 - E29

        return {
            'current': {
                'G25': current['G25'],
                'B47': current['B47'],
                'E29': E29,
                'E30': current['E30'],
                'E31': current['E31'],
                'J12': current['J12'],
            },
            'target': {
                'G25': target_G25,
                'B47': target_B47,
            },
            'verify': {
                'B47': verify_B47,
                'E30': verify_E30,
                'E31': verify_E31,
                'J12': verify_J12,
            },
            'prev_profit': self.prev_profit,
        }

    def _get_E14_formula_parts(self):
        """从生产成本月结表E14公式中提取参数 (公式: E14=a/b*c)"""
        try:
            ws_formula = self.wb_formula['生产成本月结表']
            formula = ws_formula['E14'].value
            if formula and isinstance(formula, str):
                # 尝试解析公式，格式可能是: =A/B*C 或类似
                # 返回 a, b, c 的单元格引用或值
                return formula
        except KeyError:
            pass
        return None

    def _get_margin_from_E14(self):
        """从E14公式中获取毛利率（除数b）"""
        try:
            ws_formula = self.wb_formula['生产成本月结表']
            formula = ws_formula['E14'].value
            if formula and isinstance(formula, str):
                # 公式格式: =a/b*c，提取b
                # 常见格式: =D14/1.08*E$3 或类似
                match = re.search(r'/([0-9.]+)\*', formula)
                if match:
                    return float(match.group(1))
        except KeyError:
            pass
        return 1.08  # 默认值

    def calculate_H11_from_margin(self, margin):
        """根据毛利率计算H11"""
        try:
            ws_val = self.wb_val['生产成本月结表']
            ws_formula = self.wb_formula['生产成本月结表']

            # 读取相关数据
            D14 = ws_val['D14'].value or 0
            E3 = ws_val['E3'].value or 0

            # 计算E14 = D14 / margin * E3
            E14 = D14 / margin * E3 if margin else 0

            # H11的计算依赖具体公式，这里假设H11与E14相关
            # 需要根据实际公式调整
            H8 = ws_val['H8'].value or 0
            H9 = ws_val['H9'].value or 0
            H10 = ws_val['H10'].value or 0

            # 假设 H11 = 某个计算结果 - E14 相关值
            # 实际公式需要根据Excel确定
            G11 = ws_val['G11'].value or 0
            E11 = ws_val['E11'].value or 0
            F11 = ws_val['F11'].value or 0

            # 简化计算：H11 = G11 - (E11 + F11) 或类似
            # 这里返回当前H11作为基准
            H11 = ws_val['H11'].value or 0

            return H11
        except KeyError:
            return 0

    def calculate_F20_from_B11(self, b11):
        """根据加工费B11计算F20"""
        try:
            ws_val = self.wb_val['产品成本']

            # F20的计算依赖B11
            # 实际公式需要根据Excel确定
            F20 = ws_val['F20'].value or 0

            return F20
        except KeyError:
            return 0

    def find_margin_for_target_H11(self, target=0, tolerance=5.0):
        """二分法查找毛利率使H11接近target"""
        low, high = 1.01, 1.20

        while high - low > 0.0001:
            mid = (low + high) / 2
            H11 = self.calculate_H11_from_margin(mid)

            if H11 > target:
                low = mid
            else:
                high = mid

        return mid

    def find_B11_for_target_F20(self, target=0, tolerance=20000.0):
        """二分法查找加工费B11使F20接近target"""
        try:
            ws_val = self.wb_val['产品成本']
            current_B11 = ws_val['B11'].value or 0
        except KeyError:
            current_B11 = 0

        low = current_B11 * 0.5
        high = current_B11 * 1.5

        while high - low > 1:
            mid = (low + high) / 2
            F20 = self.calculate_F20_from_B11(mid)

            if F20 > target:
                low = mid
            else:
                high = mid

        return mid

    def calculate_inventory_margin_adjustment(self):
        """
        计算库存毛利率调整方案
        目标: 使 H11 = 0.00 (±5.0), F20 = 0.00 (±20000.00)
        工作表: 生产成本月结表、产品成本
        调整变量:
          - 毛利率 (E14公式中的除数b, 公式: E14=a/b*c)
          - 加工费 (产品成本!B11)
        算法: 两阶段二分法
          1. 先调整毛利率使H11接近0
          2. 再调整加工费使F20接近0
        """
        try:
            ws_cost = self.wb_val['生产成本月结表']
            ws_product = self.wb_val['产品成本']
        except KeyError as e:
            return {
                'error': f'找不到工作表: {e}',
                'current': {},
                'target': {},
                'verify': {},
            }

        # 获取当前值
        current_margin = self._get_margin_from_E14()
        current_H11 = ws_cost['H11'].value or 0

        try:
            current_B11 = ws_product['B11'].value or 0
            current_F20 = ws_product['F20'].value or 0
        except:
            current_B11 = 0
            current_F20 = 0

        # 阶段1: 调整毛利率使H11接近0
        target_margin = self.find_margin_for_target_H11(target=0)
        verify_H11 = self.calculate_H11_from_margin(target_margin)

        # 阶段2: 调整加工费使F20接近0
        target_B11 = self.find_B11_for_target_F20(target=0)
        verify_F20 = self.calculate_F20_from_B11(target_B11)

        return {
            'current': {
                'margin': current_margin,
                'H11': current_H11,
                'B11': current_B11,
                'F20': current_F20,
            },
            'target': {
                'margin': target_margin,
                'B11': target_B11,
            },
            'verify': {
                'H11': verify_H11,
                'F20': verify_F20,
            },
        }
