#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
税负调整脚本
自动从Excel读取数据，根据目标税负率计算调整方案
"""

import os
import openpyxl
from openpyxl import load_workbook
import argparse
import sys
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
        return {
            'E17': ws['E17'].value or 0,
            'E18': ws['E18'].value or 0,
            'E21': ws['E21'].value or 0,
            'E29': ws['E29'].value or 0,
            'E30': ws['E30'].value or 0,
            'E31': ws['E31'].value or 0,
            'B46': ws['B46'].value or 0,
            'G25': ws['G25'].value or 1,  # 默认系数为1
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

    def calculate_B46_from_G25(self, G25):
        """根据G25计算B46"""
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

        # 计算销售成本
        V5 = self.T5 * G25
        I5 = (C5 + V5) / (B5 + F5) if (B5 + F5) else 0
        J5 = H5 * I5
        G6 = E6 / D6 * F6 * G25 if D6 else 0
        G7 = E7 / D7 * F7 * G25 if D7 else 0
        J12 = J5 + G6 + G7

        # 计算利润
        B2 = ws['B2'].value or 0
        B13 = ws['B13'].value or 0
        B18 = ws['B18'].value or 0
        B24 = ws['B24'].value or 0
        B36 = ws['B36'].value or 0

        B22 = B2 - J12 - B13 - B18
        B46 = B22 - B24 - B36

        return B46, J12

    def find_G25_for_target_B46(self, target_B46):
        """二分法查找G25"""
        low, high = 0.85, 1.00
        while high - low > 1e-10:
            mid = (low + high) / 2
            B46_calc, _ = self.calculate_B46_from_G25(mid)
            if B46_calc > target_B46:
                low = mid
            else:
                high = mid
        return mid

    def calculate_adjustment(self, target_rate):
        """计算调整方案"""
        current = self.get_current_data()

        E17 = current['E17']
        target_E21 = E17 * target_rate
        target_E18 = self.reverse_calculate_income(target_E21)

        target_E29 = target_E18 / 12 * 11
        target_B46 = target_E29 - self.prev_profit
        target_G25 = self.find_G25_for_target_B46(target_B46)

        verify_B46, verify_J12 = self.calculate_B46_from_G25(target_G25)
        verify_E30 = self.prev_profit + verify_B46
        verify_E31 = verify_E30 - target_E29
        verify_E21 = self.calculate_tax(target_E18)
        verify_rate = verify_E21 / E17

        return {
            'current': current,
            'target': {'rate': target_rate, 'E18': target_E18, 'E21': target_E21, 'G25': target_G25, 'B46': target_B46},
            'verify': {'B46': verify_B46, 'E30': verify_E30, 'E31': verify_E31, 'E21': verify_E21, 'rate': verify_rate, 'J12': verify_J12},
        }

    def apply_adjustment(self, target_G25, target_E18, output_path=None):
        """应用调整"""
        ws = self.wb_formula['测算表']
        ws['G25'] = target_G25
        ws['E18'] = target_E18
        save_path = output_path or self.file_path
        # 如果文件已存在，先删除
        if os.path.exists(save_path):
            os.remove(save_path)
        self.wb_formula.save(save_path)
        return save_path


def main():
    parser = argparse.ArgumentParser(description='税负调整工具')
    parser.add_argument('file', help='Excel文件路径')
    parser.add_argument('--rate', type=float, default=0.00414, help='目标税负率 (默认 0.00414)')
    parser.add_argument('--apply', action='store_true', help='应用修改到文件')
    parser.add_argument('--output', type=str, help='输出文件路径')

    args = parser.parse_args()

    try:
        adjuster = TaxAdjuster(args.file)
        result = adjuster.calculate_adjustment(args.rate)

        current = result['current']
        target = result['target']
        verify = result['verify']

        print()
        print("=" * 60)
        print(" 调整目标")
        print("=" * 60)
        print(f"  税负率:  {target['rate']*100:.4f}%")
        print(f"  E31:     -10 ~ 10 之间")
        print()

        print("=" * 60)
        print(" 需要调整的数据")
        print("=" * 60)
        print()
        print(f"  G25 (成本系数):   {current['G25']}  →  {target['G25']:.9f}")
        print(f"  E18 (年利润总额): {current['E18']:.2f}  →  {target['E18']:.2f}")
        print()

        print("=" * 60)
        print(" 调整前后对比")
        print("=" * 60)
        print(f"  {'项目':<16} {'调整前':>14} {'调整后':>14} {'变化':>12}")
        print(f"  {'-'*56}")
        print(f"  {'G25 成本系数':<14} {current['G25']:>14.9f} {target['G25']:>14.9f} {(target['G25']-current['G25'])/current['G25']*100:>+11.2f}%")
        print(f"  {'E18 年利润总额':<12} {current['E18']:>14,.2f} {target['E18']:>14,.2f} {target['E18']-current['E18']:>+12,.2f}")
        print(f"  {'J12 销售成本':<13} {current['J12']:>14,.2f} {verify['J12']:>14,.2f} {verify['J12']-current['J12']:>+12,.2f}")
        print(f"  {'B46 当月利润':<13} {current['B46']:>14,.2f} {verify['B46']:>14,.2f} {verify['B46']-current['B46']:>+12,.2f}")
        print(f"  {'E21 年应纳税额':<12} {current['E21']:>14,.2f} {verify['E21']:>14,.2f} {verify['E21']-current['E21']:>+12,.2f}")
        print(f"  {'E31 差异':<15} {current['E31']:>14,.2f} {verify['E31']:>14,.2f} {verify['E31']-current['E31']:>+12,.2f}")
        cur_rate = current['E21']/current['E17']*100
        new_rate = verify['rate']*100
        print(f"  {'税负率':<16} {cur_rate:>13.4f}% {new_rate:>13.4f}% {new_rate-cur_rate:>+11.4f}%")
        print()

        print("=" * 60)
        print(" 验证结果")
        print("=" * 60)
        rate_ok = abs(verify['rate'] - target['rate']) < 0.00001
        e31_ok = -10 <= verify['E31'] <= 10
        print(f"  税负率: {verify['rate']*100:.4f}%  {'✓' if rate_ok else '✗'}")
        print(f"  E31:    {verify['E31']:.2f}  {'✓' if e31_ok else '✗'}")
        print()

        if args.apply:
            save_path = adjuster.apply_adjustment(target['G25'], target['E18'], args.output)
            print("=" * 60)
            print(f" 已保存至: {save_path}")
            print("=" * 60)
            print()

    except Exception as e:
        print(f"错误: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()
