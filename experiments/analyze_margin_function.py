#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
分析库存毛利率搜索函数特性的实验脚本
目标：确定 H11(margin, B11) 和 F20(margin, B11) 的函数形式
"""

import sys
import os
import numpy as np
from datetime import datetime

# 添加项目路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from modules.tax_adjuster.adjust_tax import TaxAdjuster

def analyze_function_characteristics(file_path):
    """分析函数特性"""

    print(f"=" * 60)
    print(f"分析文件: {os.path.basename(file_path)}")
    print(f"时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"=" * 60)

    adjuster = TaxAdjuster(file_path)
    adjuster._load_model()

    try:
        # 定义采样点
        margins = np.linspace(0.70, 0.90, 11)  # 11 个点
        b11s = np.linspace(0, 500000, 11)      # 11 个点

        # 存储结果
        h11_matrix = np.zeros((len(margins), len(b11s)))
        f20_matrix = np.zeros((len(margins), len(b11s)))

        # 计算缓存
        def get_values(margin, b11):
            inputs = {
                adjuster._cell_key(adjuster.MARGIN_SHEET, adjuster.MARGIN_CELL): margin,
                adjuster._cell_key('产品成本', 'B11'): b11,
            }
            solution = adjuster._calculate(inputs)
            h11 = adjuster._to_number(adjuster._get_value(solution, adjuster.MARGIN_SHEET, 'H11'))
            f20 = adjuster._to_number(adjuster._get_value(solution, adjuster.MARGIN_SHEET, 'F20'))
            return h11, f20

        print("\n正在采样数据点...")
        total = len(margins) * len(b11s)
        count = 0

        for i, margin in enumerate(margins):
            for j, b11 in enumerate(b11s):
                h11, f20 = get_values(margin, b11)
                h11_matrix[i, j] = h11
                f20_matrix[i, j] = f20
                count += 1
                if count % 20 == 0:
                    print(f"  进度: {count}/{total}")

        print(f"\n采样完成，共 {total} 个点")

        # ============ 分析 1: 检查单变量关系 ============
        print("\n" + "=" * 60)
        print("分析1: 单变量关系")
        print("=" * 60)

        # 固定 B11 = 0，观察 H11 vs margin
        print("\n[H11 vs margin] (固定 B11 = 0):")
        h11_vs_margin = h11_matrix[:, 0]
        for m, h in zip(margins, h11_vs_margin):
            print(f"  margin={m:.2f}  H11={h:>10.2f}")

        # 计算差分（检查线性性）
        h11_diff = np.diff(h11_vs_margin)
        print(f"  差分: min={h11_diff.min():.2f}, max={h11_diff.max():.2f}, "
              f"std={h11_diff.std():.2f}")

        # 固定 margin = 0.8，观察 H11 vs B11
        mid_margin_idx = len(margins) // 2
        print(f"\n[H11 vs B11] (固定 margin = {margins[mid_margin_idx]:.2f}):")
        h11_vs_b11 = h11_matrix[mid_margin_idx, :]
        for b, h in zip(b11s, h11_vs_b11):
            print(f"  B11={b:>10,.0f}  H11={h:>10.2f}")

        h11_b11_diff = np.diff(h11_vs_b11)
        print(f"  差分: min={h11_b11_diff.min():.2f}, max={h11_b11_diff.max():.2f}, "
              f"std={h11_b11_diff.std():.2f}")

        # F20 同样分析
        print(f"\n[F20 vs margin] (固定 B11 = 0):")
        f20_vs_margin = f20_matrix[:, 0]
        for m, f in zip(margins, f20_vs_margin):
            print(f"  margin={m:.2f}  F20={f:>10.2f}")

        f20_diff = np.diff(f20_vs_margin)
        print(f"  差分: min={f20_diff.min():.2f}, max={f20_diff.max():.2f}, "
              f"std={f20_diff.std():.2f}")

        print(f"\n[F20 vs B11] (固定 margin = {margins[mid_margin_idx]:.2f}):")
        f20_vs_b11 = f20_matrix[mid_margin_idx, :]
        for b, f in zip(b11s, f20_vs_b11):
            print(f"  B11={b:>10,.0f}  F20={f:>10.2f}")

        f20_b11_diff = np.diff(f20_vs_b11)
        print(f"  差分: min={f20_b11_diff.min():.2f}, max={f20_b11_diff.max():.2f}, "
              f"std={f20_b11_diff.std():.2f}")

        # ============ 分析 2: 检查线性拟合 ============
        print("\n" + "=" * 60)
        print("分析2: 线性拟合检验")
        print("=" * 60)

        # 构建特征矩阵 [1, margin, B11]
        X = []
        y_h11 = []
        y_f20 = []

        for i, margin in enumerate(margins):
            for j, b11 in enumerate(b11s):
                X.append([1, margin, b11])
                y_h11.append(h11_matrix[i, j])
                y_f20.append(f20_matrix[i, j])

        X = np.array(X)
        y_h11 = np.array(y_h11)
        y_f20 = np.array(y_f20)

        # 最小二乘拟合
        coef_h11, residuals_h11, rank, s = np.linalg.lstsq(X, y_h11, rcond=None)
        coef_f20, residuals_f20, rank, s = np.linalg.lstsq(X, y_f20, rcond=None)

        # 计算 R²
        y_h11_pred = X @ coef_h11
        y_f20_pred = X @ coef_f20

        ss_tot_h11 = np.sum((y_h11 - y_h11.mean()) ** 2)
        ss_res_h11 = np.sum((y_h11 - y_h11_pred) ** 2)
        r2_h11 = 1 - ss_res_h11 / ss_tot_h11

        ss_tot_f20 = np.sum((y_f20 - y_f20.mean()) ** 2)
        ss_res_f20 = np.sum((y_f20 - y_f20_pred) ** 2)
        r2_f20 = 1 - ss_res_f20 / ss_tot_f20

        print(f"\nH11 线性拟合: H11 = {coef_h11[0]:.4f} + {coef_h11[1]:.4f}*margin + {coef_h11[2]:.8f}*B11")
        print(f"  R² = {r2_h11:.6f}")
        print(f"  最大残差: {np.abs(y_h11 - y_h11_pred).max():.4f}")

        print(f"\nF20 线性拟合: F20 = {coef_f20[0]:.4f} + {coef_f20[1]:.4f}*margin + {coef_f20[2]:.8f}*B11")
        print(f"  R² = {r2_f20:.6f}")
        print(f"  最大残差: {np.abs(y_f20 - y_f20_pred).max():.4f}")

        # ============ 分析 3: 交互项拟合 ============
        print("\n" + "=" * 60)
        print("分析3: 带交互项的拟合")
        print("=" * 60)

        # 添加交互项 [1, margin, B11, margin*B11]
        X_inter = np.column_stack([X, X[:, 1] * X[:, 2]])

        coef_h11_inter, _, _, _ = np.linalg.lstsq(X_inter, y_h11, rcond=None)
        coef_f20_inter, _, _, _ = np.linalg.lstsq(X_inter, y_f20, rcond=None)

        y_h11_pred_inter = X_inter @ coef_h11_inter
        y_f20_pred_inter = X_inter @ coef_f20_inter

        ss_res_h11_inter = np.sum((y_h11 - y_h11_pred_inter) ** 2)
        r2_h11_inter = 1 - ss_res_h11_inter / ss_tot_h11

        ss_res_f20_inter = np.sum((y_f20 - y_f20_pred_inter) ** 2)
        r2_f20_inter = 1 - ss_res_f20_inter / ss_tot_f20

        print(f"\nH11 带交互项拟合:")
        print(f"  H11 = {coef_h11_inter[0]:.4f} + {coef_h11_inter[1]:.4f}*margin + {coef_h11_inter[2]:.8f}*B11 + {coef_h11_inter[3]:.10f}*margin*B11")
        print(f"  R² = {r2_h11_inter:.6f}")
        print(f"  最大残差: {np.abs(y_h11 - y_h11_pred_inter).max():.4f}")

        print(f"\nF20 带交互项拟合:")
        print(f"  F20 = {coef_f20_inter[0]:.4f} + {coef_f20_inter[1]:.4f}*margin + {coef_f20_inter[2]:.8f}*B11 + {coef_f20_inter[3]:.10f}*margin*B11")
        print(f"  R² = {r2_f20_inter:.6f}")
        print(f"  最大残差: {np.abs(y_f20 - y_f20_pred_inter).max():.4f}")

        # ============ 分析 4: 判断 margin 和 B11 的独立性 ============
        print("\n" + "=" * 60)
        print("分析4: 变量独立性检验")
        print("=" * 60)

        # 检查 H11 对 margin 的变化率是否随 B11 变化
        h11_margin_slopes = []
        for j in range(len(b11s)):
            slope = (h11_matrix[-1, j] - h11_matrix[0, j]) / (margins[-1] - margins[0])
            h11_margin_slopes.append(slope)

        print(f"\nH11 对 margin 的斜率随 B11 变化:")
        for b, s in zip(b11s, h11_margin_slopes):
            print(f"  B11={b:>10,.0f}  斜率={s:>10.2f}")
        print(f"  斜率变化范围: {min(h11_margin_slopes):.2f} ~ {max(h11_margin_slopes):.2f}")

        # 检查 H11 对 B11 的变化率是否随 margin 变化
        h11_b11_slopes = []
        for i in range(len(margins)):
            slope = (h11_matrix[i, -1] - h11_matrix[i, 0]) / (b11s[-1] - b11s[0])
            h11_b11_slopes.append(slope)

        print(f"\nH11 对 B11 的斜率随 margin 变化:")
        for m, s in zip(margins, h11_b11_slopes):
            print(f"  margin={m:.2f}  斜率={s:>10.8f}")
        print(f"  斜率变化范围: {min(h11_b11_slopes):.8f} ~ {max(h11_b11_slopes):.8f}")

        # ============ 结论与建议 ============
        print("\n" + "=" * 60)
        print("结论与优化建议")
        print("=" * 60)

        linear_h11 = r2_h11 > 0.99
        linear_f20 = r2_f20 > 0.99

        print(f"\n1. H11 线性关系: {'是 ✓' if linear_h11 else '否 ✗'} (R²={r2_h11:.4f})")
        print(f"2. F20 线性关系: {'是 ✓' if linear_f20 else '否 ✗'} (R²={r2_f20:.4f})")

        # 检查变量独立性
        margin_slope_var = np.std(h11_margin_slopes) / np.abs(np.mean(h11_margin_slopes)) if np.mean(h11_margin_slopes) != 0 else 0
        independent = margin_slope_var < 0.1  # 变异系数小于10%认为独立

        print(f"3. margin 和 B11 对 H11 的影响独立性: {'是 ✓' if independent else '否 ✗'}")

        if linear_h11 and linear_f20:
            print("\n>>> 建议策略: 线性拟合法")
            print("    - 只需采样约 4-6 个点即可拟合线性函数")
            print("    - 直接解线性方程组求解 margin 和 B11")
            print("    - 最后用 1-2 次计算验证结果")
            print(f"\n>>> 拟合公式:")
            print(f"    H11 = {coef_h11[0]:.4f} + {coef_h11[1]:.4f}*margin + {coef_h11[2]:.8f}*B11")
            print(f"    F20 = {coef_f20[0]:.4f} + {coef_f20[1]:.4f}*margin + {coef_f20[2]:.8f}*B11")
            print(f"\n>>> 预期效率提升: 从 ~600 次计算 → ~10 次计算 (60x 提升)")
        elif independent:
            print("\n>>> 建议策略: 分步二分搜索")
            print("    - 先固定 B11=0，二分搜索使 H11=0 的 margin")
            print("    - 再固定 margin，二分搜索使 F20=0 的 B11")
            print("    - 交替迭代直到收敛")
        else:
            print("\n>>> 建议策略: 梯度下降或智能采样")
            print("    - 使用少量采样点估计梯度方向")
            print("    - 沿梯度方向快速逼近目标")

        return {
            'coef_h11': coef_h11,
            'coef_f20': coef_f20,
            'r2_h11': r2_h11,
            'r2_f20': r2_f20,
            'h11_matrix': h11_matrix,
            'f20_matrix': f20_matrix,
            'margins': margins,
            'b11s': b11s,
        }

    finally:
        adjuster._unload_model()


if __name__ == '__main__':
    # 使用测试文件
    test_files = [
        '/Users/chenjiabin/Documents/demo/accounting-assistant/洪运来2511.xlsx',
        '/Users/chenjiabin/Documents/demo/accounting-assistant/data/洪运来2512.xlsx',
    ]

    for f in test_files:
        if os.path.exists(f):
            analyze_function_characteristics(f)
            break
    else:
        print("未找到测试文件，请指定文件路径")
