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


def test_optimized_search(file_path):
    """测试优化后的搜索算法"""

    print("=" * 60)
    print(f"测试文件: {os.path.basename(file_path)}")
    print("=" * 60)

    # 进度回调
    def progress_callback(progress, message):
        print(f"  [{progress:3d}%] {message}")

    adjuster = TaxAdjuster(file_path, progress_callback=progress_callback)

    # 测试优化后的算法
    print("\n>>> 测试优化后的算法 <<<")
    start_time = time.time()
    result = adjuster.calculate_inventory_margin_adjustment(max_solutions=5)
    elapsed = time.time() - start_time

    if 'error' in result:
        print(f"\n错误: {result['error']}")
        return

    print(f"\n耗时: {elapsed:.2f} 秒")

    if 'stats' in result:
        stats = result['stats']
        print(f"计算次数: {stats.get('iterations', 'N/A')}")
        print(f"收敛: {'是' if stats.get('converged') else '否'}")

    print(f"\n当前值:")
    current = result['current']
    print(f"  毛利率: {current['margin']:.5f}")
    print(f"  B11: {current['B11']:,.2f}")
    print(f"  H11: {current['H11']:,.2f}")
    print(f"  F20: {current['F20']:,.2f}")

    print(f"\n找到 {len(result['solutions'])} 个方案:")
    print("-" * 80)
    print(f"{'方案':<15} {'毛利率':>10} {'B11':>12} {'H11':>12} {'F20':>12} {'状态':<10}")
    print("-" * 80)

    for sol in result['solutions']:
        label = sol.get('label', '')
        h11_ok = '✓' if sol.get('h11_ok') else '✗'
        f20_ok = '✓' if sol.get('f20_ok') else '✗'
        status = f"H11:{h11_ok} F20:{f20_ok}"

        print(f"{label:<15} {sol['margin']:>10.5f} {sol['B11']:>12,.0f} "
              f"{sol['H11']:>12.2f} {sol['F20']:>12.2f} {status:<10}")

    print("-" * 80)

    # 找出最优解
    optimal = None
    for sol in result['solutions']:
        if '最优' in sol.get('label', ''):
            optimal = sol
            break

    if optimal:
        print(f"\n推荐方案: {optimal.get('label')}")
        print(f"  毛利率: {optimal['margin']:.5f}")
        print(f"  B11 (加工费): {optimal['B11']:,.2f}")
        print(f"  预期 H11: {optimal['H11']:.2f}")
        print(f"  预期 F20: {optimal['F20']:.2f}")

        # 检查是否达标
        h11_target_ok = abs(optimal['H11']) < 1.0
        f20_target_ok = abs(optimal['F20']) < 100
        print(f"\n达标情况:")
        print(f"  H11 目标 (|H11| < 1.0): {'达标 ✓' if h11_target_ok else '未达标 ✗'}")
        print(f"  F20 目标 (|F20| < 100): {'达标 ✓' if f20_target_ok else '未达标 ✗'}")

    return result


if __name__ == '__main__':
    # 使用测试文件
    test_files = [
        '/Users/chenjiabin/Documents/demo/accounting-assistant/洪运来2511.xlsx',
        '/Users/chenjiabin/Documents/demo/accounting-assistant/data/洪运来2512.xlsx',
    ]

    for f in test_files:
        if os.path.exists(f):
            test_optimized_search(f)
            break
    else:
        print("未找到测试文件")
