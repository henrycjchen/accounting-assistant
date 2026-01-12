# -*- coding: utf-8 -*-
"""
测试 helpers.py 中的工具函数
"""
import pytest
from unittest.mock import MagicMock
from modules.voucher.helpers import random_range, random_pick, set_wrap_border


class TestRandomRange:
    """测试 random_range 函数"""

    def test_random_range_floor_true_returns_integer(self):
        """floor=True 时返回整数"""
        for _ in range(100):
            result = random_range(1, 10, floor=True)
            assert isinstance(result, int)
            assert 1 <= result <= 10

    def test_random_range_floor_false_returns_float(self):
        """floor=False 时返回浮点数"""
        for _ in range(100):
            result = random_range(1.0, 10.0, floor=False)
            assert isinstance(result, float)
            assert 1.0 <= result <= 10.0

    def test_random_range_floor_false_has_three_decimal_places(self):
        """floor=False 时返回值最多3位小数"""
        for _ in range(100):
            result = random_range(0.0, 1.0, floor=False)
            decimal_str = str(result).split('.')
            if len(decimal_str) > 1:
                assert len(decimal_str[1]) <= 3

    def test_random_range_same_min_max_floor_true(self):
        """min等于max时，floor=True返回该值"""
        result = random_range(5, 5, floor=True)
        assert result == 5

    def test_random_range_same_min_max_floor_false(self):
        """min等于max时，floor=False返回该值"""
        result = random_range(5.0, 5.0, floor=False)
        assert result == 5.0

    def test_random_range_negative_values(self):
        """支持负数范围"""
        for _ in range(100):
            result = random_range(-10, -1, floor=True)
            assert isinstance(result, int)
            assert -10 <= result <= -1


class TestRandomPick:
    """测试 random_pick 函数"""

    def test_random_pick_returns_correct_count(self):
        """返回指定数量的元素"""
        arr = [1, 2, 3, 4, 5]
        result = random_pick(arr, 3)
        assert len(result) == 3

    def test_random_pick_no_duplicates(self):
        """返回的元素不重复"""
        arr = [1, 2, 3, 4, 5]
        result = random_pick(arr, 5)
        assert len(result) == len(set(result))

    def test_random_pick_original_unchanged(self):
        """不修改原数组"""
        arr = [1, 2, 3, 4, 5]
        original = arr.copy()
        random_pick(arr, 3)
        assert arr == original

    def test_random_pick_count_exceeds_length(self):
        """count超过数组长度时返回所有元素"""
        arr = [1, 2, 3]
        result = random_pick(arr, 10)
        assert len(result) == 3

    def test_random_pick_empty_array(self):
        """空数组返回空列表"""
        result = random_pick([], 5)
        assert result == []

    def test_random_pick_zero_count(self):
        """count为0时返回空列表"""
        arr = [1, 2, 3]
        result = random_pick(arr, 0)
        assert result == []

    def test_random_pick_all_elements_from_source(self):
        """所有返回元素都来自源数组"""
        arr = ['a', 'b', 'c', 'd', 'e']
        result = random_pick(arr, 3)
        for item in result:
            assert item in arr


class TestSetWrapBorder:
    """测试 set_wrap_border 函数"""

    def test_set_wrap_border_sets_border(self, mock_cell):
        """设置边框"""
        set_wrap_border(mock_cell)
        assert mock_cell.border is not None

    def test_set_wrap_border_sets_alignment(self, mock_cell):
        """设置对齐方式"""
        set_wrap_border(mock_cell)
        assert mock_cell.alignment is not None

    def test_set_wrap_border_with_real_cell(self):
        """使用真实单元格测试"""
        from openpyxl import Workbook

        wb = Workbook()
        ws = wb.active
        cell = ws['A1']

        set_wrap_border(cell)

        # 检查边框已设置（有 left, right, top, bottom 属性）
        assert cell.border.left is not None
        assert cell.border.right is not None
        assert cell.border.top is not None
        assert cell.border.bottom is not None
        # 检查对齐方式
        assert cell.alignment is not None
        assert cell.alignment.vertical == 'center'
        assert cell.alignment.horizontal == 'center'
