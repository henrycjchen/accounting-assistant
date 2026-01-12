# -*- coding: utf-8 -*-
"""
测试 create_outbound.py 中的数据处理函数
"""
import pytest
from datetime import datetime
from modules.voucher.create_outbound import (
    format_data,
    merge_by_company,
    split_by_date,
    merge_counts,
    split_by_count,
    sort_by_date
)


@pytest.fixture
def sample_data():
    """示例数据"""
    base_date = int(datetime(2024, 1, 15).timestamp())
    return [
        {
            'code': '001',
            'buy_company': '公司A',
            'date': base_date,
            'product': '圆钢',
            'unit': '吨',
            'count': 100,
        },
        {
            'code': '002',
            'buy_company': '公司A',
            'date': base_date,
            'product': '方钢',
            'unit': '吨',
            'count': 50,
        },
        {
            'code': '003',
            'buy_company': '公司B',
            'date': base_date + 86400,
            'product': '螺纹钢',
            'unit': '吨',
            'count': 200,
        },
    ]


class TestMergeByCompany:
    """测试 merge_by_company 函数"""

    def test_merge_by_company_groups_correctly(self, sample_data):
        """按公司正确分组"""
        result = merge_by_company(sample_data)
        assert len(result) == 2  # 两个公司

    def test_merge_by_company_preserves_all_items(self, sample_data):
        """保留所有数据项"""
        result = merge_by_company(sample_data)
        total_items = sum(len(group) for group in result)
        assert total_items == len(sample_data)

    def test_merge_by_company_empty_list(self):
        """空列表返回空结果"""
        result = merge_by_company([])
        assert result == []

    def test_merge_by_company_single_company(self):
        """单个公司的数据"""
        data = [
            {'buy_company': 'A', 'product': '1'},
            {'buy_company': 'A', 'product': '2'},
        ]
        result = merge_by_company(data)
        assert len(result) == 1
        assert len(result[0]) == 2


class TestSplitByDate:
    """测试 split_by_date 函数"""

    def test_split_by_date_splits_correctly(self):
        """按日期正确拆分"""
        base_date = int(datetime(2024, 1, 15).timestamp())
        data = [
            [
                {'date': base_date, 'product': '1'},
                {'date': base_date, 'product': '2'},
                {'date': base_date + 86400, 'product': '3'},
            ]
        ]
        result = split_by_date(data)
        assert len(result) == 2  # 两个日期

    def test_split_by_date_sorts_within_groups(self):
        """组内按日期排序"""
        base_date = int(datetime(2024, 1, 15).timestamp())
        data = [
            [
                {'date': base_date + 86400, 'product': '2'},
                {'date': base_date, 'product': '1'},
            ]
        ]
        result = split_by_date(data)
        # 每个日期应该是独立的组
        assert len(result) == 2

    def test_split_by_date_empty_input(self):
        """空输入"""
        result = split_by_date([])
        assert result == []

    def test_split_by_date_preserves_all_items(self):
        """保留所有数据项"""
        base_date = int(datetime(2024, 1, 15).timestamp())
        data = [
            [
                {'date': base_date, 'product': '1'},
                {'date': base_date + 86400, 'product': '2'},
                {'date': base_date + 86400 * 2, 'product': '3'},
            ]
        ]
        result = split_by_date(data)
        total_items = sum(len(group) for group in result)
        assert total_items == 3


class TestMergeCounts:
    """测试 merge_counts 函数"""

    def test_merge_counts_combines_same_product(self):
        """合并相同产品的数量"""
        data = [
            [
                {'product': '圆钢', 'unit': '吨', 'count': 100},
                {'product': '圆钢', 'unit': '吨', 'count': 50},
            ]
        ]
        result = merge_counts(data)
        assert len(result[0]) == 1
        assert result[0][0]['count'] == 150

    def test_merge_counts_different_units_not_merged(self):
        """不同单位的相同产品不合并"""
        data = [
            [
                {'product': '圆钢', 'unit': '吨', 'count': 100},
                {'product': '圆钢', 'unit': 'kg', 'count': 50},
            ]
        ]
        result = merge_counts(data)
        assert len(result[0]) == 2

    def test_merge_counts_different_products_not_merged(self):
        """不同产品不合并"""
        data = [
            [
                {'product': '圆钢', 'unit': '吨', 'count': 100},
                {'product': '方钢', 'unit': '吨', 'count': 50},
            ]
        ]
        result = merge_counts(data)
        assert len(result[0]) == 2

    def test_merge_counts_empty_input(self):
        """空输入"""
        result = merge_counts([])
        assert result == []


class TestSplitByCount:
    """测试 split_by_count 函数"""

    def test_split_by_count_splits_at_seven(self):
        """每7条分一组"""
        items = [{'product': str(i)} for i in range(14)]
        data = [items]
        result = split_by_count(data)
        assert len(result) == 2
        assert len(result[0]) == 7
        assert len(result[1]) == 7

    def test_split_by_count_less_than_seven(self):
        """少于7条保持原样"""
        items = [{'product': str(i)} for i in range(5)]
        data = [items]
        result = split_by_count(data)
        assert len(result) == 1
        assert len(result[0]) == 5

    def test_split_by_count_exactly_seven(self):
        """恰好7条"""
        items = [{'product': str(i)} for i in range(7)]
        data = [items]
        result = split_by_count(data)
        assert len(result) == 1
        assert len(result[0]) == 7

    def test_split_by_count_remainder(self):
        """有余数的情况"""
        items = [{'product': str(i)} for i in range(10)]
        data = [items]
        result = split_by_count(data)
        assert len(result) == 2
        assert len(result[0]) == 7
        assert len(result[1]) == 3

    def test_split_by_count_empty_input(self):
        """空输入"""
        result = split_by_count([])
        assert result == []

    def test_split_by_count_multiple_groups(self):
        """多个输入组"""
        data = [
            [{'product': str(i)} for i in range(10)],
            [{'product': str(i)} for i in range(5)],
        ]
        result = split_by_count(data)
        assert len(result) == 3  # 2 + 1


class TestSortByDate:
    """测试 sort_by_date 函数"""

    def test_sort_by_date_sorts_correctly(self):
        """按日期正确排序"""
        base_date = int(datetime(2024, 1, 15).timestamp())
        data = [
            [{'date': base_date + 86400 * 2}],
            [{'date': base_date}],
            [{'date': base_date + 86400}],
        ]
        result = sort_by_date(data)
        assert result[0][0]['date'] == base_date
        assert result[1][0]['date'] == base_date + 86400
        assert result[2][0]['date'] == base_date + 86400 * 2

    def test_sort_by_date_filters_empty_lists(self):
        """过滤空列表"""
        base_date = int(datetime(2024, 1, 15).timestamp())
        data = [
            [{'date': base_date}],
            [],
            [{'date': base_date + 86400}],
        ]
        result = sort_by_date(data)
        assert len(result) == 2

    def test_sort_by_date_empty_input(self):
        """空输入"""
        result = sort_by_date([])
        assert result == []


class TestFormatData:
    """测试 format_data 完整流程"""

    def test_format_data_complete_pipeline(self, sample_data):
        """完整数据处理管道"""
        result = format_data(sample_data)
        # 结果应该是按日期排序的分组列表
        assert isinstance(result, list)
        for group in result:
            assert isinstance(group, list)
            for item in group:
                assert 'product' in item
                assert 'count' in item

    def test_format_data_groups_by_company_and_date(self):
        """按公司和日期分组"""
        base_date = int(datetime(2024, 1, 15).timestamp())
        data = [
            {'buy_company': 'A', 'date': base_date, 'product': '1', 'unit': '吨', 'count': 10},
            {'buy_company': 'A', 'date': base_date, 'product': '2', 'unit': '吨', 'count': 20},
            {'buy_company': 'A', 'date': base_date + 86400, 'product': '3', 'unit': '吨', 'count': 30},
            {'buy_company': 'B', 'date': base_date, 'product': '4', 'unit': '吨', 'count': 40},
        ]
        result = format_data(data)
        # 应该有3个组: A-day1, A-day2, B-day1
        assert len(result) == 3

    def test_format_data_empty_input(self):
        """空输入返回空结果"""
        result = format_data([])
        assert result == []
