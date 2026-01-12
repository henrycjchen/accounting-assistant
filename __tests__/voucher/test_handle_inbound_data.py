# -*- coding: utf-8 -*-
"""
测试 handle_inbound_data.py 中的数据处理函数
"""
import pytest
from datetime import datetime
from modules.voucher.handle_inbound_data import safe_float, wash_data


class TestSafeFloat:
    """测试 safe_float 函数（入库数据版本）"""

    def test_safe_float_with_none(self):
        """None 返回 0"""
        assert safe_float(None) == 0

    def test_safe_float_with_int(self):
        """整数转换为浮点数"""
        assert safe_float(100) == 100.0

    def test_safe_float_with_float(self):
        """浮点数保持不变"""
        assert safe_float(3.14) == 3.14

    def test_safe_float_with_string_number(self):
        """数字字符串转换为浮点数"""
        assert safe_float('123') == 123.0
        assert safe_float('  456.78  ') == 456.78

    def test_safe_float_with_chinese_label(self):
        """包含中文标签的文本返回 0"""
        assert safe_float('合计: 1000') == 0
        assert safe_float('税项：500') == 0


class TestWashDataInbound:
    """测试 wash_data 函数（入库数据）"""

    def test_wash_data_empty_list(self):
        """空列表返回空结果"""
        result = wash_data([['header']])
        assert result['valid_data'] == []

    def test_wash_data_extracts_product_info(self):
        """正确解析产品信息"""
        data = [
            ['header'],
            [None, None, None, None, None, '供应商A', None, None,
             datetime(2024, 1, 10), None, None, '*原材料*铁矿石',
             '高品位', 'kg', 2000, None, 3000, None, 390]
        ]
        result = wash_data(data)
        assert len(result['valid_data']) == 1
        item = result['valid_data'][0]
        assert item['product'] == '铁矿石'
        assert item['product_type'] == '原材料'
        assert item['specification'] == '高品位'
        assert item['unit'] == 'kg'

    def test_wash_data_filters_invalid_product_types(self):
        """过滤无效产品类型（机动车、劳务）"""
        data = [
            ['header'],
            [None, None, None, None, None, '供应商', None, None,
             datetime(2024, 1, 10), None, None, '*机动车*汽车',
             '', '辆', 1, None, 50000, None, 6500],
        ]
        result = wash_data(data)
        assert len(result['valid_data']) == 0

    def test_wash_data_filters_zero_count(self):
        """过滤数量为0的项"""
        data = [
            ['header'],
            [None, None, None, None, None, '供应商', None, None,
             datetime(2024, 1, 10), None, None, '*原材料*铁矿石',
             '', 'kg', 0, None, 3000, None, 390],
        ]
        result = wash_data(data)
        assert len(result['valid_data']) == 0

    def test_wash_data_sorts_by_date(self):
        """按日期排序"""
        data = [
            ['header'],
            [None, None, None, None, None, '供应商B', None, None,
             datetime(2024, 1, 20), None, None, '*原材料*铜',
             '', 'kg', 100, None, 5000, None, 650],
            [None, None, None, None, None, '供应商A', None, None,
             datetime(2024, 1, 10), None, None, '*原材料*铁',
             '', 'kg', 200, None, 3000, None, 390],
            [None, None, None, None, None, '供应商C', None, None,
             datetime(2024, 1, 15), None, None, '*原材料*铝',
             '', 'kg', 150, None, 4000, None, 520],
        ]
        result = wash_data(data)
        assert len(result['valid_data']) == 3
        # 验证排序顺序
        dates = [item['date'] for item in result['valid_data']]
        assert dates == sorted(dates)

    def test_wash_data_date_parsing_datetime(self):
        """解析 datetime 类型日期"""
        date = datetime(2024, 1, 15, 10, 30, 0)
        data = [
            ['header'],
            [None, None, None, None, None, '供应商', None, None,
             date, None, None, '*原材料*铁矿石',
             '', 'kg', 100, None, 3000, None, 390]
        ]
        result = wash_data(data)
        assert result['valid_data'][0]['date'] > 0

    def test_wash_data_date_parsing_string(self):
        """解析字符串格式日期"""
        data = [
            ['header'],
            [None, None, None, None, None, '供应商', None, None,
             '2024-01-15', None, None, '*原材料*铁矿石',
             '', 'kg', 100, None, 3000, None, 390]
        ]
        result = wash_data(data)
        expected_date = int(datetime(2024, 1, 15).timestamp())
        assert result['valid_data'][0]['date'] == expected_date

    def test_wash_data_preserves_all_fields(self):
        """保留所有必要字段"""
        data = [
            ['header'],
            [None, None, None, None, None, '供应商公司', None, None,
             datetime(2024, 1, 15), None, None, '*原材料*铁矿石',
             '高品位A级', 'kg', 2000.5, None, 3000.0, None, 390.0]
        ]
        result = wash_data(data)
        item = result['valid_data'][0]

        assert item['sell_company'] == '供应商公司'
        assert item['product'] == '铁矿石'
        assert item['product_type'] == '原材料'
        assert item['specification'] == '高品位A级'
        assert item['unit'] == 'kg'
        assert item['count'] == 2000.5
        assert item['price'] == 3000.0
        assert item['tax'] == 390.0

    def test_wash_data_handles_negative_count(self):
        """负数数量被过滤（count > 0 条件）"""
        data = [
            ['header'],
            [None, None, None, None, None, '供应商', None, None,
             datetime(2024, 1, 10), None, None, '*原材料*铁矿石',
             '', 'kg', -100, None, 3000, None, 390],
        ]
        result = wash_data(data)
        assert len(result['valid_data']) == 0

    def test_wash_data_handles_empty_rows(self):
        """处理空行"""
        data = [
            ['header'],
            [],
            None,
            [None, None, None, None, None, '供应商', None, None,
             datetime(2024, 1, 10), None, None, '*原材料*铁矿石',
             '', 'kg', 100, None, 3000, None, 390],
        ]
        result = wash_data(data)
        assert len(result['valid_data']) == 1
