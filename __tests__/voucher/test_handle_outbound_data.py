# -*- coding: utf-8 -*-
"""
测试 handle_outbound_data.py 中的数据处理函数
"""
import pytest
from datetime import datetime
from modules.voucher.handle_outbound_data import safe_float, wash_data


class TestSafeFloat:
    """测试 safe_float 函数"""

    def test_safe_float_with_none(self):
        """None 返回 0"""
        assert safe_float(None) == 0

    def test_safe_float_with_int(self):
        """整数转换为浮点数"""
        assert safe_float(100) == 100.0
        assert safe_float(-50) == -50.0

    def test_safe_float_with_float(self):
        """浮点数保持不变"""
        assert safe_float(3.14) == 3.14
        assert safe_float(-2.5) == -2.5

    def test_safe_float_with_string_number(self):
        """数字字符串转换为浮点数"""
        assert safe_float('123') == 123.0
        assert safe_float('  456.78  ') == 456.78

    def test_safe_float_with_chinese_label(self):
        """包含中文标签的文本返回 0"""
        assert safe_float('合计: 1000') == 0
        assert safe_float('税项：500') == 0
        assert safe_float('求和: 2000') == 0

    def test_safe_float_with_invalid_string(self):
        """无效字符串返回 0"""
        assert safe_float('abc') == 0
        assert safe_float('') == 0

    def test_safe_float_with_other_types(self):
        """其他类型返回 0"""
        assert safe_float([1, 2, 3]) == 0
        assert safe_float({'a': 1}) == 0


class TestWashData:
    """测试 wash_data 函数"""

    def test_wash_data_empty_list(self):
        """空列表返回空结果"""
        result = wash_data([['header']])
        assert result['valid_data'] == []
        assert result['invalid_data'] == []

    def test_wash_data_skips_header(self):
        """跳过表头行"""
        data = [
            ['header1', 'header2'],
            []  # 空数据行
        ]
        result = wash_data(data)
        assert result['valid_data'] == []

    def test_wash_data_extracts_product_name(self):
        """正确解析产品名称"""
        data = [
            ['header'],
            [None, None, None, '123', None, '卖家', None, '买家',
             datetime(2024, 1, 15), None, None, '*钢材*圆钢-20mm',
             None, '吨', 100, None, 5000, None, 650]
        ]
        result = wash_data(data)
        assert len(result['valid_data']) == 1
        assert result['valid_data'][0]['product'] == '圆钢-20mm'
        assert result['valid_data'][0]['product_type'] == '钢材'

    def test_wash_data_filters_invalid_product_types(self):
        """过滤无效产品类型（机动车、劳务）"""
        data = [
            ['header'],
            [None, None, None, '123', None, '卖家', None, '买家',
             datetime(2024, 1, 15), None, None, '*机动车*汽车',
             None, '辆', 1, None, 50000, None, 6500],
            [None, None, None, '124', None, '卖家', None, '买家',
             datetime(2024, 1, 15), None, None, '*劳务*服务',
             None, '次', 1, None, 1000, None, 130],
        ]
        result = wash_data(data)
        assert len(result['valid_data']) == 0

    def test_wash_data_handles_reversed_invoices(self):
        """处理被红冲蓝字发票"""
        data = [
            ['header'],
            # 原始发票
            [None, None, None, '111', None, '卖家', None, '买家',
             datetime(2024, 1, 15), None, None, '*钢材*圆钢',
             None, '吨', 100, None, 5000, None, 650,
             None, None, None, None, None, None, None, ''],
            # 红冲记录 - 引用原发票
            [None, None, None, '222', None, '卖家', None, '买家',
             datetime(2024, 1, 16), None, None, '*钢材*圆钢',
             None, '吨', -100, None, -5000, None, -650,
             None, None, None, None, None, None, None, '被红冲蓝字111'],
        ]
        result = wash_data(data)
        # 原始发票应被过滤（因为被红冲），红冲记录也被过滤
        assert len(result['valid_data']) == 0

    def test_wash_data_date_parsing_datetime(self):
        """解析 datetime 类型日期"""
        date = datetime(2024, 1, 15, 10, 30, 0)
        data = [
            ['header'],
            [None, None, None, '123', None, '卖家', None, '买家',
             date, None, None, '*钢材*圆钢',
             None, '吨', 100, None, 5000, None, 650]
        ]
        result = wash_data(data)
        assert result['valid_data'][0]['date'] > 0
        # 验证时间被归零到当天0点
        expected_date = int(datetime(2024, 1, 15).timestamp())
        assert result['valid_data'][0]['date'] == expected_date

    def test_wash_data_date_parsing_string(self):
        """解析字符串格式日期"""
        data = [
            ['header'],
            [None, None, None, '123', None, '卖家', None, '买家',
             '2024-01-15', None, None, '*钢材*圆钢',
             None, '吨', 100, None, 5000, None, 650]
        ]
        result = wash_data(data)
        expected_date = int(datetime(2024, 1, 15).timestamp())
        assert result['valid_data'][0]['date'] == expected_date

    def test_wash_data_preserves_all_fields(self):
        """保留所有必要字段"""
        data = [
            ['header'],
            [None, None, None, '12345678', None, '销售公司', None, '购买公司',
             datetime(2024, 1, 15), None, None, '*钢材*圆钢',
             None, '吨', 100.5, None, 5000.0, None, 650.0,
             None, None, None, None, None, None, None, '备注信息']
        ]
        result = wash_data(data)
        item = result['valid_data'][0]

        assert item['code'] == '12345678'
        assert item['sell_company'] == '销售公司'
        assert item['buy_company'] == '购买公司'
        assert item['product'] == '圆钢'
        assert item['product_type'] == '钢材'
        assert item['unit'] == '吨'
        assert item['count'] == 100.5
        assert item['price'] == 5000.0
        assert item['tax'] == 650.0
        assert item['notes'] == '备注信息'

    def test_wash_data_handles_short_rows(self):
        """处理字段不足的行"""
        data = [
            ['header'],
            [None, None, None, '123', None, '卖家']  # 缺少后续字段
        ]
        result = wash_data(data)
        # 应该能处理而不崩溃
        assert 'valid_data' in result

    def test_wash_data_handles_empty_product_string(self):
        """处理空产品名称"""
        data = [
            ['header'],
            [None, None, None, '123', None, '卖家', None, '买家',
             datetime(2024, 1, 15), None, None, '',
             None, '吨', 100, None, 5000, None, 650]
        ]
        result = wash_data(data)
        item = result['valid_data'][0]
        assert item['product'] == ''
        assert item['product_type'] == ''
