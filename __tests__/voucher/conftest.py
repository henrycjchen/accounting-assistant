# -*- coding: utf-8 -*-
"""
凭证生成模块测试专用 fixtures
"""
import pytest
from datetime import datetime


@pytest.fixture
def sample_outbound_row():
    """出库发票示例行数据"""
    return [
        None, None, None,           # 0-2: 空列
        '12345678',                 # 3: 发票代码
        None,                       # 4: 空列
        '销售公司A',                # 5: 销售公司
        None,                       # 6: 空列
        '购买公司B',                # 7: 购买公司
        datetime(2024, 1, 15),      # 8: 日期
        None, None,                 # 9-10: 空列
        '*钢材*圆钢-20mm',          # 11: 产品名称
        None,                       # 12: 空列
        '吨',                       # 13: 单位
        100.5,                      # 14: 数量
        None,                       # 15: 空列
        5000.0,                     # 16: 价格
        None,                       # 17: 空列
        650.0,                      # 18: 税额
        None, None, None, None,     # 19-22: 空列
        None, None, None,           # 23-25: 空列
        '',                         # 26: 备注
    ]


@pytest.fixture
def sample_inbound_row():
    """入库发票示例行数据"""
    return [
        None, None, None,           # 0-2: 空列
        '87654321',                 # 3: 发票代码
        None,                       # 4: 空列
        '供应商C',                  # 5: 销售公司
        None,                       # 6: 空列
        '本公司',                   # 7: 购买公司
        datetime(2024, 1, 10),      # 8: 日期
        None, None,                 # 9-10: 空列
        '*原材料*铁矿石',           # 11: 产品名称
        '高品位',                   # 12: 规格
        'kg',                       # 13: 单位
        2000.0,                     # 14: 数量
        None,                       # 15: 空列
        3000.0,                     # 16: 价格
        None,                       # 17: 空列
        390.0,                      # 18: 税额
    ]


@pytest.fixture
def sample_valid_outbound_data():
    """清洗后的有效出库数据"""
    base_date = int(datetime(2024, 1, 15).timestamp())
    return [
        {
            'code': '12345678',
            'sell_company': '销售公司A',
            'buy_company': '购买公司B',
            'date': base_date,
            'product': '圆钢',
            'product_type': '钢材',
            'unit': '吨',
            'count': 100.5,
            'notes': '',
            'price': 5000.0,
            'tax': 650.0
        },
        {
            'code': '12345679',
            'sell_company': '销售公司A',
            'buy_company': '购买公司B',
            'date': base_date,
            'product': '方钢',
            'product_type': '钢材',
            'unit': '吨',
            'count': 50.0,
            'notes': '',
            'price': 4500.0,
            'tax': 585.0
        },
        {
            'code': '12345680',
            'sell_company': '销售公司A',
            'buy_company': '购买公司C',
            'date': base_date + 86400,  # 第二天
            'product': '螺纹钢',
            'product_type': '钢材',
            'unit': '吨',
            'count': 200.0,
            'notes': '',
            'price': 4000.0,
            'tax': 520.0
        }
    ]
