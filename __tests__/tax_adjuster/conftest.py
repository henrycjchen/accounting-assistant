# -*- coding: utf-8 -*-
"""
税负调整模块测试专用 fixtures
"""
import pytest
from unittest.mock import MagicMock, patch
from modules.tax_adjuster.adjust_tax import TaxAdjuster


@pytest.fixture
def mock_tax_adjuster():
    """创建模拟的 TaxAdjuster 实例（绕过文件加载）"""
    with patch.object(TaxAdjuster, '__init__', lambda x, y: None):
        adjuster = TaxAdjuster(None)
        adjuster.file_path = 'test.xlsx'
        adjuster.wb_val = MagicMock()
        adjuster.wb_formula = MagicMock()
        adjuster.T5 = 100000
        adjuster.prev_profit = 50000
        return adjuster
