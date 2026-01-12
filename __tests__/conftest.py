# -*- coding: utf-8 -*-
"""
Pytest 全局配置和公用 fixtures
"""
import pytest
import sys
import os
from unittest.mock import MagicMock

# 添加项目根目录到 Python 路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


@pytest.fixture
def mock_workbook():
    """模拟 openpyxl Workbook"""
    wb = MagicMock()
    ws = MagicMock()
    wb.create_sheet.return_value = ws
    wb.__getitem__ = MagicMock(return_value=ws)
    return wb


@pytest.fixture
def mock_cell():
    """模拟 openpyxl Cell"""
    cell = MagicMock()
    cell.border = None
    cell.alignment = None
    return cell
