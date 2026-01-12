# -*- coding: utf-8 -*-
"""
测试 adjust_tax.py 中的税负调整核心逻辑
"""
import pytest
from unittest.mock import MagicMock, patch
from modules.tax_adjuster.adjust_tax import TaxAdjuster


class TestCalculateTax:
    """测试累进税率计算"""

    @pytest.fixture
    def mock_adjuster(self):
        """创建模拟的 TaxAdjuster 实例（绕过文件加载）"""
        with patch.object(TaxAdjuster, '__init__', lambda x, y: None):
            adjuster = TaxAdjuster(None)
            adjuster.file_path = 'test.xlsx'
            adjuster.wb_val = MagicMock()
            adjuster.wb_formula = MagicMock()
            adjuster.T5 = 100000
            adjuster.prev_profit = 50000
            return adjuster

    def test_calculate_tax_bracket_1(self, mock_adjuster):
        """第一档税率: 0-30000, 5%"""
        result = mock_adjuster.calculate_tax(10000)
        assert result == 10000 * 0.05  # 500

    def test_calculate_tax_bracket_1_boundary(self, mock_adjuster):
        """第一档边界: 30000"""
        result = mock_adjuster.calculate_tax(30000)
        assert result == 30000 * 0.05  # 1500

    def test_calculate_tax_bracket_2(self, mock_adjuster):
        """第二档税率: 30000-90000, 10% - 1500"""
        result = mock_adjuster.calculate_tax(60000)
        assert result == 60000 * 0.1 - 1500  # 4500

    def test_calculate_tax_bracket_2_boundary(self, mock_adjuster):
        """第二档边界: 90000"""
        result = mock_adjuster.calculate_tax(90000)
        assert result == 90000 * 0.1 - 1500  # 7500

    def test_calculate_tax_bracket_3(self, mock_adjuster):
        """第三档税率: 90000-300000, 20% - 10500"""
        result = mock_adjuster.calculate_tax(200000)
        assert result == 200000 * 0.2 - 10500  # 29500

    def test_calculate_tax_bracket_3_boundary(self, mock_adjuster):
        """第三档边界: 300000"""
        result = mock_adjuster.calculate_tax(300000)
        assert result == 300000 * 0.2 - 10500  # 49500

    def test_calculate_tax_bracket_4(self, mock_adjuster):
        """第四档税率: 300000-500000, 30% - 40500"""
        result = mock_adjuster.calculate_tax(400000)
        assert result == 400000 * 0.3 - 40500  # 79500

    def test_calculate_tax_bracket_4_boundary(self, mock_adjuster):
        """第四档边界: 500000"""
        result = mock_adjuster.calculate_tax(500000)
        assert result == 500000 * 0.3 - 40500  # 109500

    def test_calculate_tax_bracket_5(self, mock_adjuster):
        """第五档税率: >500000, 35% - 65500"""
        result = mock_adjuster.calculate_tax(1000000)
        assert result == 1000000 * 0.35 - 65500  # 284500

    def test_calculate_tax_zero_income(self, mock_adjuster):
        """零收入"""
        result = mock_adjuster.calculate_tax(0)
        assert result == 0


class TestReverseCalculateIncome:
    """测试根据税额反推应纳税所得额"""

    @pytest.fixture
    def mock_adjuster(self):
        """创建模拟的 TaxAdjuster 实例"""
        with patch.object(TaxAdjuster, '__init__', lambda x, y: None):
            adjuster = TaxAdjuster(None)
            return adjuster

    def test_reverse_bracket_1(self, mock_adjuster):
        """第一档反推: tax <= 1500"""
        tax = 500
        income = mock_adjuster.reverse_calculate_income(tax)
        assert income == tax / 0.05  # 10000

    def test_reverse_bracket_2(self, mock_adjuster):
        """第二档反推: 1500 < tax <= 7500"""
        tax = 4500
        income = mock_adjuster.reverse_calculate_income(tax)
        assert income == (tax + 1500) / 0.1  # 60000

    def test_reverse_bracket_3(self, mock_adjuster):
        """第三档反推: 7500 < tax <= 49500"""
        tax = 29500
        income = mock_adjuster.reverse_calculate_income(tax)
        assert income == (tax + 10500) / 0.2  # 200000

    def test_reverse_bracket_4(self, mock_adjuster):
        """第四档反推: 49500 < tax <= 109500"""
        tax = 79500
        income = mock_adjuster.reverse_calculate_income(tax)
        assert income == (tax + 40500) / 0.3  # 400000

    def test_reverse_bracket_5(self, mock_adjuster):
        """第五档反推: tax > 109500"""
        tax = 284500
        income = mock_adjuster.reverse_calculate_income(tax)
        assert income == (tax + 65500) / 0.35  # 1000000

    def test_calculate_and_reverse_consistency(self, mock_adjuster):
        """计算和反推的一致性"""
        test_incomes = [10000, 50000, 150000, 400000, 800000]
        for income in test_incomes:
            tax = mock_adjuster.calculate_tax(income)
            reversed_income = mock_adjuster.reverse_calculate_income(tax)
            assert abs(reversed_income - income) < 0.01


class TestExtractPrevProfit:
    """测试从公式中提取上期累计利润"""

    def test_extract_prev_profit_valid_formula(self):
        """有效公式提取"""
        with patch.object(TaxAdjuster, '__init__', lambda x, y: None):
            adjuster = TaxAdjuster(None)
            adjuster.wb_formula = MagicMock()
            adjuster.wb_formula.__getitem__ = MagicMock(return_value=MagicMock())
            adjuster.wb_formula['测算表']['E30'].value = '=50000.5+B46'

            result = adjuster._extract_prev_profit()
            assert result == 50000.5

    def test_extract_prev_profit_no_formula(self):
        """无公式返回0"""
        with patch.object(TaxAdjuster, '__init__', lambda x, y: None):
            adjuster = TaxAdjuster(None)
            adjuster.wb_formula = MagicMock()
            adjuster.wb_formula.__getitem__ = MagicMock(return_value=MagicMock())
            adjuster.wb_formula['测算表']['E30'].value = None

            result = adjuster._extract_prev_profit()
            assert result == 0

    def test_extract_prev_profit_number_value(self):
        """纯数值返回0"""
        with patch.object(TaxAdjuster, '__init__', lambda x, y: None):
            adjuster = TaxAdjuster(None)
            adjuster.wb_formula = MagicMock()
            adjuster.wb_formula.__getitem__ = MagicMock(return_value=MagicMock())
            adjuster.wb_formula['测算表']['E30'].value = 50000

            result = adjuster._extract_prev_profit()
            assert result == 0


class TestFindG25ForTargetB46:
    """测试二分法查找G25"""

    @pytest.fixture
    def mock_adjuster(self):
        """创建模拟的 TaxAdjuster 实例"""
        with patch.object(TaxAdjuster, '__init__', lambda x, y: None):
            adjuster = TaxAdjuster(None)
            adjuster.T5 = 100000

            # 模拟 calculate_B46_from_G25 方法
            # G25 越小，B46 越大（因为成本降低）
            def mock_calculate(g25):
                # 简化的模拟：B46 = 100000 - 50000 * g25
                b46 = 100000 - 50000 * g25
                j12 = 800000 * g25
                return b46, j12

            adjuster.calculate_B46_from_G25 = mock_calculate
            return adjuster

    def test_find_g25_returns_value_in_range(self, mock_adjuster):
        """返回值在有效范围内 (0.85, 1.00)"""
        target_B46 = 55000  # 目标利润
        result = mock_adjuster.find_G25_for_target_B46(target_B46)
        assert 0.85 <= result <= 1.00

    def test_find_g25_binary_search_converges(self, mock_adjuster):
        """二分查找收敛"""
        target_B46 = 52500  # 对应 G25 = 0.95
        result = mock_adjuster.find_G25_for_target_B46(target_B46)
        assert isinstance(result, float)
        # 验证收敛精度
        assert abs(result - 0.95) < 0.001


class TestGetCurrentData:
    """测试获取当前数据"""

    def test_get_current_data_returns_all_keys(self):
        """返回所有必要的键"""
        with patch.object(TaxAdjuster, '__init__', lambda x, y: None):
            adjuster = TaxAdjuster(None)
            adjuster.wb_val = MagicMock()

            mock_sheet = MagicMock()
            mock_sheet.__getitem__ = lambda self, key: MagicMock(value=0)
            adjuster.wb_val.__getitem__ = lambda self, key: mock_sheet

            result = adjuster.get_current_data()

            expected_keys = ['E17', 'E18', 'E21', 'E29', 'E30', 'E31', 'B46', 'G25', 'J12', 'B2']
            for key in expected_keys:
                assert key in result

    def test_get_current_data_default_g25(self):
        """G25 默认值为 1"""
        with patch.object(TaxAdjuster, '__init__', lambda x, y: None):
            adjuster = TaxAdjuster(None)
            adjuster.wb_val = MagicMock()

            mock_sheet = MagicMock()
            mock_sheet.__getitem__ = lambda self, key: MagicMock(value=None)
            adjuster.wb_val.__getitem__ = lambda self, key: mock_sheet

            result = adjuster.get_current_data()
            assert result['G25'] == 1


class TestCalculateAdjustment:
    """测试计算调整方案"""

    @pytest.fixture
    def mock_adjuster(self):
        """创建完整模拟的 TaxAdjuster 实例"""
        with patch.object(TaxAdjuster, '__init__', lambda x, y: None):
            adjuster = TaxAdjuster(None)
            adjuster.T5 = 100000
            adjuster.prev_profit = 50000
            adjuster.wb_val = MagicMock()

            # 模拟 get_current_data
            adjuster.get_current_data = MagicMock(return_value={
                'E17': 10000000,  # 年收入 1000万
                'E18': 500000,
                'E21': 100000,
                'E29': 400000,
                'E30': 450000,
                'E31': 50000,
                'B46': 50000,
                'G25': 1.0,
                'J12': 800000,
                'B2': 1000000,
            })

            # 模拟 calculate_B46_from_G25
            adjuster.calculate_B46_from_G25 = MagicMock(return_value=(45000, 850000))

            # 模拟 find_G25_for_target_B46
            adjuster.find_G25_for_target_B46 = MagicMock(return_value=0.95)

            return adjuster

    def test_calculate_adjustment_returns_complete_result(self, mock_adjuster):
        """返回完整的调整结果"""
        result = mock_adjuster.calculate_adjustment(0.00414)

        assert 'current' in result
        assert 'target' in result
        assert 'verify' in result

    def test_calculate_adjustment_target_contains_required_keys(self, mock_adjuster):
        """target 包含必要的键"""
        result = mock_adjuster.calculate_adjustment(0.00414)
        target = result['target']

        assert 'rate' in target
        assert 'E18' in target
        assert 'E21' in target
        assert 'G25' in target
        assert 'B46' in target

    def test_calculate_adjustment_verify_contains_required_keys(self, mock_adjuster):
        """verify 包含必要的键"""
        result = mock_adjuster.calculate_adjustment(0.00414)
        verify = result['verify']

        assert 'B46' in verify
        assert 'E30' in verify
        assert 'E31' in verify
        assert 'E21' in verify
        assert 'rate' in verify
        assert 'J12' in verify

    def test_calculate_adjustment_target_rate_matches_input(self, mock_adjuster):
        """目标税负率与输入匹配"""
        target_rate = 0.00414
        result = mock_adjuster.calculate_adjustment(target_rate)
        assert result['target']['rate'] == target_rate

    def test_calculate_adjustment_target_e21_calculation(self, mock_adjuster):
        """目标 E21 = E17 * target_rate"""
        target_rate = 0.00414
        result = mock_adjuster.calculate_adjustment(target_rate)
        expected_e21 = 10000000 * target_rate  # E17 * rate
        assert result['target']['E21'] == expected_e21


class TestApplyAdjustment:
    """测试应用调整"""

    @pytest.fixture
    def mock_adjuster(self):
        """创建模拟的 TaxAdjuster 实例"""
        with patch.object(TaxAdjuster, '__init__', lambda x, y: None):
            adjuster = TaxAdjuster(None)
            adjuster.file_path = '/tmp/test.xlsx'
            adjuster.wb_formula = MagicMock()

            mock_sheet = MagicMock()
            adjuster.wb_formula.__getitem__ = lambda self, key: mock_sheet

            return adjuster

    def test_apply_adjustment_sets_g25(self, mock_adjuster):
        """设置 G25 值"""
        with patch('os.path.exists', return_value=False):
            mock_adjuster.apply_adjustment(0.95, 400000)
            mock_adjuster.wb_formula['测算表'].__setitem__.assert_any_call('G25', 0.95)

    def test_apply_adjustment_sets_e18(self, mock_adjuster):
        """设置 E18 值"""
        with patch('os.path.exists', return_value=False):
            mock_adjuster.apply_adjustment(0.95, 400000)
            mock_adjuster.wb_formula['测算表'].__setitem__.assert_any_call('E18', 400000)

    def test_apply_adjustment_saves_file(self, mock_adjuster):
        """保存文件"""
        with patch('os.path.exists', return_value=False):
            mock_adjuster.apply_adjustment(0.95, 400000)
            mock_adjuster.wb_formula.save.assert_called_once()

    def test_apply_adjustment_returns_path(self, mock_adjuster):
        """返回保存路径"""
        with patch('os.path.exists', return_value=False):
            result = mock_adjuster.apply_adjustment(0.95, 400000)
            assert result == '/tmp/test.xlsx'

    def test_apply_adjustment_custom_output_path(self, mock_adjuster):
        """自定义输出路径"""
        with patch('os.path.exists', return_value=False):
            result = mock_adjuster.apply_adjustment(0.95, 400000, '/tmp/output.xlsx')
            assert result == '/tmp/output.xlsx'
