"""utils/wareki.py のユニットテスト"""

import pytest
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from utils.wareki import to_wareki, to_wareki_full, fiscal_year_to_wareki


class TestToWareki:
    def test_reiwa_2025(self):
        assert to_wareki(2025) == '令和7年'

    def test_reiwa_start(self):
        assert to_wareki(2019, 5, 1) == '令和1年'

    def test_heisei_last_day(self):
        assert to_wareki(2019, 4, 30) == '平成31年'

    def test_heisei_start(self):
        assert to_wareki(1989, 1, 8) == '平成1年'

    def test_showa_last_day(self):
        assert to_wareki(1989, 1, 7) == '昭和64年'

    def test_showa_1(self):
        assert to_wareki(1926, 12, 25) == '昭和1年'

    def test_reiwa_2019(self):
        assert to_wareki(2019) == '平成31年'  # 1月1日時点では平成


class TestToWarekiFull:
    def test_full_format(self):
        assert to_wareki_full(2025, 4, 1) == '令和7年4月1日'


class TestFiscalYearToWareki:
    def test_2025(self):
        assert fiscal_year_to_wareki(2025) == '令和7年度'

    def test_2019(self):
        # 2019年度 = 4月1日起算 → 平成31年度（5月1日以降でも年度は4月1日基準）
        assert fiscal_year_to_wareki(2019) == '平成31年度'
