"""utils/date_fmt.py のユニットテスト"""


from utils.date_fmt import DATE_KEYS, format_date


class TestFormatDate:
    """format_date() のテスト"""

    # ── ISO 形式 (YYYY-MM-DD) ────────────────────────────────────────────
    def test_iso_date(self):
        assert format_date('2018-06-15') == '18/06/15'

    def test_iso_date_single_digit_month_day(self):
        assert format_date('2024-1-5') == '24/01/05'

    def test_iso_date_with_time(self):
        assert format_date('2018-06-15 00:00:00') == '18/06/15'

    # ── スラッシュ形式 (YYYY/MM/DD) ──────────────────────────────────────
    def test_slash_date(self):
        assert format_date('2018/06/15') == '18/06/15'

    def test_slash_date_single_digit(self):
        assert format_date('2024/3/1') == '24/03/01'

    # ── 2000 年前後 ─────────────────────────────────────────────────────
    def test_year_2000(self):
        assert format_date('2000-01-01') == '00/01/01'

    def test_year_1999(self):
        assert format_date('1999-12-31') == '99/12/31'

    # ── Excel シリアル値 ─────────────────────────────────────────────────
    def test_excel_serial_float_string(self):
        # 43266 = 2018-06-15 (Excel serial)
        assert format_date('43266.0') == '18/06/15'

    def test_excel_serial_integer_string(self):
        # 43266 as integer string
        assert format_date('43266') == '18/06/15'

    def test_excel_serial_recent_date(self):
        # 45658 = 2025-01-01
        assert format_date('45658.0') == '25/01/01'

    # ── 変換不能 → そのまま返す ──────────────────────────────────────────
    def test_empty_string(self):
        assert format_date('') == ''

    def test_non_date_string(self):
        assert format_date('abc') == 'abc'

    def test_partial_date(self):
        assert format_date('2024-13') == '2024-13'

    def test_zero_serial(self):
        """serial <= 1 は変換しない"""
        assert format_date('0') == '0'

    def test_huge_serial(self):
        """serial >= 100000 は変換しない"""
        assert format_date('100000') == '100000'


class TestDateKeys:
    """DATE_KEYS 定数のテスト"""

    def test_contains_birthday(self):
        assert '生年月日' in DATE_KEYS

    def test_contains_enrollment(self):
        assert '入学日' in DATE_KEYS

    def test_contains_transfer(self):
        assert '転入日' in DATE_KEYS

    def test_is_frozenset(self):
        assert isinstance(DATE_KEYS, frozenset)
