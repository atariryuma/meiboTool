"""utils/address.py のユニットテスト"""

from utils.address import build_address, build_guardian_address


class TestBuildAddress:
    def test_full_fields(self):
        row = {'都道府県': '沖縄県', '市区町村': '那覇市', '町番地': '天久1-2-3', '建物名': 'コーポ天久101'}
        assert build_address(row) == '沖縄県那覇市天久1-2-3コーポ天久101'

    def test_empty_building(self):
        row = {'都道府県': '沖縄県', '市区町村': '那覇市', '町番地': '天久1-2-3', '建物名': ''}
        assert build_address(row) == '沖縄県那覇市天久1-2-3'

    def test_nan_building(self):
        row = {'都道府県': '沖縄県', '市区町村': '那覇市', '町番地': '天久1-2-3', '建物名': 'nan'}
        assert build_address(row) == '沖縄県那覇市天久1-2-3'

    def test_none_value(self):
        row = {'都道府県': '沖縄県', '市区町村': '那覇市', '町番地': None, '建物名': None}
        assert build_address(row) == '沖縄県那覇市'

    def test_missing_keys(self):
        row = {'都道府県': '沖縄県'}
        assert build_address(row) == '沖縄県'


class TestBuildGuardianAddress:
    """保護者住所結合テスト。"""

    def test_suzuki_single_field(self):
        """スズキ校務: 単一フィールドをそのまま返す。"""
        row = {'保護者住所': '東京都千代田区永田町1-7-1'}
        assert build_guardian_address(row) == '東京都千代田区永田町1-7-1'

    def test_c4th_separate_fields(self):
        """C4th: 4分割フィールドを結合して返す。"""
        row = {
            '保護者都道府県': '沖縄県',
            '保護者市区町村': '那覇市',
            '保護者町番地': '泊1-2-3',
            '保護者建物名': 'マンション泊202',
        }
        assert build_guardian_address(row) == '沖縄県那覇市泊1-2-3マンション泊202'

    def test_same_as_student_returns_doujou(self):
        """児童住所と同一 → 「同上」。"""
        row = {
            '都道府県': '沖縄県', '市区町村': '那覇市',
            '町番地': '天久1-2-3', '建物名': '',
            '保護者都道府県': '沖縄県', '保護者市区町村': '那覇市',
            '保護者町番地': '天久1-2-3', '保護者建物名': '',
        }
        assert build_guardian_address(row) == '同上'

    def test_different_from_student(self):
        """児童住所と異なる → 保護者住所を返す。"""
        row = {
            '都道府県': '沖縄県', '市区町村': '那覇市',
            '町番地': '天久1-2-3', '建物名': '',
            '保護者都道府県': '沖縄県', '保護者市区町村': '浦添市',
            '保護者町番地': '牧港5-6-7', '保護者建物名': '',
        }
        assert build_guardian_address(row) == '沖縄県浦添市牧港5-6-7'

    def test_empty_guardian_fields(self):
        """保護者住所データがない場合は空文字。"""
        assert build_guardian_address({}) == ''

    def test_nan_single_field_falls_through(self):
        """スズキ校務フィールドが nan の場合は C4th フォールバック。"""
        row = {
            'nan': '',
            '保護者住所': 'nan',
            '保護者都道府県': '沖縄県',
            '保護者市区町村': '那覇市',
            '保護者町番地': '安里1-1-1',
            '保護者建物名': '',
        }
        assert build_guardian_address(row) == '沖縄県那覇市安里1-1-1'

    def test_suzuki_single_field_priority(self):
        """スズキ校務の単一フィールドが C4th 分割より優先される。"""
        row = {
            '保護者住所': '東京都新宿区西新宿2-8-1',
            '保護者都道府県': '沖縄県',
            '保護者市区町村': '那覇市',
            '保護者町番地': '泊1-2-3',
            '保護者建物名': '',
        }
        assert build_guardian_address(row) == '東京都新宿区西新宿2-8-1'
