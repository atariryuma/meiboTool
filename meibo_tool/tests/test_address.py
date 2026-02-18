"""utils/address.py のユニットテスト"""

import pytest
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

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
