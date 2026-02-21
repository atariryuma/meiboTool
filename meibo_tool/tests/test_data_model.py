"""core/data_model.py のテスト

テスト対象:
  - EditableDataModel: セル編集、Undo/Redo、変更フラグ
"""

from __future__ import annotations

import pandas as pd

from core.data_model import EditableDataModel


def _make_model() -> EditableDataModel:
    """テスト用の 3 行モデルを作成する。"""
    df = pd.DataFrame({
        '氏名': ['山田太郎', '田中花子', '鈴木一郎'],
        '性別': ['男', '女', '男'],
        '学年': ['1', '2', '3'],
    })
    return EditableDataModel(df)


class TestEditableDataModelBasic:
    def test_len(self):
        model = _make_model()
        assert len(model) == 3

    def test_columns(self):
        model = _make_model()
        assert model.columns == ['氏名', '性別', '学年']

    def test_get_value(self):
        model = _make_model()
        assert model.get_value(0, '氏名') == '山田太郎'
        assert model.get_value(1, '性別') == '女'

    def test_get_value_nan_returns_empty(self):
        df = pd.DataFrame({'name': [None, float('nan')]})
        model = EditableDataModel(df)
        assert model.get_value(0, 'name') == ''
        assert model.get_value(1, 'name') == ''

    def test_get_df_returns_copy(self):
        model = _make_model()
        df = model.get_df()
        df.at[0, '氏名'] = 'changed'
        # 元のモデルは変更されない
        assert model.get_value(0, '氏名') == '山田太郎'

    def test_initial_not_modified(self):
        model = _make_model()
        assert not model.is_modified()


class TestSetValue:
    def test_basic_set(self):
        model = _make_model()
        model.set_value(0, '氏名', '佐藤太郎')
        assert model.get_value(0, '氏名') == '佐藤太郎'

    def test_set_marks_modified(self):
        model = _make_model()
        model.set_value(0, '氏名', '佐藤太郎')
        assert model.is_modified()

    def test_set_same_value_no_op(self):
        model = _make_model()
        model.set_value(0, '氏名', '山田太郎')  # 同じ値
        assert not model.is_modified()
        assert not model.can_undo()

    def test_set_clears_redo(self):
        model = _make_model()
        model.set_value(0, '氏名', '佐藤太郎')
        model.undo()
        assert model.can_redo()
        model.set_value(0, '氏名', '新しい名前')
        assert not model.can_redo()


class TestUndo:
    def test_undo_restores_value(self):
        model = _make_model()
        model.set_value(0, '氏名', '佐藤太郎')
        op = model.undo()
        assert model.get_value(0, '氏名') == '山田太郎'
        assert op is not None
        assert op.old_value == '山田太郎'
        assert op.new_value == '佐藤太郎'

    def test_undo_empty_returns_none(self):
        model = _make_model()
        assert model.undo() is None

    def test_undo_clears_modified(self):
        model = _make_model()
        model.set_value(0, '氏名', '佐藤太郎')
        model.undo()
        assert not model.is_modified()

    def test_multiple_undo(self):
        model = _make_model()
        model.set_value(0, '氏名', 'A')
        model.set_value(0, '氏名', 'B')
        model.undo()
        assert model.get_value(0, '氏名') == 'A'
        model.undo()
        assert model.get_value(0, '氏名') == '山田太郎'

    def test_can_undo(self):
        model = _make_model()
        assert not model.can_undo()
        model.set_value(0, '氏名', 'X')
        assert model.can_undo()
        model.undo()
        assert not model.can_undo()


class TestRedo:
    def test_redo_restores_value(self):
        model = _make_model()
        model.set_value(0, '氏名', '佐藤太郎')
        model.undo()
        op = model.redo()
        assert model.get_value(0, '氏名') == '佐藤太郎'
        assert op is not None
        assert op.new_value == '佐藤太郎'

    def test_redo_empty_returns_none(self):
        model = _make_model()
        assert model.redo() is None

    def test_redo_marks_modified(self):
        model = _make_model()
        model.set_value(0, '氏名', '佐藤太郎')
        model.undo()
        assert not model.is_modified()
        model.redo()
        assert model.is_modified()

    def test_can_redo(self):
        model = _make_model()
        assert not model.can_redo()
        model.set_value(0, '氏名', 'X')
        assert not model.can_redo()
        model.undo()
        assert model.can_redo()
        model.redo()
        assert not model.can_redo()


class TestResetModified:
    def test_reset_clears_flag(self):
        model = _make_model()
        model.set_value(0, '氏名', '佐藤太郎')
        assert model.is_modified()
        model.reset_modified()
        assert not model.is_modified()

    def test_undo_redo_still_work_after_reset(self):
        model = _make_model()
        model.set_value(0, '氏名', '佐藤太郎')
        model.reset_modified()
        assert model.can_undo()
        model.undo()
        assert model.get_value(0, '氏名') == '山田太郎'


class TestGetDfReflectsEdits:
    def test_get_df_after_edit(self):
        model = _make_model()
        model.set_value(1, '性別', '他')
        df = model.get_df()
        assert df.at[1, '性別'] == '他'

    def test_get_df_after_undo(self):
        model = _make_model()
        model.set_value(1, '性別', '他')
        model.undo()
        df = model.get_df()
        assert df.at[1, '性別'] == '女'
