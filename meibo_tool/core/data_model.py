"""編集可能データモデル — DataFrame ラッパー + Undo/Redo"""

from __future__ import annotations

from dataclasses import dataclass

import pandas as pd


@dataclass(frozen=True, slots=True)
class _EditOp:
    """単一セルの編集操作。"""
    row: int
    col: str
    old_value: str
    new_value: str


class EditableDataModel:
    """DataFrame をラップし、セル単位の編集・Undo/Redo を提供する。

    帳票生成時は ``get_df()`` で最新の DataFrame を取得する。
    """

    def __init__(self, df: pd.DataFrame) -> None:
        self._df = df.copy()
        self._undo_stack: list[_EditOp] = []
        self._redo_stack: list[_EditOp] = []
        self._modified = False

    # ── 参照 ────────────────────────────────────────────────────────────────

    @property
    def df(self) -> pd.DataFrame:
        """内部 DataFrame への直接参照。

        ``App.df_mapped`` をこの参照にバインドすることで、
        ``set_value()`` の結果が即座に反映される。
        """
        return self._df

    def get_df(self) -> pd.DataFrame:
        """現在の DataFrame のコピーを返す。"""
        return self._df.copy()

    @property
    def columns(self) -> list[str]:
        """カラム名リスト。"""
        return list(self._df.columns)

    def __len__(self) -> int:
        return len(self._df)

    def get_value(self, row: int, col: str) -> str:
        """指定セルの値を文字列で返す。"""
        val = self._df.at[row, col]
        if val is None or (isinstance(val, float) and pd.isna(val)):
            return ''
        return str(val)

    def is_modified(self) -> bool:
        """未保存の変更があるかどうか。"""
        return self._modified

    # ── 編集 ────────────────────────────────────────────────────────────────

    def set_value(self, row: int, col: str, value: str) -> None:
        """セル値を変更し、Undo スタックに積む。"""
        old = self.get_value(row, col)
        if old == value:
            return
        op = _EditOp(row=row, col=col, old_value=old, new_value=value)
        self._df.at[row, col] = value
        self._undo_stack.append(op)
        self._redo_stack.clear()
        self._modified = True

    def undo(self) -> _EditOp | None:
        """直前の操作を取り消す。取り消した操作を返す。"""
        if not self._undo_stack:
            return None
        op = self._undo_stack.pop()
        self._df.at[op.row, op.col] = op.old_value
        self._redo_stack.append(op)
        self._modified = bool(self._undo_stack)
        return op

    def redo(self) -> _EditOp | None:
        """Undo した操作をやり直す。やり直した操作を返す。"""
        if not self._redo_stack:
            return None
        op = self._redo_stack.pop()
        self._df.at[op.row, op.col] = op.new_value
        self._undo_stack.append(op)
        self._modified = True
        return op

    def can_undo(self) -> bool:
        """Undo 可能かどうか。"""
        return bool(self._undo_stack)

    def can_redo(self) -> bool:
        """Redo 可能かどうか。"""
        return bool(self._redo_stack)

    def reset_modified(self) -> None:
        """変更フラグをリセットする（保存後に呼ぶ）。"""
        self._modified = False
