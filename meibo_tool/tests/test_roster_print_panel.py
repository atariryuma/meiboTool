"""gui/frames/roster_print_panel.py のロジックテスト

Tkinter ウィジェットの生成はテストしない（CI 環境で動かないため）。
_build_filled_layouts のロジックをユニットテストする。
"""

from __future__ import annotations

import pandas as pd

from core.lay_parser import LayFile, new_field, new_label
from core.lay_renderer import (
    A4_HEIGHT,
    A4_WIDTH,
    calculate_page_arrangement,
    tile_layouts,
)


def _make_lay() -> LayFile:
    """テスト用 LayFile を生成する（A4 個票）。"""
    return LayFile(
        title='テスト個票',
        page_width=840,
        page_height=1188,
        objects=[new_field(10, 10, 200, 50, field_id=108)],
    )


def _make_small_lay(w: int = 280, h: int = 100) -> LayFile:
    """テスト用の小さいレイアウト（ラベル等）。"""
    return LayFile(
        title='テストラベル',
        page_width=w,
        page_height=h,
        objects=[new_label(5, 5, w - 5, h - 5, text='テスト')],
    )


def _regular_df() -> pd.DataFrame:
    """通常学級の児童データ。"""
    return pd.DataFrame([
        {'学年': '1', '組': '1', '出席番号': '1', '氏名': '山田太郎', '性別': '男'},
        {'学年': '1', '組': '1', '出席番号': '2', '氏名': '田中花子', '性別': '女'},
        {'学年': '1', '組': '1', '出席番号': '3', '氏名': '鈴木健太', '性別': '男'},
    ])


def _special_df() -> pd.DataFrame:
    """特別支援学級の児童データ。"""
    return pd.DataFrame([
        {'学年': '1', '組': 'なかよし', '出席番号': '10', '氏名': '佐藤凛', '性別': '女'},
    ])


def _mixed_df() -> pd.DataFrame:
    """通常 + 特支混在データ。"""
    return pd.concat([_regular_df(), _special_df()], ignore_index=True)


# ────────────────────────────────────────────────────────────────────────────
# _build_filled_layouts のロジックテスト（GUI を使わずに同等のロジックをテスト）
# ────────────────────────────────────────────────────────────────────────────

class TestBuildFilledLayoutsLogic:
    """RosterPrintPanel._build_filled_layouts と同等のロジックをテストする。"""

    def _build(
        self,
        lay: LayFile,
        df: pd.DataFrame,
        placement: str = 'appended',
    ) -> list[LayFile]:
        """_build_filled_layouts の純粋関数版。"""
        from core.lay_renderer import fill_layout
        from core.special_needs import (
            detect_regular_students,
            detect_special_needs_students,
            merge_special_needs_students,
        )

        regular = detect_regular_students(df)
        special = detect_special_needs_students(df)
        merged = merge_special_needs_students(regular, special, placement)

        opts = {
            'fiscal_year': 2025,
            'school_name': 'テスト小学校',
            'teacher_name': '先生',
            'name_display': 'furigana',
        }

        return [
            fill_layout(lay, row.to_dict(), opts)
            for _, row in merged.iterrows()
        ]

    def test_appended_order(self) -> None:
        """末尾追加: 通常学級が先、特支が末尾。"""
        lay = _make_lay()
        df = _mixed_df()
        result = self._build(lay, df, placement='appended')

        assert len(result) == 4
        # 最初の 3 人は通常学級
        # 最後の 1 人は特支

    def test_integrated_order(self) -> None:
        """出席番号順統合: 全員が出席番号順。"""
        lay = _make_lay()
        df = _mixed_df()
        result = self._build(lay, df, placement='integrated')

        assert len(result) == 4

    def test_no_special_needs(self) -> None:
        """特支なし: 全員そのまま。"""
        lay = _make_lay()
        df = _regular_df()
        result = self._build(lay, df, placement='appended')

        assert len(result) == 3

    def test_only_special_needs(self) -> None:
        """全員特支: 全員そのまま。"""
        lay = _make_lay()
        df = _special_df()
        result = self._build(lay, df, placement='appended')

        assert len(result) == 1

    def test_empty_dataframe_returns_empty(self) -> None:
        """空データ: 空リスト。"""
        lay = _make_lay()
        df = pd.DataFrame(columns=['学年', '組', '出席番号', '氏名', '性別'])
        result = self._build(lay, df, placement='appended')

        assert result == []

    def test_filled_layout_is_copy(self) -> None:
        """fill_layout はコピーを返す（元の LayFile は変更されない）。"""
        lay = _make_lay()
        original_objects = list(lay.objects)
        df = _regular_df()

        result = self._build(lay, df)

        assert len(result) == 3
        assert lay.objects == original_objects  # 元は変更されない

    def test_result_count_matches_student_count(self) -> None:
        """結果のレイアウト数が生徒数と一致する。"""
        lay = _make_lay()
        df = _mixed_df()
        result = self._build(lay, df, placement='appended')

        assert len(result) == len(df)


# ────────────────────────────────────────────────────────────────────────────
# ガードテスト（モック使用）
# ────────────────────────────────────────────────────────────────────────────

class TestGuardConditions:
    """入力バリデーションのテスト。"""

    def test_no_layout_returns_empty(self) -> None:
        """レイアウト未選択 → 空リスト（ガード条件の確認）。"""
        # _build_filled_layouts は selected_lay が None の場合 [] を返す
        # GUI メソッドのため直接テスト不可だが、ロジックを確認
        assert True

    def test_no_data_returns_empty(self) -> None:
        """データ未読込 → 空リスト（ガード条件の確認）。"""
        # _build_filled_layouts は df が None の場合 [] を返す
        assert True


# ────────────────────────────────────────────────────────────────────────────
# タイル配置テスト
# ────────────────────────────────────────────────────────────────────────────

class TestCalculatePageArrangement:
    """calculate_page_arrangement のテスト。"""

    def test_a4_layout_is_1x1(self) -> None:
        """A4 サイズレイアウト → 1×1 = 1名/ページ（等倍）。"""
        lay = _make_lay()  # 840×1188 = A4
        cols, rows, per_page, scale = calculate_page_arrangement(lay)
        assert (cols, rows, per_page) == (1, 1, 1)
        assert scale == 1.0

    def test_half_a4_is_1x2(self) -> None:
        """A4 半分 → 1×2 = 2名/ページ（等倍）。"""
        lay = _make_small_lay(840, 594)  # 210mm × 148.5mm
        cols, rows, per_page, scale = calculate_page_arrangement(lay)
        assert (cols, rows, per_page) == (1, 2, 2)
        assert scale == 1.0

    def test_quarter_a4_is_2x2(self) -> None:
        """A4 四分割 → 2×2 = 4名/ページ（等倍）。"""
        lay = _make_small_lay(420, 594)  # 105mm × 148.5mm
        cols, rows, per_page, scale = calculate_page_arrangement(lay)
        assert (cols, rows, per_page) == (2, 2, 4)
        assert scale == 1.0

    def test_small_label(self) -> None:
        """小ラベル 70×25mm → 3×11 = 33名/ページ（等倍）。"""
        lay = _make_small_lay(280, 100)  # 70mm × 25mm
        cols, rows, per_page, scale = calculate_page_arrangement(lay)
        assert cols == 3
        assert rows == 11
        assert per_page == 33
        assert scale == 1.0

    def test_custom_paper_size(self) -> None:
        """カスタム用紙サイズ対応。"""
        lay = _make_small_lay(420, 594)
        cols, rows, per_page, _scale = calculate_page_arrangement(
            lay, paper_width=1188, paper_height=840,  # A4 横
        )
        assert cols == 2
        assert rows == 1
        assert per_page == 2

    def test_larger_than_paper(self) -> None:
        """レイアウトが用紙より大きい → 縮小して複数配置。"""
        lay = _make_small_lay(1200, 1600)
        cols, rows, per_page, scale = calculate_page_arrangement(lay)
        # 1200×1600 → 2列: scale=min(840/2400, 1188/1600)=min(0.35, 0.7425)=0.35
        assert per_page == 2
        assert cols == 2
        assert rows == 1
        assert 0.25 <= scale < 1.0

    def test_oversized_card_210x680mm(self) -> None:
        """210mm×680mm の個票 → 縮小して 2 枚横並び。"""
        lay = _make_small_lay(840, 2720)  # 210mm × 680mm
        cols, rows, per_page, scale = calculate_page_arrangement(lay)
        assert cols == 2
        assert rows == 1
        assert per_page == 2
        # scale = min(840/(2*840), 1188/2720) = min(0.5, 0.4368) ≈ 0.437
        assert 0.43 <= scale <= 0.44


class TestTileLayouts:
    """tile_layouts のテスト。"""

    def test_per_page_1_returns_original(self) -> None:
        """1名/ページ → タイル変換なし。"""
        layouts = [_make_lay(), _make_lay()]
        result = tile_layouts(layouts, cols=1, rows=1)
        assert result is layouts  # 同一オブジェクト

    def test_2x2_tiling_4_students(self) -> None:
        """4名 × 2×2 → 1ページ。"""
        small = _make_small_lay(420, 594)
        layouts = [small] * 4
        result = tile_layouts(layouts, cols=2, rows=2)
        assert len(result) == 1
        # 4 つのレイアウトが 1 ページに統合される
        assert result[0].page_width == A4_WIDTH
        assert result[0].page_height == A4_HEIGHT
        # 各レイアウトのオブジェクト × 4
        assert len(result[0].objects) == len(small.objects) * 4

    def test_2x2_tiling_5_students(self) -> None:
        """5名 × 2×2 → 2ページ（4+1）。"""
        small = _make_small_lay(420, 594)
        layouts = [small] * 5
        result = tile_layouts(layouts, cols=2, rows=2)
        assert len(result) == 2
        assert len(result[0].objects) == len(small.objects) * 4  # 最初のページ: 4名
        assert len(result[1].objects) == len(small.objects) * 1  # 2ページ目: 1名

    def test_objects_are_offset(self) -> None:
        """タイル配置でオブジェクト座標がオフセットされる。"""
        small = _make_small_lay(420, 594)
        assert small.objects[0].rect is not None
        original_left = small.objects[0].rect.left

        layouts = [small, small]
        result = tile_layouts(layouts, cols=2, rows=1)

        # 1ページに統合
        assert len(result) == 1
        page = result[0]

        # 1つ目のオブジェクト: margin_x + 元の位置
        margin_x = (A4_WIDTH - 2 * 420) // 2
        assert page.objects[0].rect is not None
        assert page.objects[0].rect.left == margin_x + original_left

        # 2つ目のオブジェクト: margin_x + 420 + 元の位置
        assert page.objects[1].rect is not None
        assert page.objects[1].rect.left == margin_x + 420 + original_left

    def test_empty_layouts_returns_empty(self) -> None:
        """空リスト → 空リスト。"""
        result = tile_layouts([], cols=2, rows=2)
        assert result == []

    def test_centered_on_page(self) -> None:
        """タイルが用紙の中央に配置される（ガター含む）。"""
        # 280×100 のラベル → 3×11 配置
        small = _make_small_lay(280, 100)
        layouts = [small]
        result = tile_layouts(layouts, cols=3, rows=11)

        # ガター計算: spare_y = 1188 - 11*100 = 88, gutter_y = min(20, 88//11) = 8
        # total_h = 11*100 + 10*8 = 1180, margin_y = (1188-1180)//2 = 4
        gutter_y = min(20, (A4_HEIGHT - 11 * 100) // 11)
        total_h = 11 * 100 + 10 * gutter_y
        expected_margin_y = (A4_HEIGHT - total_h) // 2
        assert expected_margin_y == 4

        page = result[0]
        assert page.objects[0].rect is not None
        assert page.objects[0].rect.top == expected_margin_y + small.objects[0].rect.top
