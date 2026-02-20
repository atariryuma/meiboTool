"""win_printer テスト

GDI 印刷エンジンのロジックをモックでテストする。
実際のプリンター出力はテストしない。
"""

from __future__ import annotations

from unittest.mock import MagicMock, patch

import pytest

from core.lay_parser import (
    LayFile,
    LayoutObject,
    ObjectType,
    new_field,
    new_label,
    new_line,
)
from core.lay_renderer import fill_layout, model_to_printer


def _make_layout(*objects: LayoutObject) -> LayFile:
    return LayFile(
        title='テスト',
        page_width=840,
        page_height=1188,
        objects=list(objects),
    )


# ── model_to_printer テスト ──────────────────────────────────────────────────


class TestModelToPrinter:
    """モデル座標→プリンタドット変換のテスト。"""

    def test_zero(self) -> None:
        assert model_to_printer(0, 300) == 0

    def test_a4_width_300dpi(self) -> None:
        # 840 × 0.25mm = 210mm → 300dpi: 210/25.4*300 ≈ 2480
        dots = model_to_printer(840, 300)
        assert 2478 <= dots <= 2481

    def test_a4_height_300dpi(self) -> None:
        # 1188 × 0.25mm = 297mm → 300dpi: 297/25.4*300 ≈ 3508
        dots = model_to_printer(1188, 300)
        assert 3506 <= dots <= 3509

    def test_600dpi_doubles(self) -> None:
        dots_300 = model_to_printer(100, 300)
        dots_600 = model_to_printer(100, 600)
        assert dots_600 == pytest.approx(dots_300 * 2, abs=1)


# ── enumerate_printers テスト ────────────────────────────────────────────────


class TestEnumeratePrinters:
    """プリンター列挙のテスト。"""

    def test_returns_list(self) -> None:
        from core.win_printer import enumerate_printers
        result = enumerate_printers()
        assert isinstance(result, list)

    def test_default_printer_returns_string_or_none(self) -> None:
        from core.win_printer import get_default_printer
        result = get_default_printer()
        assert result is None or isinstance(result, str)


# ── PrintJob テスト（モック）──────────────────────────────────────────────────


class TestPrintJobMock:
    """PrintJob のモックテスト。"""

    @pytest.fixture()
    def mock_win32(self) -> dict:
        """win32 モジュールをモックする。"""
        mock_dc = MagicMock()
        mock_dc.GetDeviceCaps.return_value = 300  # 300 DPI
        mock_dc.DrawText.return_value = (100, 50, (0, 0, 100, 50))

        mock_win32ui = MagicMock()
        mock_win32ui.CreateDC.return_value = mock_dc
        mock_win32ui.CreateFont.return_value = MagicMock()
        mock_win32ui.CreatePen.return_value = MagicMock()

        mock_win32print = MagicMock()
        mock_win32print.OpenPrinter.return_value = MagicMock()
        mock_win32print.GetPrinter.return_value = {'pDevMode': MagicMock()}

        return {
            'dc': mock_dc,
            'win32ui': mock_win32ui,
            'win32print': mock_win32print,
        }

    def test_print_label(self, mock_win32: dict) -> None:
        """LABEL オブジェクトが DrawText で描画される。"""
        lay = _make_layout(new_label(10, 20, 200, 50, text='テスト'))

        with patch.dict('core.win_printer.__dict__', {
            'win32ui': mock_win32['win32ui'],
            'win32print': mock_win32['win32print'],
            'win32con': MagicMock(),
            'HAS_WIN32': True,
        }):
            from core.win_printer import PrintJob
            job = PrintJob.__new__(PrintJob)
            job._printer_name = 'TestPrinter'
            job._dc = mock_win32['dc']
            job._dpi_x = 300
            job._dpi_y = 300
            job._started = True

            job.print_page(lay)

            mock_win32['dc'].StartPage.assert_called_once()
            mock_win32['dc'].EndPage.assert_called_once()
            mock_win32['dc'].DrawText.assert_called()

    def test_print_line(self, mock_win32: dict) -> None:
        """LINE オブジェクトが MoveTo/LineTo で描画される。"""
        lay = _make_layout(new_line(0, 0, 100, 100))

        with patch.dict('core.win_printer.__dict__', {
            'win32ui': mock_win32['win32ui'],
            'win32print': mock_win32['win32print'],
            'win32con': MagicMock(),
            'HAS_WIN32': True,
        }):
            from core.win_printer import PrintJob
            job = PrintJob.__new__(PrintJob)
            job._printer_name = 'TestPrinter'
            job._dc = mock_win32['dc']
            job._dpi_x = 300
            job._dpi_y = 300
            job._started = True

            job.print_page(lay)

            mock_win32['dc'].MoveTo.assert_called()
            mock_win32['dc'].LineTo.assert_called()

    def test_print_empty_layout(self, mock_win32: dict) -> None:
        """空レイアウトでもエラーにならない。"""
        lay = _make_layout()

        with patch.dict('core.win_printer.__dict__', {
            'win32ui': mock_win32['win32ui'],
            'win32print': mock_win32['win32print'],
            'win32con': MagicMock(),
            'HAS_WIN32': True,
        }):
            from core.win_printer import PrintJob
            job = PrintJob.__new__(PrintJob)
            job._printer_name = 'TestPrinter'
            job._dc = mock_win32['dc']
            job._dpi_x = 300
            job._dpi_y = 300
            job._started = True

            job.print_page(lay)
            mock_win32['dc'].StartPage.assert_called_once()
            mock_win32['dc'].EndPage.assert_called_once()

    def test_print_not_started_raises(self) -> None:
        """start() 前に print_page を呼ぶとエラー。"""
        from core.win_printer import HAS_WIN32
        if not HAS_WIN32:
            pytest.skip('pywin32 未インストール')

        from core.win_printer import PrintJob
        job = PrintJob.__new__(PrintJob)
        job._started = False
        job._dc = None

        with pytest.raises(RuntimeError, match='start'):
            job.print_page(_make_layout())

    def test_context_manager(self, mock_win32: dict) -> None:
        """with 文で正しくクリーンアップされる。"""
        with patch.dict('core.win_printer.__dict__', {
            'win32ui': mock_win32['win32ui'],
            'win32print': mock_win32['win32print'],
            'win32con': MagicMock(),
            'HAS_WIN32': True,
        }):
            from core.win_printer import PrintJob
            job = PrintJob.__new__(PrintJob)
            job._printer_name = 'TestPrinter'
            job._dc = mock_win32['dc']
            job._started = True

            with job:
                pass

            # end() が呼ばれて DC が削除される
            mock_win32['dc'].EndDoc.assert_called()
            mock_win32['dc'].DeleteDC.assert_called()
            assert job._dc is None
            assert job._started is False


# ── fill + print 統合テスト ──────────────────────────────────────────────────


class TestFillAndPrintIntegration:
    """fill_layout + PrintJob の統合テスト。"""

    def test_fill_then_print_no_fields(self) -> None:
        """fill_layout 後は FIELD がないので _render_field は呼ばれない。"""
        objects = [
            new_label(0, 0, 100, 30, text='氏名'),
            new_field(100, 0, 300, 30, field_id=108),
            new_line(0, 30, 300, 30),
        ]
        lay = _make_layout(*objects)
        data = {'氏名': '田中太郎'}
        filled = fill_layout(lay, data)

        # FIELD がないことを確認
        for obj in filled.objects:
            assert obj.obj_type != ObjectType.FIELD

    def test_filled_objects_have_text(self) -> None:
        """fill 後のオブジェクトにはテキストが入っている。"""
        obj = new_field(10, 20, 200, 50, field_id=108)
        lay = _make_layout(obj)
        filled = fill_layout(lay, {'氏名': '印刷テスト'})
        assert filled.objects[0].text == '印刷テスト'
