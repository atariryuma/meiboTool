"""win_printer テスト

PIL 画像ベース印刷エンジンのロジックをモックでテストする。
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
        mock_dc.GetSafeHdc.return_value = 99999

        mock_win32ui = MagicMock()
        mock_win32ui.CreateDC.return_value = mock_dc
        mock_win32ui.CreateDCFromHandle.return_value = mock_dc

        mock_win32print = MagicMock()

        mock_win32gui = MagicMock()
        mock_win32gui.CreateDC.return_value = 12345

        mock_win32con = MagicMock()
        mock_win32con.HORZRES = 8
        mock_win32con.VERTRES = 10
        mock_win32con.LOGPIXELSX = 88
        mock_win32con.LOGPIXELSY = 90

        return {
            'dc': mock_dc,
            'win32ui': mock_win32ui,
            'win32print': mock_win32print,
            'win32gui': mock_win32gui,
            'win32con': mock_win32con,
        }

    def test_print_page_calls_render_and_blit(self, mock_win32: dict) -> None:
        """print_page が render_layout_to_image → _blit_pil_image を呼ぶ。"""
        lay = _make_layout(new_label(10, 20, 200, 50, text='テスト'))

        mock_img = MagicMock()
        mock_img.convert.return_value = mock_img
        mock_img.size = (2480, 3508)
        mock_img.tobytes.return_value = b'\x00' * (2480 * 3508 * 3)

        with (
            patch.dict('core.win_printer.__dict__', {
                'win32ui': mock_win32['win32ui'],
                'win32print': mock_win32['win32print'],
                'win32gui': mock_win32['win32gui'],
                'win32con': mock_win32['win32con'],
                'HAS_WIN32': True,
            }),
            patch('core.win_printer.render_layout_to_image', return_value=mock_img) as mock_render,
            patch('core.win_printer._blit_pil_image') as mock_blit,
        ):
            from core.win_printer import PrintJob
            job = PrintJob.__new__(PrintJob)
            job._printer_name = 'TestPrinter'
            job._dc = mock_win32['dc']
            job._hdc = None
            job._dpi_x = 300
            job._dpi_y = 300
            job._started = True

            job.print_page(lay)

            mock_win32['dc'].StartPage.assert_called_once()
            mock_win32['dc'].EndPage.assert_called_once()
            mock_render.assert_called_once_with(lay, dpi=300, for_print=True)
            mock_blit.assert_called_once_with(
                mock_win32['dc'], mock_img, 300, 300,
            )

    def test_print_empty_layout(self, mock_win32: dict) -> None:
        """空レイアウトでもエラーにならない。"""
        lay = _make_layout()
        mock_img = MagicMock()

        with (
            patch.dict('core.win_printer.__dict__', {
                'win32ui': mock_win32['win32ui'],
                'win32print': mock_win32['win32print'],
                'win32gui': mock_win32['win32gui'],
                'win32con': mock_win32['win32con'],
                'HAS_WIN32': True,
            }),
            patch('core.win_printer.render_layout_to_image', return_value=mock_img),
            patch('core.win_printer._blit_pil_image'),
        ):
            from core.win_printer import PrintJob
            job = PrintJob.__new__(PrintJob)
            job._printer_name = 'TestPrinter'
            job._dc = mock_win32['dc']
            job._hdc = None
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
        job._hdc = None

        with pytest.raises(RuntimeError, match='start'):
            job.print_page(_make_layout())

    def test_context_manager(self, mock_win32: dict) -> None:
        """with 文で正しくクリーンアップされる。"""
        with patch.dict('core.win_printer.__dict__', {
            'win32ui': mock_win32['win32ui'],
            'win32print': mock_win32['win32print'],
            'win32gui': mock_win32['win32gui'],
            'win32con': mock_win32['win32con'],
            'HAS_WIN32': True,
        }):
            from core.win_printer import PrintJob
            job = PrintJob.__new__(PrintJob)
            job._printer_name = 'TestPrinter'
            job._dc = mock_win32['dc']
            job._hdc = None
            job._started = True

            with job:
                pass

            # end() が呼ばれて DC が削除される
            mock_win32['dc'].EndDoc.assert_called()
            mock_win32['dc'].DeleteDC.assert_called()
            assert job._dc is None
            assert job._started is False

    def test_context_manager_with_hdc(self, mock_win32: dict) -> None:
        """hdc がある場合、win32gui.DeleteDC も呼ばれる。"""
        with patch.dict('core.win_printer.__dict__', {
            'win32ui': mock_win32['win32ui'],
            'win32print': mock_win32['win32print'],
            'win32gui': mock_win32['win32gui'],
            'win32con': mock_win32['win32con'],
            'HAS_WIN32': True,
        }):
            from core.win_printer import PrintJob
            job = PrintJob.__new__(PrintJob)
            job._printer_name = 'TestPrinter'
            job._dc = mock_win32['dc']
            job._hdc = 12345
            job._started = True

            with job:
                pass

            mock_win32['dc'].DeleteDC.assert_called()
            mock_win32['win32gui'].DeleteDC.assert_called_once_with(12345)
            assert job._hdc is None


# ── _blit_pil_image テスト ──────────────────────────────────────────────────


class TestBlitPilImage:
    """_blit_pil_image のテスト。"""

    def test_calls_stretch_dib_bits(self) -> None:
        """StretchDIBits が正しく呼ばれる。"""
        from core.win_printer import HAS_WIN32
        if not HAS_WIN32:
            pytest.skip('pywin32 未インストール')

        from PIL import Image

        from core.win_printer import _blit_pil_image

        # 10x10 白画像
        img = Image.new('RGB', (10, 10), (255, 255, 255))

        mock_dc = MagicMock()
        mock_dc.GetSafeHdc.return_value = 99999
        mock_dc.GetDeviceCaps.return_value = 2480

        with patch('core.win_printer.ctypes') as mock_ctypes:
            mock_gdi32 = MagicMock()
            mock_ctypes.windll.gdi32 = mock_gdi32

            _blit_pil_image(mock_dc, img, 300, 300)

            mock_gdi32.StretchDIBits.assert_called_once()
            args = mock_gdi32.StretchDIBits.call_args[0]
            assert args[0] == 99999  # hdc
            assert args[1] == 0      # xDest
            assert args[2] == 0      # yDest
            assert args[5] == 0      # xSrc
            assert args[6] == 0      # ySrc
            assert args[7] == 10     # wSrc
            assert args[8] == 10     # hSrc

    def test_stride_padding(self) -> None:
        """幅が 4 バイト境界でない場合パディングされる。"""
        from core.win_printer import HAS_WIN32
        if not HAS_WIN32:
            pytest.skip('pywin32 未インストール')

        from PIL import Image

        from core.win_printer import _blit_pil_image

        # 幅 5px → stride = (5*3+3)&~3 = 16 (padding=1)
        img = Image.new('RGB', (5, 2), (255, 0, 0))

        mock_dc = MagicMock()
        mock_dc.GetSafeHdc.return_value = 99999
        mock_dc.GetDeviceCaps.return_value = 100

        with patch('core.win_printer.ctypes') as mock_ctypes:
            mock_gdi32 = MagicMock()
            mock_ctypes.windll.gdi32 = mock_gdi32

            _blit_pil_image(mock_dc, img, 300, 300)

            mock_gdi32.StretchDIBits.assert_called_once()
            args = mock_gdi32.StretchDIBits.call_args[0]
            bits = args[9]
            # stride=16, 2行 → 32バイト
            assert len(bits) == 32


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
