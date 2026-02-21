"""Windows GDI 印刷エンジン

LayFile を PILBackend で高解像度画像にレンダリングし、
GDI StretchDIBits でプリンター DC に転送する方式を採用。
プレビュー (PILBackend) と印刷結果が同一コードパスで生成されるため、
表示と印刷が完全に一致する。

DC 作成は win32gui.CreateDC("WINSPOOL", ..., devmode) で DevMode を直接適用し、
win32ui.CreateDCFromHandle() で PyCDC にラップする方式を採用。
これにより用紙サイズ・印刷方向が確実に反映される。

条件付きインポートにより、Windows 以外の環境でもインポートエラーにならない。
"""

from __future__ import annotations

import contextlib
import ctypes
import struct
from typing import TYPE_CHECKING

from core.lay_renderer import render_layout_to_image

if TYPE_CHECKING:
    from core.lay_parser import LayFile

# ── 条件付きインポート ────────────────────────────────────────────────────────

try:
    import pywintypes
    import win32con
    import win32gui
    import win32print
    import win32ui

    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False


# ── プリンター列挙 ───────────────────────────────────────────────────────────


def enumerate_printers() -> list[dict]:
    """利用可能プリンター一覧を返す。

    Returns:
        [{'name': 'Microsoft Print to PDF', 'default': True}, ...]
    """
    if not HAS_WIN32:
        return []

    default_name = ''
    with contextlib.suppress(Exception):
        default_name = win32print.GetDefaultPrinter()

    printers = []
    flags = (
        win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    )
    for _flags, _desc, name, _comment in win32print.EnumPrinters(flags):
        printers.append({
            'name': name,
            'default': name == default_name,
        })
    return printers


def get_default_printer() -> str | None:
    """デフォルトプリンター名を返す。"""
    if not HAS_WIN32:
        return None
    with contextlib.suppress(Exception):
        return win32print.GetDefaultPrinter()
    return None


# ── DevMode 作成 ──────────────────────────────────────────────────────────────


def _create_devmode(printer_name: str) -> object | None:
    """ドライバプライベートデータ付きの DevMode を正しく作成する。

    DocumentProperties で完全なバッファサイズを取得し、ドライバ固有データも含めた
    DevMode を構築する。GetPrinter(2) より信頼性が高い。
    """
    hprinter = win32print.OpenPrinter(printer_name)
    try:
        # ドライバが必要とする DevMode バッファサイズを取得
        dmsize = win32print.DocumentProperties(
            0, hprinter, printer_name, None, None, 0,
        )
        if dmsize <= 0:
            return None

        driverextra = max(0, dmsize - pywintypes.DEVMODEType().Size)
        devmode = pywintypes.DEVMODEType(driverextra)

        # 現在のデフォルト設定を取得
        win32print.DocumentProperties(
            0, hprinter, printer_name, devmode, None,
            win32con.DM_OUT_BUFFER,
        )

        # A4 縦に設定し、Fields ビットマスクで変更箇所を明示
        devmode.Fields |= (
            win32con.DM_PAPERSIZE | win32con.DM_ORIENTATION
        )
        devmode.PaperSize = win32con.DMPAPER_A4
        devmode.Orientation = win32con.DMORIENT_PORTRAIT

        # ドライバに検証させる（無効な組み合わせを自動補正）
        win32print.DocumentProperties(
            0, hprinter, printer_name, devmode, devmode,
            win32con.DM_IN_BUFFER | win32con.DM_OUT_BUFFER,
        )

        return devmode
    except Exception:
        return None
    finally:
        win32print.ClosePrinter(hprinter)


# ── PIL Image → GDI DC 転送 ──────────────────────────────────────────────────


def _blit_pil_image(
    dc: object, img: object, dpi_x: int, dpi_y: int,
) -> None:
    """PIL Image を GDI DC に StretchDIBits で描画する。

    Args:
        dc: PyCDC オブジェクト
        img: PIL.Image.Image (RGB)
        dpi_x: プリンターの水平 DPI
        dpi_y: プリンターの垂直 DPI
    """
    img_rgb = img.convert('RGB')
    w, h = img_rgb.size

    # BITMAPINFOHEADER (40 bytes)
    bih = struct.pack(
        '<IiiHHIIiiII',
        40,         # biSize
        w,          # biWidth
        -h,         # biHeight (負値=トップダウン)
        1,          # biPlanes
        24,         # biBitCount
        0,          # biCompression (BI_RGB)
        0,          # biSizeImage
        round(dpi_x * 39.3701),  # biXPelsPerMeter
        round(dpi_y * 39.3701),  # biYPelsPerMeter
        0, 0,       # biClrUsed, biClrImportant
    )

    # ピクセルデータ (RGB → BGR 変換 + 各行4バイト境界アライメント)
    # PIL の tobytes('raw', 'BGR') で BGR 順に取得
    stride = (w * 3 + 3) & ~3
    padding = stride - w * 3

    if padding == 0:
        bits = img_rgb.tobytes('raw', 'BGR')
    else:
        # 行ごとにパディング追加
        raw = img_rgb.tobytes('raw', 'BGR')
        rows = []
        for y_row in range(h):
            row_data = raw[y_row * w * 3:(y_row + 1) * w * 3]
            rows.append(row_data + b'\x00' * padding)
        bits = b''.join(rows)

    hdc = dc.GetSafeHdc()

    # 印刷領域サイズ (プリンタードット)
    page_w = dc.GetDeviceCaps(win32con.HORZRES)
    page_h = dc.GetDeviceCaps(win32con.VERTRES)

    # StretchDIBits で画像をページ全体にフィット
    gdi32 = ctypes.windll.gdi32
    gdi32.StretchDIBits(
        hdc,
        0, 0, page_w, page_h,     # 出力先: ページ全体
        0, 0, w, h,                # ソース: 画像全体
        bits,                      # ピクセルデータ
        bih,                       # BITMAPINFOHEADER
        0,                         # DIB_RGB_COLORS
        0x00CC0020,                # SRCCOPY
    )


# ── 印刷ジョブ ───────────────────────────────────────────────────────────────


class PrintJob:
    """GDI 印刷ジョブ。

    使用方法:
        with PrintJob('Microsoft Print to PDF') as job:
            job.start('名簿印刷')
            for lay in filled_layouts:
                job.print_page(lay)
    """

    def __init__(self, printer_name: str) -> None:
        if not HAS_WIN32:
            msg = 'pywin32 が必要です。pip install pywin32 を実行してください。'
            raise RuntimeError(msg)
        self._printer_name = printer_name
        self._dc = None
        self._hdc = None  # win32gui 生ハンドル (DevMode 適用用)
        self._dpi_x = 300
        self._dpi_y = 300
        self._started = False

    def start(self, doc_name: str = '名簿印刷') -> None:
        """印刷ジョブを開始する。

        DocumentProperties で正しい DevMode を構築し、
        win32gui.CreateDC で DevMode 付き DC を作成する。
        """
        devmode = _create_devmode(self._printer_name)

        if devmode is not None:
            # DevMode 付きで DC を作成（用紙サイズ・方向が確実に反映される）
            self._hdc = win32gui.CreateDC(
                'WINSPOOL', self._printer_name, devmode,
            )
            self._dc = win32ui.CreateDCFromHandle(self._hdc)
        else:
            # DevMode 取得失敗時のフォールバック
            self._dc = win32ui.CreateDC()
            self._dc.CreatePrinterDC(self._printer_name)
            self._hdc = None

        self._dpi_x = self._dc.GetDeviceCaps(win32con.LOGPIXELSX)
        self._dpi_y = self._dc.GetDeviceCaps(win32con.LOGPIXELSY)

        self._dc.StartDoc(doc_name)
        self._started = True

    def print_page(self, lay: LayFile) -> None:
        """1 ページ分のレイアウトを印刷する。

        PILBackend で高解像度画像を生成し、StretchDIBits で DC に転送する。
        プレビューと同一の描画コードパスを使用するため、表示と印刷が一致する。
        """
        if not self._started or self._dc is None:
            msg = 'start() を先に呼び出してください。'
            raise RuntimeError(msg)

        self._dc.StartPage()

        # プリンター解像度で PIL 画像を生成（ページ外枠はスキップ）
        dpi = min(self._dpi_x, self._dpi_y)
        img = render_layout_to_image(lay, dpi=dpi, for_print=True)

        # PIL Image → GDI DC に転送
        _blit_pil_image(self._dc, img, self._dpi_x, self._dpi_y)

        self._dc.EndPage()

    def end(self) -> None:
        """印刷ジョブを終了する。"""
        if self._dc is not None:
            if self._started:
                with contextlib.suppress(Exception):
                    self._dc.EndDoc()
            with contextlib.suppress(Exception):
                self._dc.DeleteDC()
            # CreateDCFromHandle の場合、生ハンドルも明示的に解放
            if self._hdc is not None:
                with contextlib.suppress(Exception):
                    win32gui.DeleteDC(self._hdc)
                self._hdc = None
            self._dc = None
        self._started = False

    def __enter__(self) -> PrintJob:
        return self

    def __exit__(self, *_args: object) -> None:
        self.end()
