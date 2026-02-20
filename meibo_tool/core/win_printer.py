"""Windows GDI 印刷エンジン

LayFile をベクター品質で Windows プリンターに直接印刷する。
pywin32 (win32print / win32ui) を使用し、テキストはフォント指定付き DrawText、
罫線は MoveTo/LineTo で描画する。

条件付きインポートにより、Windows 以外の環境でもインポートエラーにならない。
"""

from __future__ import annotations

import contextlib
from typing import TYPE_CHECKING

from core.lay_parser import LayoutObject, ObjectType, resolve_field_name
from core.lay_renderer import model_to_printer

if TYPE_CHECKING:
    from core.lay_parser import LayFile

# ── 条件付きインポート ────────────────────────────────────────────────────────

try:
    import win32con
    import win32print
    import win32ui

    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

# ── 定数 ─────────────────────────────────────────────────────────────────────

# DrawText フラグマッピング
_H_ALIGN_FLAGS = {
    0: 0,  # DT_LEFT = 0
    1: 1,  # DT_CENTER = 1
    2: 2,  # DT_RIGHT = 2
}

_V_ALIGN_TOP = 0
_V_ALIGN_CENTER = 1
_V_ALIGN_BOTTOM = 2


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


# ── 印刷ジョブ ───────────────────────────────────────────────────────────────


class PrintJob:
    """GDI 印刷ジョブ。

    使用方法:
        job = PrintJob('Microsoft Print to PDF')
        job.start('名簿印刷')
        for lay in filled_layouts:
            job.print_page(lay)
        job.end()
    """

    def __init__(self, printer_name: str) -> None:
        if not HAS_WIN32:
            msg = 'pywin32 が必要です。pip install pywin32 を実行してください。'
            raise RuntimeError(msg)
        self._printer_name = printer_name
        self._dc = None
        self._dpi_x = 300
        self._dpi_y = 300
        self._started = False

    def start(self, doc_name: str = '名簿印刷') -> None:
        """印刷ジョブを開始する。"""
        hprinter = win32print.OpenPrinter(self._printer_name)
        try:
            devmode = win32print.GetPrinter(hprinter, 2)['pDevMode']
        finally:
            win32print.ClosePrinter(hprinter)

        # A4 用紙設定
        if devmode:
            devmode.PaperSize = win32con.DMPAPER_A4
            devmode.Orientation = win32con.DMORIENT_PORTRAIT

        self._dc = win32ui.CreateDC()
        self._dc.CreatePrinterDC(self._printer_name)

        if devmode:
            self._dc.ResetDC(devmode)

        self._dpi_x = self._dc.GetDeviceCaps(win32con.LOGPIXELSX)
        self._dpi_y = self._dc.GetDeviceCaps(win32con.LOGPIXELSY)

        self._dc.StartDoc(doc_name)
        self._started = True

    def print_page(self, lay: LayFile) -> None:
        """1 ページ分のレイアウトを印刷する。"""
        if not self._started or self._dc is None:
            msg = 'start() を先に呼び出してください。'
            raise RuntimeError(msg)

        self._dc.StartPage()

        # LABEL / FIELD を先に描画、LINE は最前面
        for obj in lay.objects:
            if obj.obj_type != ObjectType.LINE:
                self._render_object(obj)
        for obj in lay.objects:
            if obj.obj_type == ObjectType.LINE:
                self._render_object(obj)

        self._dc.EndPage()

    def end(self) -> None:
        """印刷ジョブを終了する。"""
        if self._dc is not None:
            if self._started:
                with contextlib.suppress(Exception):
                    self._dc.EndDoc()
            with contextlib.suppress(Exception):
                self._dc.DeleteDC()
            self._dc = None
        self._started = False

    def __enter__(self) -> PrintJob:
        return self

    def __exit__(self, *_args: object) -> None:
        self.end()

    # ── 内部描画 ─────────────────────────────────────────────────────────

    def _model_x(self, x: int) -> int:
        """モデル X 座標 → プリンタドット。"""
        return model_to_printer(x, self._dpi_x)

    def _model_y(self, y: int) -> int:
        """モデル Y 座標 → プリンタドット。"""
        return model_to_printer(y, self._dpi_y)

    def _render_object(self, obj: LayoutObject) -> None:
        """1 オブジェクトを GDI で描画する。"""
        if obj.obj_type == ObjectType.LABEL:
            self._render_label(obj)
        elif obj.obj_type == ObjectType.FIELD:
            self._render_field(obj)
        elif obj.obj_type == ObjectType.LINE:
            self._render_line(obj)

    def _create_font(
        self, name: str, size_pt: float,
        bold: bool = False, italic: bool = False,
    ) -> object:
        """GDI フォントを作成する。"""
        # ポイント → デバイスピクセル
        height = -round(size_pt * self._dpi_y / 72)
        font = win32ui.CreateFont({
            'name': name or 'IPAmj明朝',
            'height': height,
            'weight': 700 if bold else 400,
            'italic': 1 if italic else 0,
            'charset': 128,  # SHIFTJIS_CHARSET
        })
        return font

    @staticmethod
    def _is_vertical_text(obj: LayoutObject, text: str) -> bool:
        """縦書きテキストかどうかを判定する。"""
        if obj.rect is None:
            return False
        w = obj.rect.width
        h = obj.rect.height
        has_newline = '\r' in text or '\n' in text
        return (
            not has_newline
            and len(text) > 1
            and w < h * 0.5
            and w < 120
        )

    @staticmethod
    def _is_multiline_text(text: str) -> bool:
        """複数行テキストかどうかを判定する。"""
        return '\r' in text or '\n' in text

    def _render_label(self, obj: LayoutObject) -> None:
        """LABEL を GDI で描画する。"""
        if obj.rect is None:
            return

        text = obj.prefix + obj.text + obj.suffix
        if not text:
            return

        r = obj.rect
        left = self._model_x(r.left)
        top = self._model_y(r.top)
        right = self._model_x(r.right)
        bottom = self._model_y(r.bottom)

        if self._is_vertical_text(obj, text):
            self._render_vertical_text(
                text, left, top, right, bottom,
                obj.font.name, obj.font.size_pt,
                obj.font.bold, obj.font.italic,
                obj.h_align, obj.v_align,
            )
            return

        if self._is_multiline_text(text):
            self._render_multiline_text(
                text, left, top, right, bottom,
                obj.font.name, obj.font.size_pt,
                obj.font.bold, obj.font.italic,
                obj.h_align, obj.v_align,
            )
            return

        # ── 単一行テキスト: 自動文字サイズ調整 ──
        size_pt = obj.font.size_pt
        box_w = right - left
        font = self._create_font(
            obj.font.name, size_pt, obj.font.bold, obj.font.italic,
        )
        old_font = self._dc.SelectObject(font)

        calc_flags = (
            _H_ALIGN_FLAGS.get(obj.h_align, 0)
            | win32con.DT_SINGLELINE | win32con.DT_NOPREFIX | win32con.DT_CALCRECT
        )
        _h, _w, calc_rect = self._dc.DrawText(
            text, (left, top, left + 10000, bottom), calc_flags,
        )
        text_w = calc_rect[2] - calc_rect[0]
        while text_w > box_w and size_pt > 3:
            self._dc.SelectObject(old_font)
            font.DeleteObject()
            size_pt -= 0.5
            font = self._create_font(
                obj.font.name, size_pt, obj.font.bold, obj.font.italic,
            )
            old_font = self._dc.SelectObject(font)
            _h, _w, calc_rect = self._dc.DrawText(
                text, (left, top, left + 10000, bottom), calc_flags,
            )
            text_w = calc_rect[2] - calc_rect[0]

        flags = _H_ALIGN_FLAGS.get(obj.h_align, 0)
        flags |= win32con.DT_SINGLELINE | win32con.DT_NOPREFIX

        rect = (left, top, right, bottom)
        if obj.v_align in (_V_ALIGN_CENTER, _V_ALIGN_BOTTOM):
            calc_flags = flags | win32con.DT_CALCRECT
            _h, _w, calc_rect = self._dc.DrawText(text, rect, calc_flags)
            text_h = calc_rect[3] - calc_rect[1]
            box_h = bottom - top
            if obj.v_align == _V_ALIGN_CENTER:
                top += (box_h - text_h) // 2
            else:
                top += box_h - text_h
            rect = (left, top, right, bottom)

        self._dc.DrawText(text, rect, flags)
        self._dc.SelectObject(old_font)
        font.DeleteObject()

    def _render_vertical_text(
        self, text: str,
        left: int, top: int, right: int, bottom: int,
        font_name: str, size_pt: float,
        bold: bool, italic: bool,
        h_align: int, v_align: int,
    ) -> None:
        """縦書きテキスト: 各文字を縦に1文字ずつ描画する。"""
        chars = list(text)
        n = len(chars)
        if n == 0:
            return

        box_h = bottom - top

        # 各文字の高さを計算し、収まるまでサイズ縮小
        font = self._create_font(font_name, size_pt, bold, italic)
        old_font = self._dc.SelectObject(font)

        calc_flags = (
            win32con.DT_SINGLELINE | win32con.DT_NOPREFIX | win32con.DT_CALCRECT
        )
        _h, _w, cr = self._dc.DrawText('国', (0, 0, 10000, 10000), calc_flags)
        char_h = cr[3] - cr[1]

        total_h = char_h * n
        while total_h > box_h and size_pt > 3:
            self._dc.SelectObject(old_font)
            font.DeleteObject()
            size_pt -= 0.5
            font = self._create_font(font_name, size_pt, bold, italic)
            old_font = self._dc.SelectObject(font)
            _h, _w, cr = self._dc.DrawText(
                '国', (0, 0, 10000, 10000), calc_flags,
            )
            char_h = cr[3] - cr[1]
            total_h = char_h * n

        # 開始Y位置
        if v_align == _V_ALIGN_CENTER:
            start_y = top + (box_h - total_h) // 2
        elif v_align == _V_ALIGN_BOTTOM:
            start_y = bottom - total_h
        else:
            start_y = top

        # 各文字を描画（水平中央揃え）
        flags = win32con.DT_CENTER | win32con.DT_SINGLELINE | win32con.DT_NOPREFIX
        for i, ch in enumerate(chars):
            cy = start_y + char_h * i
            self._dc.DrawText(ch, (left, cy, right, cy + char_h), flags)

        self._dc.SelectObject(old_font)
        font.DeleteObject()

    def _render_multiline_text(
        self, text: str,
        left: int, top: int, right: int, bottom: int,
        font_name: str, size_pt: float,
        bold: bool, italic: bool,
        h_align: int, v_align: int,
    ) -> None:
        """複数行テキスト: \\r\\n を改行として描画する。"""
        # \r\n → \n に正規化
        text = text.replace('\r\n', '\n').replace('\r', '\n')

        font = self._create_font(font_name, size_pt, bold, italic)
        old_font = self._dc.SelectObject(font)

        # DT_WORDBREAK で複数行描画（DT_SINGLELINE なし）
        flags = _H_ALIGN_FLAGS.get(h_align, 0)
        flags |= win32con.DT_NOPREFIX | win32con.DT_WORDBREAK

        rect = (left, top, right, bottom)
        if v_align in (_V_ALIGN_CENTER, _V_ALIGN_BOTTOM):
            calc_flags = flags | win32con.DT_CALCRECT
            _h, _w, calc_rect = self._dc.DrawText(text, rect, calc_flags)
            text_h = calc_rect[3] - calc_rect[1]
            box_h = bottom - top
            if v_align == _V_ALIGN_CENTER:
                top += (box_h - text_h) // 2
            else:
                top += box_h - text_h
            rect = (left, top, right, bottom)

        self._dc.DrawText(text, rect, flags)
        self._dc.SelectObject(old_font)
        font.DeleteObject()

    def _render_field(self, obj: LayoutObject) -> None:
        """FIELD を GDI で描画する（通常は fill_layout 後なので呼ばれない）。"""
        if obj.rect is None:
            return

        r = obj.rect
        name = resolve_field_name(obj.field_id)
        text = f'{obj.prefix}{{{{{name}}}}}{obj.suffix}'

        left = self._model_x(r.left)
        top = self._model_y(r.top)
        right = self._model_x(r.right)
        bottom = self._model_y(r.bottom)

        font = self._create_font(
            obj.font.name, obj.font.size_pt, obj.font.bold, obj.font.italic,
        )
        old_font = self._dc.SelectObject(font)

        flags = _H_ALIGN_FLAGS.get(obj.h_align, 0)
        flags |= win32con.DT_SINGLELINE | win32con.DT_NOPREFIX

        self._dc.DrawText(text, (left, top, right, bottom), flags)
        self._dc.SelectObject(old_font)
        font.DeleteObject()

    def _render_line(self, obj: LayoutObject) -> None:
        """LINE を GDI で描画する。"""
        if obj.line_start is None or obj.line_end is None:
            return

        x1 = self._model_x(obj.line_start.x)
        y1 = self._model_y(obj.line_start.y)
        x2 = self._model_x(obj.line_end.x)
        y2 = self._model_y(obj.line_end.y)

        # 1pt 幅の黒ペン
        pen = win32ui.CreatePen(
            win32con.PS_SOLID, max(1, self._dpi_x // 72), 0x000000,
        )
        old_pen = self._dc.SelectObject(pen)

        self._dc.MoveTo((x1, y1))
        self._dc.LineTo((x2, y2))

        self._dc.SelectObject(old_pen)
        pen.DeleteObject()
