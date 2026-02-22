"""インタラクティブ レイアウト Canvas

LayFile のオブジェクトを PIL でレンダリングし、選択・移動・リサイズの
マウスインタラクションを処理する。
"""

from __future__ import annotations

import dataclasses
import math
import tkinter as tk
from collections.abc import Callable

import customtkinter as ctk
from PIL import ImageTk

from core.lay_parser import (
    LayFile,
    LayoutObject,
    Point,
    Rect,
)
from core.lay_renderer import (
    canvas_to_model,
    model_to_canvas,
    render_layout_to_image,
)

# ── 定数 ─────────────────────────────────────────────────────────────────────

_CANVAS_MARGIN = 30     # ページ外マージン (px)
_MIN_SCALE = 0.15       # ズーム最小 (≈25%)
_MAX_SCALE = 1.2        # ズーム最大 (≈200%)
_SCALE_STEP = 0.05      # ズーム増分
_HANDLE_HIT = 8         # ハンドルヒット判定半径 (px)
_HANDLE_SIZE = 5        # ハンドル矩形の半径 (px)
_SELECT_OUTLINE = '#1A73E8'
_HANDLE_COLOR = '#1A73E8'
_LINE_HIT_THRESHOLD = 5  # LINE ヒット判定距離 (モデル単位)


def _point_near_line(
    mx: int, my: int,
    p1: Point, p2: Point,
    threshold: int = _LINE_HIT_THRESHOLD,
) -> bool:
    """点 (mx, my) が線分 (p1, p2) の近傍にあるか判定する。"""
    dx = p2.x - p1.x
    dy = p2.y - p1.y
    length_sq = dx * dx + dy * dy
    if length_sq == 0:
        return math.hypot(mx - p1.x, my - p1.y) <= threshold

    t = max(0.0, min(1.0, ((mx - p1.x) * dx + (my - p1.y) * dy) / length_sq))
    proj_x = p1.x + t * dx
    proj_y = p1.y + t * dy
    return math.hypot(mx - proj_x, my - proj_y) <= threshold


class LayoutCanvas(ctk.CTkFrame):
    """インタラクティブ Canvas。

    選択・移動・リサイズを処理し、コールバックで変更を通知する。
    """

    def __init__(
        self, master: ctk.CTkBaseClass,
        on_select: Callable[[int], None] | None = None,
        on_modify: Callable[[int, LayoutObject], None] | None = None,
        on_cursor: Callable[[int, int], None] | None = None,
    ) -> None:
        super().__init__(master)
        self._on_select = on_select
        self._on_modify = on_modify
        self._on_cursor = on_cursor

        self._lay: LayFile | None = None
        self._scale = 0.5
        self._selected_idx: int = -1
        self._layout_registry: dict[str, LayFile] = {}
        self._photo_image: ImageTk.PhotoImage | None = None  # GC 防止

        # ドラッグ状態
        self._dragging = False
        self._drag_start_x = 0.0
        self._drag_start_y = 0.0
        self._drag_handle: str | None = None  # リサイズ中のハンドル名
        self._drag_orig_obj: LayoutObject | None = None  # ドラッグ前のスナップショット

        self._build_canvas()

    # ── UI 構築 ───────────────────────────────────────────────────────────

    def _build_canvas(self) -> None:
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self._canvas = tk.Canvas(
            self, bg='#E0E0E0', highlightthickness=0,
        )
        self._canvas.grid(row=0, column=0, sticky='nsew')

        # スクロールバー
        self._vscroll = ctk.CTkScrollbar(
            self, orientation='vertical', command=self._canvas.yview,
        )
        self._vscroll.grid(row=0, column=1, sticky='ns')
        self._hscroll = ctk.CTkScrollbar(
            self, orientation='horizontal', command=self._canvas.xview,
        )
        self._hscroll.grid(row=1, column=0, sticky='ew')

        self._canvas.configure(
            xscrollcommand=self._hscroll.set,
            yscrollcommand=self._vscroll.set,
        )

        # イベントバインド
        self._canvas.bind('<Button-1>', self._on_click)
        self._canvas.bind('<B1-Motion>', self._on_drag)
        self._canvas.bind('<ButtonRelease-1>', self._on_release)
        self._canvas.bind('<Motion>', self._on_motion)
        self._canvas.bind('<MouseWheel>', self._on_mousewheel)
        self._canvas.bind('<Shift-MouseWheel>', self._on_shift_mousewheel)
        self._canvas.bind('<Control-MouseWheel>', self._on_ctrl_mousewheel)

    # ── 公開メソッド ─────────────────────────────────────────────────────

    def set_layout(
        self, lay: LayFile,
        layout_registry: dict[str, LayFile] | None = None,
    ) -> None:
        """レイアウトを設定して描画する。"""
        self._lay = lay
        self._selected_idx = -1
        if layout_registry is not None:
            self._layout_registry = layout_registry
        self.refresh()
        self.zoom_to_fit()

    def refresh(self) -> None:
        """全オブジェクトを再描画する。"""
        self._canvas.delete('all')
        if self._lay is None:
            return

        # PIL でレンダリング → PhotoImage として Canvas に配置
        unit_mm = self._lay.paper.unit_mm if self._lay.paper else 0.25
        dpi = max(1, int(self._scale * 25.4 / unit_mm))
        img = render_layout_to_image(
            self._lay, dpi=dpi,
            layout_registry=self._layout_registry,
            editor_mode=True,
        )
        self._photo_image = ImageTk.PhotoImage(img)
        self._canvas.create_image(
            _CANVAS_MARGIN, _CANVAS_MARGIN,
            image=self._photo_image, anchor='nw',
        )

        # 選択中ならハンドルをオーバーレイ
        if 0 <= self._selected_idx < len(self._lay.objects):
            self._draw_selection_handles(self._lay.objects[self._selected_idx])

        self._update_scroll_region()

    def select(self, index: int) -> None:
        """指定インデックスのオブジェクトを選択する。"""
        self._selected_idx = index
        self.refresh()

    def get_selected_index(self) -> int:
        return self._selected_idx

    def set_zoom(self, scale: float) -> None:
        """ズームレベルを設定する。"""
        self._scale = max(_MIN_SCALE, min(_MAX_SCALE, scale))
        self.refresh()

    def get_zoom(self) -> float:
        return self._scale

    def zoom_in(self) -> None:
        self.set_zoom(self._scale + _SCALE_STEP)

    def zoom_out(self) -> None:
        self.set_zoom(self._scale - _SCALE_STEP)

    def zoom_to_fit(self) -> None:
        """Canvas サイズに合わせてズームする。"""
        if self._lay is None:
            return
        self.update_idletasks()
        cw = self._canvas.winfo_width()
        ch = self._canvas.winfo_height()
        if cw <= 1 or ch <= 1:
            return
        pw = max(1, self._lay.page_width)
        ph = max(1, self._lay.page_height)
        sx = (cw - _CANVAS_MARGIN * 2) / pw
        sy = (ch - _CANVAS_MARGIN * 2) / ph
        self.set_zoom(min(sx, sy))

    @property
    def canvas(self) -> tk.Canvas:
        return self._canvas

    # ── 内部メソッド ─────────────────────────────────────────────────────

    def _update_scroll_region(self) -> None:
        if self._lay is None:
            return
        w = self._lay.page_width * self._scale + _CANVAS_MARGIN * 2
        h = self._lay.page_height * self._scale + _CANVAS_MARGIN * 2
        self._canvas.configure(scrollregion=(0, 0, w, h))

    def _find_object_at(self, cx: float, cy: float) -> int:
        """Canvas 座標位置のオブジェクトインデックスを返す。見つからなければ -1。"""
        if self._lay is None:
            return -1

        mx, my = canvas_to_model(
            cx, cy, self._scale,
            _CANVAS_MARGIN, _CANVAS_MARGIN,
        )

        for i in reversed(range(len(self._lay.objects))):
            obj = self._lay.objects[i]
            if obj.rect:
                r = obj.rect
                if r.left <= mx <= r.right and r.top <= my <= r.bottom:
                    return i
            if obj.line_start and obj.line_end and _point_near_line(mx, my, obj.line_start, obj.line_end):
                return i
        return -1

    def _find_handle_at(self, cx: float, cy: float) -> str | None:
        """ハンドル名を返す。ハンドル上でなければ None。"""
        items = self._canvas.find_overlapping(
            cx - _HANDLE_HIT, cy - _HANDLE_HIT,
            cx + _HANDLE_HIT, cy + _HANDLE_HIT,
        )
        for item in reversed(items):
            tags = self._canvas.gettags(item)
            for t in tags:
                if t.startswith('handle_'):
                    return t[7:]  # 'nw', 'se', etc.
        return None

    def _draw_selection_handles(self, obj: LayoutObject) -> None:
        """選択オブジェクトの周囲にリサイズハンドルを Canvas 上にオーバーレイ描画する。"""
        self._canvas.delete('handles')

        if obj.rect is not None:
            r = obj.rect
            px1, py1 = model_to_canvas(
                r.left, r.top, self._scale,
                _CANVAS_MARGIN, _CANVAS_MARGIN,
            )
            px2, py2 = model_to_canvas(
                r.right, r.bottom, self._scale,
                _CANVAS_MARGIN, _CANVAS_MARGIN,
            )
        elif obj.line_start is not None and obj.line_end is not None:
            px1, py1 = model_to_canvas(
                obj.line_start.x, obj.line_start.y, self._scale,
                _CANVAS_MARGIN, _CANVAS_MARGIN,
            )
            px2, py2 = model_to_canvas(
                obj.line_end.x, obj.line_end.y, self._scale,
                _CANVAS_MARGIN, _CANVAS_MARGIN,
            )
        else:
            return

        # 選択枠
        self._canvas.create_rectangle(
            px1 - 1, py1 - 1, px2 + 1, py2 + 1,
            outline=_SELECT_OUTLINE, width=2, dash=(4, 4),
            tags=('handles',),
        )

        # 8 ハンドル (矩形) / LINE は 2 ハンドル
        hs = _HANDLE_SIZE
        if obj.rect is not None:
            cx_h, cy_h = (px1 + px2) / 2, (py1 + py2) / 2
            handle_positions = {
                'nw': (px1, py1), 'n': (cx_h, py1), 'ne': (px2, py1),
                'w': (px1, cy_h), 'e': (px2, cy_h),
                'sw': (px1, py2), 's': (cx_h, py2), 'se': (px2, py2),
            }
        else:
            handle_positions = {
                'start': (px1, py1),
                'end': (px2, py2),
            }

        for name, (hx, hy) in handle_positions.items():
            self._canvas.create_rectangle(
                hx - hs, hy - hs, hx + hs, hy + hs,
                fill=_HANDLE_COLOR, outline='white', width=1,
                tags=('handles', f'handle_{name}'),
            )

    # ── マウスイベント ───────────────────────────────────────────────────

    def _on_click(self, event: tk.Event) -> None:
        cx = self._canvas.canvasx(event.x)
        cy = self._canvas.canvasy(event.y)

        # まずハンドルチェック
        handle = self._find_handle_at(cx, cy)
        if handle and self._selected_idx >= 0:
            self._dragging = True
            self._drag_handle = handle
            self._drag_start_x = cx
            self._drag_start_y = cy
            obj = self._lay.objects[self._selected_idx]
            self._drag_orig_obj = dataclasses.replace(obj)
            if obj.rect:
                self._drag_orig_obj.rect = Rect(
                    obj.rect.left, obj.rect.top,
                    obj.rect.right, obj.rect.bottom,
                )
            return

        # オブジェクト選択
        idx = self._find_object_at(cx, cy)
        self._selected_idx = idx

        if self._on_select:
            self._on_select(idx)

        if idx >= 0:
            self._dragging = True
            self._drag_handle = None
            self._drag_start_x = cx
            self._drag_start_y = cy
            obj = self._lay.objects[idx]
            self._drag_orig_obj = dataclasses.replace(obj)
            if obj.rect:
                self._drag_orig_obj.rect = Rect(
                    obj.rect.left, obj.rect.top,
                    obj.rect.right, obj.rect.bottom,
                )
            if obj.line_start:
                self._drag_orig_obj.line_start = Point(
                    obj.line_start.x, obj.line_start.y,
                )
            if obj.line_end:
                self._drag_orig_obj.line_end = Point(
                    obj.line_end.x, obj.line_end.y,
                )

        self.refresh()

    def _on_drag(self, event: tk.Event) -> None:
        if not self._dragging or self._lay is None or self._selected_idx < 0:
            return

        cx = self._canvas.canvasx(event.x)
        cy = self._canvas.canvasy(event.y)
        dx_px = cx - self._drag_start_x
        dy_px = cy - self._drag_start_y

        # ピクセルデルタ → モデルデルタ
        dm_x = round(dx_px / self._scale)
        dm_y = round(dy_px / self._scale)

        obj = self._lay.objects[self._selected_idx]
        orig = self._drag_orig_obj

        if self._drag_handle:
            # リサイズ
            self._apply_resize(obj, orig, dm_x, dm_y)
        else:
            # 移動
            self._apply_move(obj, orig, dm_x, dm_y)

        self.refresh()

    def _on_release(self, event: tk.Event) -> None:
        if self._dragging and self._on_modify and self._drag_orig_obj is not None:
            self._on_modify(self._selected_idx, self._drag_orig_obj)
        self._dragging = False
        self._drag_handle = None
        self._drag_orig_obj = None

    def _on_motion(self, event: tk.Event) -> None:
        """カーソル位置をモデル座標で通知する。"""
        if self._on_cursor and self._lay:
            cx = self._canvas.canvasx(event.x)
            cy = self._canvas.canvasy(event.y)
            mx, my = canvas_to_model(
                cx, cy, self._scale,
                _CANVAS_MARGIN, _CANVAS_MARGIN,
            )
            self._on_cursor(mx, my)

    # ── 移動・リサイズ ───────────────────────────────────────────────────

    def _apply_move(
        self, obj: LayoutObject, orig: LayoutObject,
        dm_x: int, dm_y: int,
    ) -> None:
        """オブジェクトを移動する。"""
        if orig.rect and obj.rect:
            obj.rect = Rect(
                orig.rect.left + dm_x, orig.rect.top + dm_y,
                orig.rect.right + dm_x, orig.rect.bottom + dm_y,
            )
        if orig.line_start and obj.line_start:
            obj.line_start = Point(
                orig.line_start.x + dm_x, orig.line_start.y + dm_y,
            )
        if orig.line_end and obj.line_end:
            obj.line_end = Point(
                orig.line_end.x + dm_x, orig.line_end.y + dm_y,
            )

    def _apply_resize(
        self, obj: LayoutObject, orig: LayoutObject,
        dm_x: int, dm_y: int,
    ) -> None:
        """リサイズハンドルに応じてオブジェクトをリサイズする。"""
        if orig.rect is None or obj.rect is None:
            return

        handle = self._drag_handle
        left = orig.rect.left
        top = orig.rect.top
        right = orig.rect.right
        bottom = orig.rect.bottom

        if 'w' in handle:
            left += dm_x
        if 'e' in handle:
            right += dm_x
        if 'n' in handle:
            top += dm_y
        if 's' in handle:
            bottom += dm_y

        # 最小サイズ保証
        if right - left < 4:
            right = left + 4
        if bottom - top < 4:
            bottom = top + 4

        obj.rect = Rect(left, top, right, bottom)

    # ── マウスホイール ───────────────────────────────────────────────────

    def _on_mousewheel(self, event: tk.Event) -> None:
        """垂直スクロール。"""
        self._canvas.yview_scroll(-1 * (event.delta // 120), 'units')

    def _on_shift_mousewheel(self, event: tk.Event) -> None:
        """水平スクロール。"""
        self._canvas.xview_scroll(-1 * (event.delta // 120), 'units')

    def _on_ctrl_mousewheel(self, event: tk.Event) -> None:
        """ズーム。"""
        if event.delta > 0:
            self.zoom_in()
        else:
            self.zoom_out()
