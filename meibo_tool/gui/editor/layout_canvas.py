"""インタラクティブ レイアウト Canvas

LayFile のオブジェクトを描画し、選択・移動・リサイズの
マウスインタラクションを処理する。
"""

from __future__ import annotations

import dataclasses
import tkinter as tk
from collections.abc import Callable

import customtkinter as ctk

from core.lay_parser import (
    LayFile,
    LayoutObject,
    Point,
    Rect,
)
from core.lay_renderer import (
    CanvasBackend,
    LayRenderer,
    canvas_to_model,
    clear_selection_handles,
    draw_selection_handles,
)

# ── 定数 ─────────────────────────────────────────────────────────────────────

_CANVAS_MARGIN = 30     # ページ外マージン (px)
_MIN_SCALE = 0.15       # ズーム最小 (≈25%)
_MAX_SCALE = 1.2        # ズーム最大 (≈200%)
_SCALE_STEP = 0.05      # ズーム増分
_HANDLE_HIT = 8         # ハンドルヒット判定半径 (px)


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

    # ── 公開メソッド ─────────────────────────────────────────────────────

    def set_layout(self, lay: LayFile) -> None:
        """レイアウトを設定して描画する。"""
        self._lay = lay
        self._selected_idx = -1
        self.refresh()

    def refresh(self) -> None:
        """全オブジェクトを再描画する。"""
        self._canvas.delete('all')
        if self._lay is None:
            return

        backend = self._make_backend()
        renderer = LayRenderer(self._lay, backend)
        renderer.render_all()

        # 選択中なら再描画
        if 0 <= self._selected_idx < len(self._lay.objects):
            draw_selection_handles(
                self._canvas,
                self._lay.objects[self._selected_idx],
                backend,
            )

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

    @property
    def canvas(self) -> tk.Canvas:
        return self._canvas

    # ── 内部メソッド ─────────────────────────────────────────────────────

    def _make_backend(self) -> CanvasBackend:
        return CanvasBackend(
            self._canvas, self._scale,
            offset_x=_CANVAS_MARGIN, offset_y=_CANVAS_MARGIN,
        )

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

        items = self._canvas.find_overlapping(
            cx - 3, cy - 3, cx + 3, cy + 3,
        )

        for item in reversed(items):
            tags = self._canvas.gettags(item)
            for t in tags:
                if t.startswith('obj_'):
                    try:
                        return int(t[4:])
                    except ValueError:
                        pass
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
        else:
            clear_selection_handles(self._canvas)

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
