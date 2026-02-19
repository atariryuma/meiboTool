"""名簿帳票ツール — メインウィンドウ

SPEC.md §3.1〜§3.4 参照。
"""

from __future__ import annotations

import os
import threading
import tkinter.messagebox as mb
import tkinter.ttk as ttk

import customtkinter as ctk
import pandas as pd

from core.config import get_output_dir, get_template_dir, load_config, save_config
from gui.frames.class_select_panel import ClassSelectPanel
from gui.frames.import_frame import ImportFrame
from gui.frames.output_frame import OutputFrame
from gui.frames.select_frame import SelectFrame


def _validate_required_columns(meta: dict, df: pd.DataFrame) -> None:
    """
    テンプレートの required_columns が df に揃っているか検証する。
    不足している場合は ValueError を raise する（OutputFrame がキャッチして表示）。
    """
    required = meta.get('required_columns', [])
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(
            f'必須カラムが見つかりません: 「{"」「".join(missing)}」\n'
            'テンプレートに必要な列がファイルに含まれているか確認してください。'
        )


class _BatchGenerator:
    """複数クラスを一括生成するラッパー。generate() は出力フォルダを返す。"""

    def __init__(self, generators: list, output_dir: str) -> None:
        self._generators = generators
        self._output_dir = output_dir

    def generate(self) -> str:
        for gen in self._generators:
            gen.generate()
        # OutputFrame は dirname(result) を startfile するため、
        # "出力フォルダ/一括生成完了" を返してフォルダを開かせる
        return os.path.join(self._output_dir, '一括生成完了')


class App(ctk.CTk):
    """メインウィンドウ。状態管理と各フレームの調整役。"""

    def __init__(self) -> None:
        super().__init__()
        self.title('名簿帳票ツール v1.0')
        self.geometry('960x700')
        self.minsize(800, 600)
        ctk.set_appearance_mode('light')
        ctk.set_default_color_theme('blue')

        self.config = load_config()

        # データ状態
        self.df_mapped: pd.DataFrame | None = None

        self._build_layout()
        self._check_update_bg()
        self._sync_data_bg()
        # 起動時に前回ファイルを自動再読込
        self.after(100, self._on_startup)

    # ────────────────────────────────────────────────────────────────────────
    # レイアウト構築
    # ────────────────────────────────────────────────────────────────────────

    def _build_layout(self) -> None:
        self.grid_columnconfigure(0, weight=0, minsize=320)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # ── ヘッダーバー（設定ボタン） ───────────────────────────────────
        header_bar = ctk.CTkFrame(self, height=36, corner_radius=0, fg_color='transparent')
        header_bar.grid(row=0, column=0, columnspan=2, sticky='ew', padx=5, pady=(5, 0))
        header_bar.grid_columnconfigure(0, weight=1)

        ctk.CTkButton(
            header_bar, text='設定', width=60, height=28,
            command=self._open_settings,
        ).grid(row=0, column=1, padx=5, sticky='e')

        # 左パネル（スクロール可）
        self._left = ctk.CTkScrollableFrame(self, width=310, corner_radius=0)
        self._left.grid(row=1, column=0, sticky='nsew', padx=(5, 2), pady=5)
        self._left.grid_columnconfigure(0, weight=1)

        # 右パネル
        self._right = _PreviewPanel(self)
        self._right.grid(row=1, column=1, sticky='nsew', padx=(2, 5), pady=5)

        # セクション配置
        self.import_frame = ImportFrame(self._left, on_load=self._on_import)
        self.import_frame.grid(row=0, column=0, sticky='ew', padx=5, pady=(5, 3))

        self.class_select = ClassSelectPanel(
            self._left, on_select=self._on_class_select
        )
        self.class_select.grid(row=1, column=0, sticky='ew', padx=5, pady=3)

        self.select_frame = SelectFrame(
            self._left,
            config=self.config,
            on_teacher_save=self._on_teacher_save,
            on_school_name_save=self._on_school_name_save,
        )
        self.select_frame.grid(row=2, column=0, sticky='ew', padx=5, pady=3)
        self.select_frame.set_enabled(False)

        self.output_frame = OutputFrame(
            self._left,
            on_generate=self._create_generator,
            config=self.config,
        )
        self.output_frame.grid(row=3, column=0, sticky='ew', padx=5, pady=(3, 5))
        self.output_frame.set_enabled(False)

    # ────────────────────────────────────────────────────────────────────────
    # 起動時処理
    # ────────────────────────────────────────────────────────────────────────

    def _on_startup(self) -> None:
        """前回読み込んだファイルを自動再読込する。"""
        path = self.config.get('last_loaded_file', '')
        if path and os.path.exists(path):
            self.import_frame.load_from_path(path)

    # ────────────────────────────────────────────────────────────────────────
    # コールバック
    # ────────────────────────────────────────────────────────────────────────

    def _on_import(
        self, df_mapped: pd.DataFrame, unmapped: list[str], source_path: str
    ) -> None:
        """ファイル読込完了後に呼ばれる。"""
        # 組・出席番号の存在チェック
        missing = [col for col in ('組', '出席番号') if col not in df_mapped.columns]
        if missing:
            mb.showerror(
                '列が見つかりません',
                f'ファイルに「{"」「".join(missing)}」の列が必要です。\n'
                'ファイルを確認して再選択してください。',
            )
            return

        self.df_mapped = df_mapped

        # パスを config に保存
        self.config['last_loaded_file'] = source_path
        save_config(self.config)

        # ClassSelectPanel にデータをセット（自動でクラス選択・コールバックが走る）
        self.class_select.set_data(df_mapped)

        # プレビュー更新
        self._right.show_data(df_mapped)

    def _on_class_select(self, filter_dict: dict) -> None:
        """クラス/学年選択時に呼ばれる。"""
        g = filter_dict.get('学年')
        k = filter_dict.get('組')

        # 単一クラス選択時は担任名を自動セット
        if g and k:
            key = f'{g}-{k}'
            teacher = self.config.get('homeroom_teachers', {}).get(key, '')
            self.select_frame.set_teacher(teacher)

        # 以降のセクションを有効化
        self.select_frame.set_enabled(True)
        self.output_frame.set_enabled(True)

        # プレビューをフィルタリングして更新
        df_preview = self._filter_df(g, k)
        self._right.show_data(df_preview)

    def _on_teacher_save(self, name: str) -> None:
        """SelectFrame の「保存」ボタン押下時に呼ばれる。"""
        f = self.class_select.get_filter()
        g = f.get('学年')
        k = f.get('組')
        if g and k:
            key = f'{g}-{k}'
            if 'homeroom_teachers' not in self.config:
                self.config['homeroom_teachers'] = {}
            self.config['homeroom_teachers'][key] = name
            save_config(self.config)

    def _on_school_name_save(self, name: str) -> None:
        """SelectFrame の学校名「保存」ボタン押下時に呼ばれる。"""
        self.config['school_name'] = name
        save_config(self.config)

    def _create_generator(self):
        """
        OutputFrame のスレッドから呼ばれる。
        生成可能な BaseGenerator インスタンス（or None）を返す。
        一括生成の場合は _BatchGenerator を返す。
        """
        if self.df_mapped is None:
            return None

        template_key = self.select_frame.get_selected_template()
        if not template_key:
            return None

        options = self.select_frame.get_options()
        options['template_dir'] = get_template_dir(self.config)

        from templates.template_registry import TEMPLATES
        actual = template_key
        if actual not in TEMPLATES:
            return None

        meta = TEMPLATES[actual]
        output_dir = get_output_dir(self.config)

        f = self.class_select.get_filter()
        g = f.get('学年')
        k = f.get('組')

        from core.generator import create_generator

        if k is not None:
            # 単一クラス
            df_filtered = self._filter_df(g, k)
            _validate_required_columns(meta, df_filtered)
            suffix = f'{g}年{k}組'
            out_file = meta['file'].replace('.xlsx', f'_{suffix}.xlsx')
            output_path = os.path.join(output_dir, out_file)
            return create_generator(actual, output_path, df_filtered, options)

        # 一括生成（学年単位 or 全校）— 最初のクラスでカラムを検証
        classes = self.class_select.get_available_classes(学年=g)
        if classes:
            _validate_required_columns(meta, self._filter_df(*classes[0]))
        generators = []
        homeroom = self.config.get('homeroom_teachers', {})
        for (gakunen, kumi) in classes:
            df_filtered = self._filter_df(gakunen, kumi)
            opts = dict(options)
            # クラスごとの担任名を config から上書き（登録済みの場合）
            saved_teacher = homeroom.get(f'{gakunen}-{kumi}', '')
            if saved_teacher:
                opts['teacher_name'] = saved_teacher
            suffix = f'{gakunen}年{kumi}組'
            out_file = meta['file'].replace('.xlsx', f'_{suffix}.xlsx')
            output_path = os.path.join(output_dir, out_file)
            generators.append(create_generator(actual, output_path, df_filtered, opts))

        return _BatchGenerator(generators, output_dir) if generators else None

    # ────────────────────────────────────────────────────────────────────────
    # ヘルパー
    # ────────────────────────────────────────────────────────────────────────

    def _filter_df(self, 学年: str | None, 組: str | None) -> pd.DataFrame:
        """df_mapped を学年・組でフィルタリングして返す。"""
        df = self.df_mapped.copy()
        if 学年 and '学年' in df.columns:
            df = df[df['学年'] == 学年]
        if 組 and '組' in df.columns:
            df = df[df['組'] == 組]
        return df

    # ────────────────────────────────────────────────────────────────────────
    # バックグラウンド更新チェック
    # ────────────────────────────────────────────────────────────────────────

    def _check_update_bg(self) -> None:
        """バックグラウンドでアプリの新バージョンを確認する。"""
        def _task() -> None:
            try:
                from core.updater import check_for_update
                info = check_for_update(self.config)
                if info:
                    self.after(0, lambda: self._show_update_dialog(info))
            except Exception:
                pass  # ネットワークエラー等は無視

        threading.Thread(target=_task, daemon=True).start()

    def _show_update_dialog(self, info) -> None:
        """更新ダイアログを表示する。"""
        from gui.dialogs.update_dialog import UpdateDialog
        UpdateDialog(self, info, self.config)

    def _sync_data_bg(self) -> None:
        """バックグラウンドで名簿データを同期する。"""
        def _task() -> None:
            try:
                from core.data_sync import sync
                result = sync(self.config)
                self.after(0, lambda: self._on_sync_result(result))
            except Exception:
                pass  # 同期エラーは無視

        threading.Thread(target=_task, daemon=True).start()

    def _on_sync_result(self, result) -> None:
        """データ同期結果を処理する（メインスレッドで呼ばれる）。"""
        # config 更新をメインスレッドで適用
        if result.config_updates:
            ds = self.config.setdefault('data_source', {})
            ds.update(result.config_updates)
            import contextlib
            with contextlib.suppress(OSError):
                save_config(self.config)

        if result.status == 'updated' and result.path:
            self.import_frame.load_from_path(result.path)
            self.import_frame.show_sync_status(
                f'名簿データを自動更新しました ({result.path})'
            )
        elif result.status == 'unavailable' and result.path:
            self.import_frame.load_from_path(result.path)
            self.import_frame.show_sync_status(result.message, warning=True)
        elif result.status == 'decrypt_error':
            mb.showwarning('復号エラー', result.message)
        elif result.status == 'error' and result.message:
            self.import_frame.show_sync_status(result.message, warning=True)

    def _open_settings(self) -> None:
        """設定ダイアログを開く。"""
        from gui.dialogs.settings_dialog import SettingsDialog
        SettingsDialog(self, self.config)


# ────────────────────────────────────────────────────────────────────────────
# 右パネル — データプレビュー
# ────────────────────────────────────────────────────────────────────────────

class _PreviewPanel(ctk.CTkFrame):
    """右パネル: 読込データを ttk.Treeview で表示する。"""

    _PRIORITY_COLS = ('出席番号', '組', '学年', '氏名', '氏名かな', '性別', '生年月日')

    def __init__(self, master) -> None:
        super().__init__(master, corner_radius=6)
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self._header = ctk.CTkLabel(
            self,
            text='データプレビュー',
            font=ctk.CTkFont(size=13, weight='bold'),
        )
        self._header.grid(row=0, column=0, sticky='w', padx=10, pady=(8, 4))

        self._placeholder = ctk.CTkLabel(
            self,
            text='名簿ファイルを読み込むと\nここにプレビューが表示されます',
            text_color='gray',
        )
        self._placeholder.grid(row=1, column=0, padx=20, pady=40)

        self._tree_container: ctk.CTkFrame | None = None

    def show_data(self, df: pd.DataFrame) -> None:
        """DataFrame を Treeview に表示する。"""
        if self._tree_container:
            self._tree_container.destroy()
            self._tree_container = None
        self._placeholder.grid_remove()

        show_cols = [c for c in self._PRIORITY_COLS if c in df.columns]
        if not show_cols:
            show_cols = list(df.columns[:10])

        container = ctk.CTkFrame(self, corner_radius=0, fg_color='transparent')
        container.grid(row=1, column=0, sticky='nsew', padx=5, pady=5)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)
        self._tree_container = container

        style = ttk.Style()
        style.configure('Roster.Treeview', font=('メイリオ', 9), rowheight=22)
        style.configure('Roster.Treeview.Heading', font=('メイリオ', 9, 'bold'))

        tree = ttk.Treeview(
            container,
            columns=show_cols,
            show='headings',
            height=25,
            style='Roster.Treeview',
        )
        for col in show_cols:
            w = 70 if col in ('出席番号', '組', '学年', '性別') else 110
            tree.heading(col, text=col)
            tree.column(col, width=w, minwidth=50, anchor='center')

        for _, row in df.iterrows():
            tree.insert('', 'end', values=[str(row.get(c, '')) for c in show_cols])

        vsb = ttk.Scrollbar(container, orient='vertical', command=tree.yview)
        hsb = ttk.Scrollbar(container, orient='horizontal', command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')

        self._header.configure(text=f'データプレビュー — {len(df)} 名')
