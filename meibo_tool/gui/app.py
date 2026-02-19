"""名簿帳票ツール — メインウィンドウ

SPEC.md §3.1〜§3.5 参照。
"""

from __future__ import annotations

import contextlib
import logging
import os
import tempfile
import threading
import tkinter.messagebox as mb
import tkinter.ttk as ttk

import customtkinter as ctk
import pandas as pd
from PIL import Image as PILImage

from core.config import get_output_dir, get_template_dir, load_config, save_config
from core.mapper import ensure_fallback_columns
from gui.frames.class_select_panel import ClassSelectPanel
from gui.frames.import_frame import ImportFrame
from gui.frames.output_frame import OutputFrame
from gui.frames.select_frame import SelectFrame
from utils.date_fmt import DATE_KEYS, format_date

logger = logging.getLogger(__name__)


def _validate_required_columns(meta: dict, df: pd.DataFrame) -> None:
    """
    テンプレートの required_columns が df に揃っているか検証する。
    不足列がある場合でも生成は続行する（空欄で出力される）。
    mandatory_columns（組・出席番号）が不足している場合のみ ValueError。
    """
    mandatory = meta.get('mandatory_columns', [])
    missing_mandatory = [c for c in mandatory if c not in df.columns]
    if missing_mandatory:
        raise ValueError(
            f'必須カラムが見つかりません: 「{"」「".join(missing_mandatory)}」\n'
            'ファイルを確認して再選択してください。'
        )

    required = meta.get('required_columns', [])
    missing = [c for c in required if c not in df.columns]
    if missing:
        logger.warning(
            'テンプレート「%s」の推奨カラムが不足: %s（空欄で出力されます）',
            meta.get('file', '?'), missing,
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
            template_dir=get_template_dir(self.config),
            on_teacher_save=self._on_teacher_save,
            on_school_name_save=self._on_school_name_save,
            on_template_change=self._request_preview,
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
        # 未マップ列がある場合はマッピングダイアログを表示
        if unmapped:
            self._show_mapping_dialog(df_mapped, unmapped, source_path)
            return

        self._finalize_import(df_mapped, source_path)

    def _show_mapping_dialog(
        self,
        df_mapped: pd.DataFrame,
        unmapped: list[str],
        source_path: str,
    ) -> None:
        """マッピングダイアログを表示し、確定後に _finalize_import を呼ぶ。"""
        from gui.dialogs.mapping_dialog import MappingDialog

        existing = set(df_mapped.columns)

        def _on_confirm(mapping: dict[str, str]) -> None:
            if mapping:
                df_mapped.rename(columns=mapping, inplace=True)
            self._finalize_import(df_mapped, source_path)

        MappingDialog(self, unmapped, existing, _on_confirm)

    def _finalize_import(
        self, df_mapped: pd.DataFrame, source_path: str,
    ) -> None:
        """マッピング完了後のインポート処理。"""
        # 組・出席番号の存在チェック
        missing = [col for col in ('組', '出席番号') if col not in df_mapped.columns]
        if missing:
            mb.showerror(
                '列が見つかりません',
                f'ファイルに「{"」「".join(missing)}」の列が必要です。\n'
                'ファイルを確認して再選択してください。',
            )
            return

        # 正式氏名 ↔ 氏名 等のフォールバック列を自動補完
        ensure_fallback_columns(df_mapped)

        self.df_mapped = df_mapped

        # パスを config に保存
        self.config['last_loaded_file'] = source_path
        save_config(self.config)

        # ClassSelectPanel にデータをセット（自動でクラス選択・コールバックが走る）
        self.class_select.set_data(df_mapped)

        # 特別支援学級がある場合はオプションを表示
        self.select_frame.set_special_needs_visible(
            self.class_select.has_special_needs()
        )

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

        # プレビューをフィルタリングして更新（特支オプション考慮）
        opts = self.select_frame.get_options()
        df_preview = self._filter_df(
            g, k,
            include_special_needs=opts.get('include_special_needs', False),
            special_needs_placement=opts.get('special_needs_placement', 'appended'),
        )
        self._right.show_data(df_preview)

        # 帳票プレビューも更新
        self._request_preview()

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

        from templates.template_registry import get_all_templates
        all_templates = get_all_templates(options['template_dir'])
        actual = template_key
        if actual not in all_templates:
            return None

        meta = all_templates[actual]
        output_dir = get_output_dir(self.config)

        f = self.class_select.get_filter()
        g = f.get('学年')
        k = f.get('組')

        from core.generator import create_generator

        sn_include = options.get('include_special_needs', False)
        sn_placement = options.get('special_needs_placement', 'appended')

        if k is not None:
            # 単一クラス
            df_filtered = self._filter_df(
                g, k,
                include_special_needs=sn_include,
                special_needs_placement=sn_placement,
            )
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
            df_filtered = self._filter_df(
                gakunen, kumi,
                include_special_needs=sn_include,
                special_needs_placement=sn_placement,
            )
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

    def _filter_df(
        self,
        学年: str | None,
        組: str | None,
        include_special_needs: bool = False,
        special_needs_placement: str = 'appended',
    ) -> pd.DataFrame:
        """df_mapped を学年・組でフィルタリングし、出席番号順にソートして返す。

        include_special_needs=True の場合、同学年の特別支援学級児童も含める。
        """
        from core.special_needs import (
            detect_regular_students,
            detect_special_needs_students,
            merge_special_needs_students,
        )

        df = self.df_mapped.copy()
        if 学年 and '学年' in df.columns:
            df = df[df['学年'] == 学年]

        if 組 and '組' in df.columns:
            # 通常学級でフィルタ
            df_regular = df[df['組'] == 組].copy()
        else:
            # 学年全体/全校の場合は通常学級のみ
            df_regular = detect_regular_students(df)

        # 出席番号で数値ソート
        if '出席番号' in df_regular.columns:
            df_regular['_sort_key'] = pd.to_numeric(
                df_regular['出席番号'], errors='coerce',
            ).fillna(0)
            df_regular = df_regular.sort_values('_sort_key').drop(columns='_sort_key')
            df_regular = df_regular.reset_index(drop=True)

        if include_special_needs and 学年:
            # 同学年の特別支援学級児童を取得
            df_grade = self.df_mapped.copy()
            if '学年' in df_grade.columns:
                df_grade = df_grade[df_grade['学年'] == 学年]
            special_df = detect_special_needs_students(df_grade)
            if not special_df.empty:
                return merge_special_needs_students(
                    df_regular, special_df, placement=special_needs_placement,
                )

        return df_regular

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

    # ────────────────────────────────────────────────────────────────────────
    # 帳票プレビュー
    # ────────────────────────────────────────────────────────────────────────

    _PREVIEW_MAX_STUDENTS = 5
    _PREVIEW_DEBOUNCE_MS = 300

    def _request_preview(self) -> None:
        """プレビュー要求（デバウンス付き）。テンプレート/モード変更時に呼ばれる。"""
        if hasattr(self, '_preview_debounce_id'):
            self.after_cancel(self._preview_debounce_id)
        self._preview_debounce_id = self.after(
            self._PREVIEW_DEBOUNCE_MS, self._generate_preview,
        )

    def _generate_preview(self) -> None:
        """メインスレッドで UI 状態を収集し、バックグラウンドでプレビューを生成する。"""
        if self.df_mapped is None:
            return

        template_key = self.select_frame.get_selected_template()
        if not template_key:
            return

        options = self.select_frame.get_options()
        options['template_dir'] = get_template_dir(self.config)

        from templates.template_registry import get_all_templates
        all_templates = get_all_templates(options['template_dir'])
        if template_key not in all_templates:
            return

        meta = all_templates[template_key]

        # テンプレートファイル存在確認
        tmpl_file = os.path.join(options['template_dir'], meta['file'])
        if not os.path.exists(tmpl_file):
            self._right.show_preview_error(
                f'テンプレートが見つかりません:\n{meta["file"]}'
            )
            return

        # フィルタリング: 現在の学年・組から先頭 N 名（特支オプション考慮）
        f = self.class_select.get_filter()
        df_filtered = self._filter_df(
            f.get('学年'), f.get('組'),
            include_special_needs=options.get('include_special_needs', False),
            special_needs_placement=options.get('special_needs_placement', 'appended'),
        )
        if df_filtered.empty:
            self._right.show_preview_error('表示するデータがありません')
            return

        df_sample = df_filtered.head(self._PREVIEW_MAX_STUDENTS).copy()

        # ローディング表示
        self._right.show_preview_loading()

        # バックグラウンドスレッドで生成
        threading.Thread(
            target=self._preview_worker,
            args=(template_key, df_sample, options),
            daemon=True,
        ).start()

    def _preview_worker(
        self,
        template_key: str,
        df_sample: pd.DataFrame,
        options: dict,
    ) -> None:
        """バックグラウンドでテンプレート生成 → レンダリングを行う。"""
        tmp_path = None
        try:
            from core.generator import create_generator
            from gui.preview_renderer import render_worksheet

            # 一時ファイルに生成
            with tempfile.NamedTemporaryFile(
                suffix='.xlsx', delete=False,
            ) as tmp:
                tmp_path = tmp.name

            gen = create_generator(template_key, tmp_path, df_sample, options)
            gen.generate()

            # openpyxl で読み戻してレンダリング
            from openpyxl import load_workbook
            wb = load_workbook(tmp_path)
            try:
                ws = wb.active
                img = render_worksheet(ws, max_rows=60, max_cols=15, scale=1.5)
            finally:
                wb.close()

            # メインスレッドで表示（ウィンドウ破棄済みなら無視）
            if self.winfo_exists():
                self.after(0, lambda: self._right.show_preview_image(img))

        except Exception:
            logger.exception('プレビュー生成中にエラーが発生しました')
            if self.winfo_exists():
                self.after(
                    0,
                    lambda: self._right.show_preview_error(
                        'プレビューの生成中にエラーが発生しました'
                    ),
                )
        finally:
            if tmp_path:
                with contextlib.suppress(OSError):
                    os.unlink(tmp_path)


# ────────────────────────────────────────────────────────────────────────────
# 右パネル — データプレビュー
# ────────────────────────────────────────────────────────────────────────────

def _fmt_cell(col: str, value) -> str:
    """データプレビュー用のセル値フォーマッター。"""
    s = str(value) if value is not None else ''
    if s.lower() == 'nan' or s == '':
        return ''
    if col in DATE_KEYS:
        return format_date(s)
    return s


class _PreviewPanel(ctk.CTkFrame):
    """右パネル: データ一覧 + 帳票プレビュー（CTkTabview）。"""

    _PRIORITY_COLS = ('出席番号', '組', '学年', '氏名', '氏名かな', '性別', '生年月日')

    def __init__(self, master) -> None:
        super().__init__(master, corner_radius=6)
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # タブビュー
        self._tabview = ctk.CTkTabview(self, corner_radius=6)
        self._tabview.grid(row=0, column=0, sticky='nsew', padx=4, pady=4)

        self._tab_data = self._tabview.add('データ')
        self._tab_preview = self._tabview.add('プレビュー')

        # ── データタブ ─────────────────────────────────────────────────────
        self._tab_data.grid_rowconfigure(1, weight=1)
        self._tab_data.grid_columnconfigure(0, weight=1)

        self._header = ctk.CTkLabel(
            self._tab_data,
            text='データプレビュー',
            font=ctk.CTkFont(size=13, weight='bold'),
        )
        self._header.grid(row=0, column=0, sticky='w', padx=6, pady=(4, 2))

        self._data_placeholder = ctk.CTkLabel(
            self._tab_data,
            text='名簿ファイルを読み込むと\nここにプレビューが表示されます',
            text_color='gray',
        )
        self._data_placeholder.grid(row=1, column=0, padx=20, pady=40)

        self._tree_container: ctk.CTkFrame | None = None

        # ── プレビュータブ ─────────────────────────────────────────────────
        self._tab_preview.grid_rowconfigure(0, weight=1)
        self._tab_preview.grid_columnconfigure(0, weight=1)

        self._preview_scroll = ctk.CTkScrollableFrame(
            self._tab_preview, fg_color='transparent',
        )
        self._preview_scroll.grid(row=0, column=0, sticky='nsew')
        self._preview_scroll.grid_columnconfigure(0, weight=1)

        self._preview_label = ctk.CTkLabel(
            self._preview_scroll,
            text='テンプレートを選択すると\n帳票プレビューが表示されます',
            text_color='gray',
        )
        self._preview_label.grid(row=0, column=0, padx=20, pady=40)

        # CTkImage の GC 防止用リファレンス
        self._preview_image_ref: ctk.CTkImage | None = None

    # ── データタブ API ─────────────────────────────────────────────────────

    def show_data(self, df: pd.DataFrame) -> None:
        """DataFrame を Treeview に表示する。"""
        if self._tree_container:
            self._tree_container.destroy()
            self._tree_container = None
        self._data_placeholder.grid_remove()

        show_cols = [c for c in self._PRIORITY_COLS if c in df.columns]
        if not show_cols:
            show_cols = list(df.columns[:10])

        container = ctk.CTkFrame(self._tab_data, corner_radius=0, fg_color='transparent')
        container.grid(row=1, column=0, sticky='nsew', padx=2, pady=2)
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
            tree.insert('', 'end', values=[
                _fmt_cell(c, row.get(c, '')) for c in show_cols
            ])

        vsb = ttk.Scrollbar(container, orient='vertical', command=tree.yview)
        hsb = ttk.Scrollbar(container, orient='horizontal', command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')

        self._header.configure(text=f'データプレビュー — {len(df)} 名')

    # ── プレビュータブ API ─────────────────────────────────────────────────

    def show_preview_loading(self) -> None:
        """プレビュー生成中のローディング表示。"""
        self._preview_image_ref = None
        self._preview_label.configure(
            text='プレビューを生成中…', text_color='gray', image=None,
        )
        self._preview_label.grid(row=0, column=0, padx=20, pady=40)

    def show_preview_image(self, pil_image: PILImage.Image) -> None:
        """PIL Image をプレビュータブに表示する。"""
        ctk_img = ctk.CTkImage(
            light_image=pil_image,
            size=(pil_image.width, pil_image.height),
        )
        self._preview_image_ref = ctk_img  # GC 防止
        self._preview_label.configure(text='', image=ctk_img, text_color='gray')
        self._preview_label.grid(row=0, column=0, padx=4, pady=4)
        # プレビュータブに自動切替
        self._tabview.set('プレビュー')

    def show_preview_error(self, msg: str) -> None:
        """プレビューエラーメッセージを表示する。"""
        self._preview_image_ref = None
        self._preview_label.configure(
            text=msg, text_color='#CC4444', image=None,
        )
        self._preview_label.grid(row=0, column=0, padx=20, pady=40)
