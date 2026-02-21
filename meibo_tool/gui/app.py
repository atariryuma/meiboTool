"""名簿帳票ツール — メインウィンドウ

SPEC.md §3.1〜§3.5 参照。
"""

from __future__ import annotations

import contextlib
import logging
import os
import tempfile
import threading
import tkinter as tk
import tkinter.messagebox as mb
import tkinter.ttk as ttk

import customtkinter as ctk
import pandas as pd
from PIL import Image as PILImage

from core.config import get_layout_dir, get_output_dir, get_template_dir, load_config, save_config
from core.data_model import EditableDataModel
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


def _sort_by_attendance(df: pd.DataFrame) -> pd.DataFrame:
    """出席番号で数値ソートして返す。元の DataFrame インデックスは保持する。"""
    if '出席番号' not in df.columns:
        return df
    result = df.copy()
    result['_sort_key'] = pd.to_numeric(
        result['出席番号'], errors='coerce',
    ).fillna(0)
    result = result.sort_values('_sort_key').drop(columns='_sort_key')
    return result


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

        # レイアウトライブラリが空なら同梱 .lay から自動インポート
        from core.layout_registry import ensure_default_layouts
        ensure_default_layouts(get_layout_dir(self.config))

        # データ状態
        self.df_mapped: pd.DataFrame | None = None
        self._data_model: EditableDataModel | None = None

        self._build_layout()
        self._check_update_bg()
        self._sync_data_bg()
        # 起動時に前回ファイルを自動再読込
        self.after(100, self._on_startup)

    # ────────────────────────────────────────────────────────────────────────
    # レイアウト構築
    # ────────────────────────────────────────────────────────────────────────

    def _build_layout(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # ── ヘッダーバー ────────────────────────────────────────────────
        header_bar = ctk.CTkFrame(self, height=36, corner_radius=0, fg_color='transparent')
        header_bar.grid(row=0, column=0, sticky='ew', padx=5, pady=(5, 0))
        header_bar.grid_columnconfigure(0, weight=1)

        ctk.CTkButton(
            header_bar, text='レイアウト編集', width=120, height=28,
            command=self._open_editor,
        ).grid(row=0, column=1, padx=5, sticky='e')

        ctk.CTkButton(
            header_bar, text='設定', width=60, height=28,
            command=self._open_settings,
        ).grid(row=0, column=2, padx=5, sticky='e')

        # ── メインタブ ──────────────────────────────────────────────────
        self._main_tabview = ctk.CTkTabview(
            self, corner_radius=6, command=self._on_tab_change,
        )
        self._main_tabview.grid(row=1, column=0, sticky='nsew', padx=5, pady=5)

        tab_print = self._main_tabview.add('名簿印刷')
        tab_excel = self._main_tabview.add('帳票生成')

        # ── 帳票生成タブ（既存の左右パネル） ──────────────────────────
        tab_excel.grid_columnconfigure(0, weight=0, minsize=320)
        tab_excel.grid_columnconfigure(1, weight=1)
        tab_excel.grid_rowconfigure(0, weight=1)

        # 左パネル（スクロール可）
        self._left = ctk.CTkScrollableFrame(tab_excel, width=310, corner_radius=0)
        self._left.grid(row=0, column=0, sticky='nsew', padx=(0, 2))
        self._left.grid_columnconfigure(0, weight=1)

        # 右パネル
        self._right = _PreviewPanel(tab_excel, on_data_edit=self._on_data_edited)
        self._right.grid(row=0, column=1, sticky='nsew', padx=(2, 0))

        # セクション配置
        self.import_frame = ImportFrame(self._left, on_load=self._on_import)
        self.import_frame.grid(row=0, column=0, sticky='ew', padx=5, pady=(5, 3))

        self.class_select = ClassSelectPanel(
            self._left, on_select=self._on_class_select,
        )
        self.class_select.grid(row=1, column=0, sticky='ew', padx=5, pady=3)

        self.select_frame = SelectFrame(
            self._left,
            config=self.config,
            template_dir=get_template_dir(self.config),
            on_teacher_save=self._on_teacher_save,
            on_school_name_save=self._on_school_name_save,
            on_template_change=self._request_preview,
            on_exchange_class_edit=self._show_exchange_class_dialog,
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

        # ── 名簿印刷タブ ────────────────────────────────────────────────
        tab_print.grid_columnconfigure(0, weight=1)
        tab_print.grid_rowconfigure(0, weight=1)

        from gui.frames.roster_print_panel import RosterPrintPanel
        self._roster_print = RosterPrintPanel(
            tab_print, self.config, on_import=self._on_import,
        )
        self._roster_print.grid(row=0, column=0, sticky='nsew')

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

        self._data_model = EditableDataModel(df_mapped)
        self.df_mapped = self._data_model.df

        # パスを config に保存
        self.config['last_loaded_file'] = source_path
        save_config(self.config)

        # ClassSelectPanel にデータをセット（自動でクラス選択・コールバックが走る）
        self.class_select.set_data(df_mapped)

        # 特別支援学級がある場合はオプションを表示 + 未割り当てダイアログ
        has_sn = self.class_select.has_special_needs()
        self.select_frame.set_special_needs_visible(has_sn)

        if has_sn:
            from core.special_needs import (
                detect_special_needs_students,
                get_unassigned_students,
            )
            special_df = detect_special_needs_students(df_mapped)
            assignments = self.config.get('special_needs_assignments', {})
            unassigned = get_unassigned_students(special_df, assignments)
            if not unassigned.empty:
                self._show_exchange_class_dialog()

        # プレビュー更新
        self._right.show_data(df_mapped, model=self._data_model)

        # 名簿印刷タブにもデータを共有
        self._roster_print.set_data(df_mapped)

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

        # プレビューをフィルタリングして更新（割り当て済み特支児童を自動含む）
        df_preview = self._filter_df(g, k)
        self._right.show_data(df_preview, model=self._data_model)

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

    def _show_exchange_class_dialog(self) -> None:
        """交流学級割り当てダイアログを表示する。"""
        if self.df_mapped is None:
            return

        from core.special_needs import detect_special_needs_students
        from gui.dialogs.exchange_class_dialog import ExchangeClassDialog

        special_df = detect_special_needs_students(self.df_mapped)
        if special_df.empty:
            return

        regular_classes = self.class_select.get_available_classes()
        existing = self.config.get('special_needs_assignments', {})

        def _on_confirm(assignments: dict[str, str]) -> None:
            self.config['special_needs_assignments'] = assignments
            save_config(self.config)
            # プレビューを更新（割り当て反映）
            f = self.class_select.get_filter()
            df_preview = self._filter_df(f.get('学年'), f.get('組'))
            self._right.show_data(df_preview, model=self._data_model)
            self._request_preview()

        ExchangeClassDialog(
            self, special_df, regular_classes, existing, _on_confirm,
        )

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

    def _on_data_edited(self) -> None:
        """_PreviewPanel でデータが編集された時に呼ばれる。"""
        self._request_preview()

    # ────────────────────────────────────────────────────────────────────────
    # ヘルパー
    # ────────────────────────────────────────────────────────────────────────

    def _filter_df(
        self,
        学年: str | None,
        組: str | None,
    ) -> pd.DataFrame:
        """df_mapped を学年・組でフィルタリングし、出席番号順にソートして返す。

        config の ``special_needs_assignments`` に基づいて、割り当て済みの
        特支児童を自動的に含める。
        """
        from core.special_needs import (
            detect_special_needs_students,
            get_assigned_students,
            merge_special_needs_students,
        )

        df = self.df_mapped.copy()
        if 学年 and '学年' in df.columns:
            df = df[df['学年'] == 学年]

        if 組 and '組' in df.columns:
            # 特定クラスでフィルタ
            df_regular = df[df['組'] == 組].copy()
            df_regular = _sort_by_attendance(df_regular)

            # 割り当て済み特支児童を自動で取得
            assignments = self.config.get('special_needs_assignments', {})
            if assignments:
                target_class = f'{学年}-{組}'
                special_all = detect_special_needs_students(self.df_mapped)
                assigned = get_assigned_students(
                    special_all, assignments, target_class,
                )
                if not assigned.empty:
                    placement = self.config.get(
                        'special_needs_placement', 'appended',
                    )
                    return merge_special_needs_students(
                        df_regular, assigned, placement=placement,
                    )
            return df_regular

        # 学年全体/全校: 全員を含む（特支含む）
        return _sort_by_attendance(df)

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

    def _open_editor(self) -> None:
        """レイアウトライブラリを開く。"""
        from gui.editor.layout_manager_dialog import LayoutManagerDialog
        LayoutManagerDialog(
            self, self.config, on_open=self._open_editor_with_file,
        )

    def _open_editor_with_file(self, path: str) -> None:
        """指定ファイルでレイアウトエディターを開く。空パスなら新規。"""
        from gui.editor.editor_window import EditorWindow
        layout_dir = get_layout_dir(self.config)
        if path:
            if path.lower().endswith('.lay'):
                from core.lay_parser import parse_lay
                EditorWindow(self, lay=parse_lay(path), layout_dir=layout_dir)
            else:
                from core.lay_serializer import load_layout
                ew = EditorWindow(
                    self, lay=load_layout(path), layout_dir=layout_dir,
                )
                ew._current_file = path
        else:
            EditorWindow(self, layout_dir=layout_dir)

    def _on_tab_change(self) -> None:
        """タブ切り替え時にデータとレイアウト一覧を同期する。"""
        if self._main_tabview.get() == '名簿印刷':
            self._roster_print.refresh_layouts()
            # 帳票生成タブで読み込み済みのデータを共有
            if self.df_mapped is not None:
                self._roster_print.set_data(self.df_mapped)

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

        # フィルタリング: 現在の学年・組から先頭 N 名
        f = self.class_select.get_filter()
        df_filtered = self._filter_df(f.get('学年'), f.get('組'))
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
    """右パネル: データ一覧（編集可能） + 帳票プレビュー（CTkTabview）。"""

    _PRIORITY_COLS = ('出席番号', '組', '学年', '氏名', '氏名かな', '性別', '生年月日')

    def __init__(self, master, on_data_edit: callable | None = None) -> None:
        super().__init__(master, corner_radius=6)
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self._on_data_edit = on_data_edit
        self._model: EditableDataModel | None = None
        self._show_cols: list[str] = []
        self._tree: ttk.Treeview | None = None
        self._edit_entry: tk.Entry | None = None
        self._edit_item: str | None = None
        self._edit_col_idx: int = -1

        # タブビュー
        self._tabview = ctk.CTkTabview(self, corner_radius=6)
        self._tabview.grid(row=0, column=0, sticky='nsew', padx=4, pady=4)

        self._tab_data = self._tabview.add('データ')
        self._tab_preview = self._tabview.add('プレビュー')

        # ── データタブ ─────────────────────────────────────────────────────
        self._tab_data.grid_rowconfigure(2, weight=1)
        self._tab_data.grid_columnconfigure(0, weight=1)

        # ツールバー（ヘッダー + 検索 + エクスポート）
        toolbar = ctk.CTkFrame(self._tab_data, fg_color='transparent')
        toolbar.grid(row=0, column=0, sticky='ew', padx=4, pady=(4, 0))
        toolbar.grid_columnconfigure(1, weight=1)

        self._header = ctk.CTkLabel(
            toolbar,
            text='データ編集',
            font=ctk.CTkFont(size=13, weight='bold'),
        )
        self._header.grid(row=0, column=0, sticky='w', padx=2)

        self._export_btn = ctk.CTkButton(
            toolbar, text='エクスポート', width=90, height=26,
            command=self._on_export, state='disabled',
        )
        self._export_btn.grid(row=0, column=2, padx=(4, 2))

        # 検索バー
        search_frame = ctk.CTkFrame(self._tab_data, fg_color='transparent')
        search_frame.grid(row=1, column=0, sticky='ew', padx=4, pady=(2, 0))
        search_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(search_frame, text='検索:', font=ctk.CTkFont(size=11)).grid(
            row=0, column=0, padx=(2, 4),
        )
        self._search_var = tk.StringVar()
        self._search_entry = ctk.CTkEntry(
            search_frame, textvariable=self._search_var,
            placeholder_text='氏名・かな等で絞り込み…', height=26,
        )
        self._search_entry.grid(row=0, column=1, sticky='ew', padx=2)
        self._search_var.trace_add('write', self._on_search_changed)

        self._current_df: pd.DataFrame | None = None

        self._data_placeholder = ctk.CTkLabel(
            self._tab_data,
            text='名簿ファイルを読み込むと\nここにプレビューが表示されます',
            text_color='gray',
        )
        self._data_placeholder.grid(row=2, column=0, padx=20, pady=40)

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

    def show_data(
        self, df: pd.DataFrame, *, model: EditableDataModel | None = None,
    ) -> None:
        """DataFrame を Treeview に表示する。model を渡すとセル編集が有効になる。"""
        self._cancel_edit()
        if self._tree_container:
            self._tree_container.destroy()
            self._tree_container = None
        self._data_placeholder.grid_remove()
        self._model = model
        self._current_df = df
        self._search_var.set('')
        self._export_btn.configure(state='normal')

        show_cols = [c for c in self._PRIORITY_COLS if c in df.columns]
        if not show_cols:
            show_cols = list(df.columns[:10])
        self._show_cols = show_cols

        container = ctk.CTkFrame(self._tab_data, corner_radius=0, fg_color='transparent')
        container.grid(row=2, column=0, sticky='nsew', padx=2, pady=2)
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
        tree.tag_configure('edited', background='#FFFACD')
        tree.tag_configure('warning', background='#FFE4E1')

        # 必須フィールド（空欄の場合に警告ハイライトする）
        required_fields = {'氏名', '性別', '生年月日'}

        for col in show_cols:
            w = 70 if col in ('出席番号', '組', '学年', '性別') else 110
            tree.heading(col, text=col)
            tree.column(col, width=w, minwidth=50, anchor='center')

        for idx, row in df.iterrows():
            values = [_fmt_cell(c, row.get(c, '')) for c in show_cols]
            # 必須フィールドに空欄があれば警告タグ
            tags: tuple[str, ...] = ()
            for i, col in enumerate(show_cols):
                if col in required_fields and not values[i]:
                    tags = ('warning',)
                    break
            tree.insert('', 'end', iid=str(idx), values=values, tags=tags)

        vsb = ttk.Scrollbar(container, orient='vertical', command=tree.yview)
        hsb = ttk.Scrollbar(container, orient='horizontal', command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')

        self._tree = tree

        # ダブルクリックでセル編集
        if model is not None:
            tree.bind('<Double-1>', self._on_double_click)
            tree.bind('<Control-z>', lambda _e: self._undo())
            tree.bind('<Control-y>', lambda _e: self._redo())

        self._update_header(len(df))

    # ── インライン編集 ─────────────────────────────────────────────────────

    def _update_header(self, count: int, *, filtered: int | None = None) -> None:
        """データタブのヘッダーを更新する。"""
        if filtered is not None:
            text = f'データ編集 — {filtered}/{count} 名'
        else:
            text = f'データ編集 — {count} 名'
        if self._model and self._model.is_modified():
            text += '  (*変更あり)'
        self._header.configure(text=text)

    def _on_double_click(self, event: tk.Event) -> None:
        """Treeview セルのダブルクリックで編集を開始する。"""
        tree = self._tree
        if tree is None:
            return
        item = tree.identify_row(event.y)
        col_id = tree.identify_column(event.x)
        if not item or not col_id:
            return
        # col_id は '#1', '#2', ... 形式
        col_idx = int(col_id.replace('#', '')) - 1
        if col_idx < 0 or col_idx >= len(self._show_cols):
            return
        self._start_edit(item, col_idx)

    def _start_edit(self, item: str, col_idx: int) -> None:
        """指定セルの編集を開始する。"""
        tree = self._tree
        if tree is None or self._model is None:
            return

        self._cancel_edit()

        col_name = self._show_cols[col_idx]
        row_idx = int(item)
        current_value = self._model.get_value(row_idx, col_name)

        # セルの位置を取得
        bbox = tree.bbox(item, column=self._show_cols[col_idx])
        if not bbox:
            return
        x, y, w, h = bbox

        entry = tk.Entry(tree, font=('メイリオ', 9), justify='center')
        entry.place(x=x, y=y, width=w, height=h)
        entry.insert(0, current_value)
        entry.select_range(0, 'end')
        entry.focus_set()

        entry.bind('<Return>', lambda _e: self._commit_edit())
        entry.bind('<Escape>', lambda _e: self._cancel_edit())
        entry.bind('<FocusOut>', lambda _e: self._commit_edit())

        self._edit_entry = entry
        self._edit_item = item
        self._edit_col_idx = col_idx

    def _commit_edit(self) -> None:
        """編集を確定する。"""
        if self._edit_entry is None or self._model is None:
            return

        new_value = self._edit_entry.get()
        item = self._edit_item
        col_idx = self._edit_col_idx
        col_name = self._show_cols[col_idx]
        row_idx = int(item)

        # Entry を先に破棄（FocusOut の再帰防止）
        entry = self._edit_entry
        self._edit_entry = None
        self._edit_item = None
        entry.destroy()

        # モデルに反映
        self._model.set_value(row_idx, col_name, new_value)

        # Treeview の表示を更新
        if self._tree is not None:
            values = list(self._tree.item(str(row_idx), 'values'))
            values[col_idx] = _fmt_cell(col_name, new_value)
            self._tree.item(str(row_idx), values=values, tags=('edited',))

        # ヘッダー更新
        tree = self._tree
        if tree:
            self._update_header(len(tree.get_children()))

        # コールバック（帳票プレビュー更新等）
        if self._on_data_edit:
            self._on_data_edit()

    def _cancel_edit(self) -> None:
        """編集をキャンセルする。"""
        if self._edit_entry is not None:
            entry = self._edit_entry
            self._edit_entry = None
            self._edit_item = None
            entry.destroy()

    def _undo(self) -> None:
        """Undo: 直前の編集を取り消す。"""
        if self._model is None:
            return
        op = self._model.undo()
        if op is None:
            return
        self._refresh_cell(op.row, op.col, op.old_value)
        if self._on_data_edit:
            self._on_data_edit()

    def _redo(self) -> None:
        """Redo: 取り消した編集をやり直す。"""
        if self._model is None:
            return
        op = self._model.redo()
        if op is None:
            return
        self._refresh_cell(op.row, op.col, op.new_value)
        if self._on_data_edit:
            self._on_data_edit()

    def _refresh_cell(self, row_idx: int, col_name: str, value: str) -> None:
        """Treeview の特定セルを更新する。"""
        tree = self._tree
        if tree is None:
            return
        item = str(row_idx)
        if not tree.exists(item):
            return
        col_idx = self._show_cols.index(col_name) if col_name in self._show_cols else -1
        if col_idx < 0:
            return
        values = list(tree.item(item, 'values'))
        values[col_idx] = _fmt_cell(col_name, value)
        tags = ('edited',) if self._model and self._model.is_modified() else ()
        tree.item(item, values=values, tags=tags)
        self._update_header(len(tree.get_children()))

    # ── 検索・フィルタ ─────────────────────────────────────────────────────

    def _on_search_changed(self, *_args: object) -> None:
        """検索文字列の変更時にフィルタを適用する。"""
        tree = self._tree
        if tree is None or self._current_df is None:
            return

        query = self._search_var.get().strip().lower()

        for item in tree.get_children():
            tree.delete(item)

        for idx, row in self._current_df.iterrows():
            values = [_fmt_cell(c, row.get(c, '')) for c in self._show_cols]
            if query and not any(query in str(v).lower() for v in values):
                continue
            tree.insert('', 'end', iid=str(idx), values=values)

        visible = len(tree.get_children())
        total = len(self._current_df)
        if query:
            self._update_header(total, filtered=visible)
        else:
            self._update_header(total)

    # ── エクスポート ───────────────────────────────────────────────────────

    def _on_export(self) -> None:
        """編集済みデータをファイルにエクスポートする。"""
        import tkinter.filedialog as fd

        if self._model is None and self._current_df is None:
            return

        df = self._model.get_df() if self._model else self._current_df

        path = fd.asksaveasfilename(
            title='データをエクスポート',
            defaultextension='.csv',
            filetypes=[
                ('CSV (UTF-8)', '*.csv'),
                ('Excel', '*.xlsx'),
            ],
        )
        if not path:
            return

        from core.exporter import export_csv, export_excel
        try:
            if path.lower().endswith('.xlsx'):
                export_excel(df, path)
            else:
                export_csv(df, path)
            if self._model:
                self._model.reset_modified()
                self._update_header(len(df))
        except Exception as e:
            import tkinter.messagebox as _mb
            _mb.showerror('エクスポートエラー', str(e))

    # ── プレビュータブ API ─────────────────────────────────────────────────

    def show_preview_loading(self) -> None:
        """プレビュー生成中のローディング表示。"""
        self._preview_image_ref = None
        with contextlib.suppress(tk.TclError):
            self._preview_label.configure(
                text='プレビューを生成中…', text_color='gray', image=None,
            )
            self._preview_label.grid(row=0, column=0, padx=20, pady=40)

    def show_preview_image(self, pil_image: PILImage.Image) -> None:
        """PIL Image をプレビュータブに表示する。"""
        try:
            ctk_img = ctk.CTkImage(
                light_image=pil_image,
                size=(pil_image.width, pil_image.height),
            )
            self._preview_image_ref = ctk_img  # GC 防止
            self._preview_label.configure(
                text='', image=ctk_img, text_color='gray',
            )
            self._preview_label.grid(row=0, column=0, padx=4, pady=4)
            # プレビュータブに自動切替
            self._tabview.set('プレビュー')
        except tk.TclError:
            pass

    def show_preview_error(self, msg: str) -> None:
        """プレビューエラーメッセージを表示する。"""
        self._preview_image_ref = None
        with contextlib.suppress(tk.TclError):
            self._preview_label.configure(
                text=msg, text_color='#CC4444', image=None,
            )
            self._preview_label.grid(row=0, column=0, padx=20, pady=40)
