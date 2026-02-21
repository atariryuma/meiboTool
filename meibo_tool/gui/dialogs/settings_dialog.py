"""管理者設定ダイアログ

データソース設定 (LAN / Google Drive / 手動) と暗号化操作を提供する。
"""

from __future__ import annotations

import tkinter.filedialog as fd
import tkinter.messagebox as mb
from typing import Any

import customtkinter as ctk

from core.config import save_config
from core.crypto import encrypt_file, protect_password, unprotect_password


class SettingsDialog(ctk.CTkToplevel):
    """管理者設定ダイアログ。"""

    def __init__(self, master: ctk.CTk, config: dict[str, Any]) -> None:
        super().__init__(master)
        self.title('設定')
        self.geometry('520x520')
        self.resizable(False, False)
        self.transient(master)
        self.grab_set()

        self._config = config
        self._ds = dict(config.get('data_source', {}))
        self._update = dict(config.get('update', {}))

        self._build_ui()
        self._load_values()

    # ────────────────────────────────────────────────────────────────────────
    # UI 構築
    # ────────────────────────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)

        # ── データソース設定 ──────────────────────────────────────────────
        ds_frame = ctk.CTkFrame(self, corner_radius=6)
        ds_frame.grid(row=0, column=0, sticky='ew', padx=12, pady=(12, 6))
        ds_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            ds_frame, text='データソース設定',
            font=ctk.CTkFont(size=14, weight='bold'),
        ).grid(row=0, column=0, columnspan=3, sticky='w', padx=10, pady=(8, 6))

        # モード選択
        ctk.CTkLabel(ds_frame, text='モード:').grid(
            row=1, column=0, sticky='w', padx=(10, 5), pady=4,
        )
        self._mode_var = ctk.StringVar(value='manual')
        mode_frame = ctk.CTkFrame(ds_frame, fg_color='transparent')
        mode_frame.grid(row=1, column=1, columnspan=2, sticky='w', pady=4)

        for i, (val, label) in enumerate([
            ('manual', '手動'),
            ('lan', 'LAN'),
            ('gdrive', 'Google Drive'),
        ]):
            ctk.CTkRadioButton(
                mode_frame, text=label, variable=self._mode_var,
                value=val, command=self._on_mode_change,
            ).grid(row=0, column=i, padx=(0, 12))

        # LAN パス
        self._lan_label = ctk.CTkLabel(ds_frame, text='共有フォルダ:')
        self._lan_label.grid(row=2, column=0, sticky='w', padx=(10, 5), pady=4)
        self._lan_entry = ctk.CTkEntry(ds_frame, placeholder_text='\\\\server\\共有\\名簿.xlsx')
        self._lan_entry.grid(row=2, column=1, sticky='ew', padx=2, pady=4)
        self._lan_browse = ctk.CTkButton(
            ds_frame, text='参照...', width=60, command=self._on_browse_lan,
        )
        self._lan_browse.grid(row=2, column=2, padx=(2, 10), pady=4)

        # Google Drive ファイル ID
        self._gd_id_label = ctk.CTkLabel(ds_frame, text='ファイル ID:')
        self._gd_id_label.grid(row=3, column=0, sticky='w', padx=(10, 5), pady=4)
        self._gd_id_entry = ctk.CTkEntry(ds_frame, placeholder_text='Google Drive ファイル ID')
        self._gd_id_entry.grid(row=3, column=1, columnspan=2, sticky='ew', padx=(2, 10), pady=4)

        # パスワード
        self._pw_label = ctk.CTkLabel(ds_frame, text='パスワード:')
        self._pw_label.grid(row=4, column=0, sticky='w', padx=(10, 5), pady=4)
        self._pw_entry = ctk.CTkEntry(ds_frame, show='*', placeholder_text='暗号化パスワード')
        self._pw_entry.grid(row=4, column=1, columnspan=2, sticky='ew', padx=(2, 10), pady=4)

        # 暗号化ボタン
        self._encrypt_btn = ctk.CTkButton(
            ds_frame, text='名簿ファイルを暗号化...', command=self._on_encrypt_click,
        )
        self._encrypt_btn.grid(row=5, column=0, columnspan=3, padx=10, pady=(4, 10))

        # ── GitHub 更新設定 ──────────────────────────────────────────────
        upd_frame = ctk.CTkFrame(self, corner_radius=6)
        upd_frame.grid(row=1, column=0, sticky='ew', padx=12, pady=6)
        upd_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            upd_frame, text='アプリ更新設定',
            font=ctk.CTkFont(size=14, weight='bold'),
        ).grid(row=0, column=0, columnspan=2, sticky='w', padx=10, pady=(8, 6))

        ctk.CTkLabel(upd_frame, text='リポジトリ:').grid(
            row=1, column=0, sticky='w', padx=(10, 5), pady=4,
        )
        self._repo_entry = ctk.CTkEntry(upd_frame, placeholder_text='owner/repo')
        self._repo_entry.grid(row=1, column=1, sticky='ew', padx=(2, 10), pady=4)

        self._auto_check_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            upd_frame, text='起動時に自動チェック', variable=self._auto_check_var,
        ).grid(row=2, column=0, columnspan=2, sticky='w', padx=10, pady=(4, 10))

        # ── ボタン ───────────────────────────────────────────────────────
        btn_frame = ctk.CTkFrame(self, fg_color='transparent')
        btn_frame.grid(row=2, column=0, sticky='ew', padx=12, pady=(6, 12))
        btn_frame.grid_columnconfigure((0, 1), weight=1)

        ctk.CTkButton(
            btn_frame, text='保存', command=self._on_save, width=120,
        ).grid(row=0, column=0, padx=5, pady=5, sticky='e')
        ctk.CTkButton(
            btn_frame, text='キャンセル', command=self.destroy,
            width=120, fg_color='gray',
        ).grid(row=0, column=1, padx=5, pady=5, sticky='w')

    # ────────────────────────────────────────────────────────────────────────
    # 値の読み込み
    # ────────────────────────────────────────────────────────────────────────

    def _load_values(self) -> None:
        self._mode_var.set(self._ds.get('mode', 'manual'))

        lan_path = self._ds.get('lan_path', '')
        if lan_path:
            self._lan_entry.insert(0, lan_path)

        gd_id = self._ds.get('gdrive_file_id', '')
        if gd_id:
            self._gd_id_entry.insert(0, gd_id)

        pw = unprotect_password(self._ds.get('encryption_password', ''))
        if pw:
            self._pw_entry.insert(0, pw)

        repo = self._update.get('github_repo', '')
        if repo:
            self._repo_entry.insert(0, repo)

        self._auto_check_var.set(self._update.get('check_on_startup', True))
        self._on_mode_change()

    # ────────────────────────────────────────────────────────────────────────
    # イベントハンドラ
    # ────────────────────────────────────────────────────────────────────────

    def _on_mode_change(self) -> None:
        """モード変更時に表示/非表示を切り替える。"""
        mode = self._mode_var.get()
        is_lan = mode == 'lan'
        is_gdrive = mode == 'gdrive'

        for w in (self._lan_label, self._lan_entry, self._lan_browse):
            if is_lan:
                w.grid()
            else:
                w.grid_remove()

        for w in (self._gd_id_label, self._gd_id_entry, self._pw_label, self._pw_entry, self._encrypt_btn):
            if is_gdrive:
                w.grid()
            else:
                w.grid_remove()

    def _on_browse_lan(self) -> None:
        """LAN パスのファイル選択。"""
        path = fd.askopenfilename(
            title='名簿ファイルを選択',
            filetypes=[
                ('名簿ファイル', '*.xlsx *.xls *.csv'),
                ('Excel', '*.xlsx *.xls'),
                ('CSV', '*.csv'),
                ('すべて', '*.*'),
            ],
        )
        if path:
            self._lan_entry.delete(0, 'end')
            self._lan_entry.insert(0, path)

    def _on_encrypt_click(self) -> None:
        """名簿ファイルを暗号化する。"""
        password = self._pw_entry.get().strip()
        if not password:
            mb.showwarning('パスワード未入力', 'パスワードを入力してください。', parent=self)
            return

        source = fd.askopenfilename(
            title='暗号化する名簿ファイルを選択',
            filetypes=[
                ('名簿ファイル', '*.xlsx *.xls *.csv'),
                ('Excel', '*.xlsx *.xls'),
                ('CSV', '*.csv'),
            ],
        )
        if not source:
            return

        dest = source + '.encrypted'
        try:
            encrypt_file(source, dest, password)
        except Exception as exc:
            mb.showerror('暗号化エラー', f'暗号化に失敗しました:\n{exc}', parent=self)
            return

        mb.showinfo(
            '暗号化完了',
            f'暗号化ファイルを作成しました:\n{dest}\n\n'
            'このファイルを Google Drive にアップロードし、\n'
            '共有設定を「リンクを知っている全員」に変更してください。',
            parent=self,
        )

    def _on_save(self) -> None:
        """設定を保存する。"""
        mode = self._mode_var.get()

        # LAN モード時は共有フォルダパス必須
        lan_path = self._lan_entry.get().strip()
        if mode == 'lan' and not lan_path:
            mb.showwarning(
                'パス未入力',
                'LAN モードでは共有フォルダのパスを入力してください。',
                parent=self,
            )
            return

        # Google Drive モード時はファイル ID 必須
        gd_id = self._gd_id_entry.get().strip()
        if mode == 'gdrive' and not gd_id:
            mb.showwarning(
                'ファイル ID 未入力',
                'Google Drive モードではファイル ID を入力してください。',
                parent=self,
            )
            return

        self._ds['mode'] = mode
        self._ds['lan_path'] = lan_path
        self._ds['gdrive_file_id'] = gd_id

        pw = self._pw_entry.get().strip()
        if pw:
            self._ds['encryption_password'] = protect_password(pw)
        else:
            self._ds['encryption_password'] = ''

        self._update['github_repo'] = self._repo_entry.get().strip()
        self._update['check_on_startup'] = self._auto_check_var.get()

        self._config['data_source'] = self._ds
        self._config['update'] = self._update
        save_config(self._config)

        self.destroy()
