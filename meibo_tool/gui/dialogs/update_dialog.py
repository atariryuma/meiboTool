"""アプリ更新確認ダイアログ

UpdateInfo を受け取り、ユーザーに更新を確認する。
ダウンロード進捗をプログレスバーで表示する。
"""

from __future__ import annotations

import contextlib
import os
import sys
import threading
import tkinter.messagebox as mb
from typing import Any

import customtkinter as ctk

from core.config import save_config
from core.updater import (
    UpdateInfo,
    compute_local_manifest,
    diff_manifests,
    download_manifest,
    download_release_asset,
    extract_changed_files,
    generate_update_batch,
)


class UpdateDialog(ctk.CTkToplevel):
    """更新確認ダイアログ。"""

    def __init__(
        self,
        master: ctk.CTk,
        update_info: UpdateInfo,
        config: dict[str, Any],
    ) -> None:
        super().__init__(master)
        self.title('更新のお知らせ')
        self.geometry('440x320')
        self.resizable(False, False)
        self.transient(master)
        self.grab_set()

        self._info = update_info
        self._config = config
        self._master = master
        self._build_ui()

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)

        # バージョン情報
        ctk.CTkLabel(
            self,
            text='新しいバージョンが利用可能です',
            font=ctk.CTkFont(size=15, weight='bold'),
        ).grid(row=0, column=0, padx=20, pady=(16, 4))

        ctk.CTkLabel(
            self,
            text=f'{self._info.current_version} → {self._info.new_version}',
            font=ctk.CTkFont(size=13),
            text_color='#2563EB',
        ).grid(row=1, column=0, padx=20, pady=2)

        # リリースノート
        if self._info.release_notes:
            notes_box = ctk.CTkTextbox(self, height=100, corner_radius=6)
            notes_box.grid(row=2, column=0, sticky='ew', padx=20, pady=8)
            notes_box.insert('1.0', self._info.release_notes)
            notes_box.configure(state='disabled')

        # プログレスバー（初期非表示）
        self._progress = ctk.CTkProgressBar(self, mode='determinate')
        self._progress.set(0)
        self._progress_label = ctk.CTkLabel(self, text='', font=ctk.CTkFont(size=11))

        # ボタン
        btn_frame = ctk.CTkFrame(self, fg_color='transparent')
        btn_frame.grid(row=5, column=0, sticky='ew', padx=20, pady=(4, 16))
        btn_frame.grid_columnconfigure((0, 1, 2), weight=1)

        self._update_btn = ctk.CTkButton(
            btn_frame, text='今すぐ更新', command=self._on_update, width=110,
        )
        self._update_btn.grid(row=0, column=0, padx=4, pady=4)

        ctk.CTkButton(
            btn_frame, text='後で', command=self.destroy,
            width=90, fg_color='gray',
        ).grid(row=0, column=1, padx=4, pady=4)

        self._skip_btn = ctk.CTkButton(
            btn_frame, text='スキップ', command=self._on_skip,
            width=90, fg_color='gray',
        )
        self._skip_btn.grid(row=0, column=2, padx=4, pady=4)

    def _on_update(self) -> None:
        """バックグラウンドでダウンロードを開始する。"""
        self._update_btn.configure(state='disabled', text='ダウンロード中...')
        self._skip_btn.configure(state='disabled')

        # プログレスバー表示
        self._progress.grid(row=3, column=0, sticky='ew', padx=20, pady=(4, 2))
        self._progress_label.grid(row=4, column=0, padx=20, pady=0)

        threading.Thread(target=self._download_task, daemon=True).start()

    def _download_task(self) -> None:
        """バックグラウンドスレッドでダウンロードを実行する。

        2 フェーズ方式:
          1. manifest.json でローカルと比較 → 変更なしなら終了
          2. 変更ありなら zip DL → 変更分のみ展開 → バッチ実行
        manifest がないリリースではフル置換にフォールバックする。
        """
        try:
            if getattr(sys, 'frozen', False):
                current_dir = os.path.dirname(sys.executable)
            else:
                current_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

            parent_dir = os.path.dirname(current_dir)
            changed_files: list[str] | None = None
            staging_dir: str | None = None

            # ── フェーズ 1: マニフェスト比較 ──
            if self._info.manifest_url:
                self.after(0, lambda: self._progress_label.configure(
                    text='変更内容を確認中…',
                ))
                try:
                    remote_manifest = download_manifest(self._info.manifest_url)
                    remote_files = remote_manifest.get('files', {})
                    local_files = compute_local_manifest(current_dir)
                    changed_files = diff_manifests(remote_files, local_files)

                    if not changed_files:
                        self.after(0, self._no_changes_found)
                        return

                    n = len(changed_files)
                    self.after(0, lambda: self._progress_label.configure(
                        text=f'{n} ファイルが変更されました。ダウンロード中…',
                    ))
                except Exception:
                    # manifest 取得/比較に失敗 → フル置換にフォールバック
                    changed_files = None
                    self.after(0, lambda: self._progress_label.configure(
                        text='ダウンロード中…',
                    ))

            # ── フェーズ 2: zip ダウンロード → 展開 ──
            dest = os.path.join(parent_dir, self._info.asset_name)

            def _progress_cb(fraction: float) -> None:
                self.after(0, lambda: self._update_progress(fraction))

            download_release_asset(self._info.asset_url, dest, _progress_cb)

            # 差分展開
            if changed_files is not None:
                staging_dir = os.path.join(parent_dir, '_update_staging')
                extract_changed_files(dest, staging_dir, changed_files)

            if getattr(sys, 'frozen', False):
                self.after(0, lambda: self._apply_update(
                    dest, current_dir, changed_files, staging_dir,
                ))
            else:
                self.after(0, lambda: self._download_complete_dev(dest))

        except Exception as exc:
            err_msg = str(exc)
            self.after(0, lambda: self._download_failed(err_msg))

    def _update_progress(self, fraction: float) -> None:
        """プログレスバーを更新する。"""
        self._progress.set(fraction)
        pct = int(fraction * 100)
        self._progress_label.configure(text=f'{pct}%')

    def _no_changes_found(self) -> None:
        """マニフェスト比較で変更なし。"""
        mb.showinfo(
            '更新チェック',
            'すべてのファイルが最新です。更新の必要はありません。',
            parent=self,
        )
        self.destroy()

    def _apply_update(
        self,
        zip_path: str,
        current_dir: str,
        changed_files: list[str] | None = None,
        staging_dir: str | None = None,
    ) -> None:
        """ダウンロード完了後、バッチファイルで更新を適用する。"""
        if changed_files is not None:
            msg = f'ダウンロードが完了しました（{len(changed_files)} ファイル更新）。\nアプリを再起動して更新を適用しますか？'
        else:
            msg = 'ダウンロードが完了しました。\nアプリを再起動して更新を適用しますか？'

        if mb.askyesno('更新の適用', msg, parent=self):
            generate_update_batch(
                zip_path, current_dir,
                changed_files=changed_files,
                staging_dir=staging_dir,
            )
            # generate_update_batch が sys.exit(0) を呼ぶ

    def _download_complete_dev(self, zip_path: str) -> None:
        """開発環境でのダウンロード完了。"""
        mb.showinfo(
            'ダウンロード完了',
            f'ダウンロードしました:\n{zip_path}\n\n（開発環境のため自動更新は行いません）',
            parent=self,
        )
        self.destroy()

    def _download_failed(self, error: str) -> None:
        """ダウンロード失敗。"""
        mb.showerror(
            'ダウンロードエラー',
            f'ダウンロードに失敗しました:\n{error}',
            parent=self,
        )
        self._update_btn.configure(state='normal', text='今すぐ更新')
        self._skip_btn.configure(state='normal')
        self._progress.grid_remove()
        self._progress_label.grid_remove()

    def _on_skip(self) -> None:
        """このバージョンをスキップする。"""
        self._config.setdefault('update', {})['skip_version'] = self._info.new_version
        with contextlib.suppress(OSError):
            save_config(self._config)
        self.destroy()
