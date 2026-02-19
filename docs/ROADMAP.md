# 名簿帳票ツール 開発ロードマップ

> **最終更新**: 2026-02-19（品質評価に基づく全面改修 — バグ修正6件 + テスト追加41件 + 設計書同期）
> このファイルは作業完了のたびに更新する。CLAUDE.md の「現在地」セクションも合わせて更新すること。

---

## フェーズ概要

| フェーズ | 目標 | 状態 |
|---------|------|------|
| Phase 0 | 開発基盤（環境・テスト・リント） | ✅ **完了** |
| Phase 1 | MVP: 名列表を1枚出力できる | ✅ **完了**（GUI 含む） |
| Phase 2 | 名札3種 + 台帳2種 + 調べ表 + 追加帳票 | ✅ **完了** |
| Phase 3 | 自動更新 + 暗号化同期 + 品質改善 | ✅ **完了** |
| Phase 4 | 個票2種 + ビルド + 配布 | ⏳ 未着手 |

---

## Phase 0 — 開発基盤 ✅

**目標**: バグを書く前に検出できる環境を整える

- [x] `venv` + `requirements.txt`（customtkinter / openpyxl / pandas / ruff / pytest / cryptography）
- [x] `pyproject.toml`（`pythonpath`・`testpaths`・ruff ルール一元管理）
- [x] `tests/conftest.py`（`dummy_df` / `default_options` / `default_config` フィクスチャ）
- [x] `sys.path` ハック全廃
- [x] `.gitignore`（`.pytest_cache/` `.ruff_cache/` `.claude/` 追加）
- [x] Git 初期化
- [x] `docs/` 整備（SPEC.md / ROADMAP.md / images/ / .docx）
- [x] ruff クリーン・テスト全パス

---

## Phase 1 — MVP ✅

**目標**: ダミーデータを読み込んで「掲示用名列表.xlsx」を出力できる

- [x] `utils/wareki.py` — 西暦→和暦変換（令和/平成/昭和）
- [x] `utils/address.py` — 住所4フィールド結合・NaN除外
- [x] `utils/font_helper.py` — IPAmj明朝フォント適用
- [x] `core/config.py` — config.json 読み書き・deep_merge・パス解決
- [x] `core/mapper.py` — C4th 全50カラム → 内部論理名マッピング
- [x] `core/importer.py` — ヘッダー行自動検出 + dtype=str 読込
- [x] `core/generator.py` — BaseGenerator / GridGenerator / ListGenerator / IndividualGenerator
- [x] `templates/template_registry.py` — テンプレート15種メタデータ
- [x] `templates/generators/gen_meireihyo.py` — 掲示用名列表テンプレート生成
- [x] `templates/generators/generate_all.py` — 全テンプレート一括生成
- [x] `gui/app.py` — CustomTkinter メインウィンドウ（2カラムレイアウト）
- [x] `gui/frames/import_frame.py` — ファイル選択・読込件数・同期ステータス表示
- [x] `gui/frames/class_select_panel.py` — 学年・組選択パネル
- [x] `gui/frames/select_frame.py` — テンプレート選択・担任名・学校名
- [x] `gui/frames/output_frame.py` — 生成ボタン・進捗バー
- [ ] 統合確認: ダミーデータ読込 → 組設定 → 名列表.xlsx 出力まで一気通貫（手動確認未実施）

---

## Phase 2 — 名札3種 + 台帳2種 + 調べ表 + 追加帳票 ✅

**目標**: 実際の授業開始に必要な帳票をすべて出力できる

### テンプレート生成スクリプト ✅

- [x] `gen_nafuda.py` — 名札3種（通常/装飾/1年生）+ ラベル3種（大2/小/特大）
- [x] `gen_daicho.py` — 台帳2種（修了/卒業）
- [x] `gen_shirabehyo.py` — 調べ表（6列×10行グリッド）
- [x] `gen_from_legacy.py` — レガシーテンプレート変換（横名簿/縦一週間）

### ジェネレーター拡張 ✅

- [x] `GridGenerator._populate()` — ページネーション完全実装
- [x] `ListGenerator._find_template_row()` — 全名前系プレースホルダー検出
- [x] `IndividualGenerator._populate()` — fill 前にシート複製する方式に修正
- [x] `generate_all.py` — 全テンプレートを一括生成
- [x] `template_registry.py` — 15種メタデータ・`get_display_groups()` 追加

---

## Phase 3 — 自動更新 + 暗号化同期 + 品質改善 ✅

**目標**: アプリ更新・名簿データ同期の自動化

### 暗号化基盤 ✅
- [x] `core/crypto.py` — AES-256-GCM 暗号化/復号 + PBKDF2 鍵導出 + DPAPI パスワード保護

### データ同期 ✅
- [x] `core/data_sync.py` — 3モード同期（manual/lan/gdrive）+ SyncResult（config_updates 方式）

### アプリ更新 ✅
- [x] `core/updater.py` — GitHub Releases API でバージョン確認 + zip ダウンロード + バッチ更新
- [x] `gui/dialogs/update_dialog.py` — 更新確認ダイアログ（進捗バー付き）

### 管理者設定 ✅
- [x] `gui/dialogs/settings_dialog.py` — 同期モード設定・暗号化パスワード・LAN パス

### GUI 統合 ✅
- [x] `gui/app.py` — `_check_update_bg()` / `_sync_data_bg()` / 設定ボタン追加
- [x] `gui/frames/output_frame.py` — メインスレッドでジェネレーター作成（スレッド安全性修正）

### CI/CD ✅
- [x] `.github/workflows/build-release.yml` — テスト → ビルド → Release

### 品質改善（バグ修正） ✅
- [x] `generator.py` — `copy_sheet_with_images`: `Image(img.ref)` → `Image(BytesIO(img._data()))` に修正
- [x] `generator.py` — `IndividualGenerator`: fill 前にシート複製する方式に修正
- [x] `output_frame.py` — ジェネレーター作成をメインスレッドに移動（スレッド安全性）
- [x] `data_sync.py` — config 変更を `config_updates` で返し、メインスレッドで適用する方式に修正
- [x] `updater.py` — バッチファイルの `ren` → `move` に修正（パスクォート問題）
- [x] `config.py` — `fiscal_year` を関数化（インポート時凍結を修正）+ JSON パースエラー処理追加

### テスト拡充 ✅（153ケース全パス）

- [x] `test_config.py` — deep_merge / fiscal_year / load_config / malformed JSON（13ケース）
- [x] `test_importer.py` — ヘッダー検出 / インポート / 全角スペースヘッダー（11ケース）
- [x] `test_mapper_resolve.py` — resolve_name_fields フォールバック（6ケース）
- [x] `test_list_individual_gen.py` — ListGenerator / IndividualGenerator 統合（11ケース）

---

## Phase 4 — 個票2種 + ビルド + 配布

**目標**: 全テンプレート対応 + 配布可能な exe

### 個票テンプレート

- [ ] `gen_katei_chousahyo.py` — 家庭調査票（A4縦・1名/シート）
- [ ] `gen_gakkyuu_kojihyo.py` — 学級編成用個票（A4縦・1名/シート）

### GUI 追加

- [ ] `gui/dialogs/mapping_dialog.py` — カラムマッピング手動調整画面
- [ ] ステータスバー（更新通知・生成完了メッセージ）

### ビルド・配布

- [x] `build.spec` 作成（`--onedir`・CustomTkinter データファイル同梱）
- [ ] `pyinstaller build.spec` で `dist/名簿帳票ツール/` 生成確認
- [ ] 別 PC（Python なし環境）での起動確認

---

## 既知の落とし穴（実装時に必ず確認）

| 場所 | 落とし穴 | 対策 |
|------|---------|------|
| openpyxl | `cell.font.bold = True` → AttributeError | `cell.font = Font(bold=True)` で新インスタンス |
| openpyxl | `copy_worksheet` は画像をコピーしない | `_data()` → `BytesIO` → `Image()` で再構築 |
| openpyxl | 結合セルの左上以外に書き込み → AttributeError | `MergedCell` チェックを入れる |
| openpyxl | `textRotation=255` のみが縦書き | §4.3 参照 |
| CustomTkinter | バックグラウンドスレッドから UI 操作 → クラッシュ | `self.after(0, callback)` で委譲 |
| generator | fill 後にシートコピーするとプレースホルダー消失 | 全シート複製後に fill（GridGen/IndividualGen 共通） |
| C4th | 「保護者1　続柄」は全角スペース（U+3000） | `normalize_header()` で統一済み |
| PyInstaller | `--onefile` では CustomTkinter の .json/.otf が見つからない | `--onedir` のみ使用 |
| バッチファイル | 日本語パスが文字化け | `chcp 65001` 必須 |
| バッチファイル | `ren` はスペース含むパスで失敗 | `move` コマンドを使う |

---

## 参照先早見表

| やりたいこと | 参照先 |
|-------------|--------|
| C4th全50カラムの正確な列名を確認 | SPEC.md §2.1 |
| テンプレートのレイアウト詳細（列幅・行高） | SPEC.md §4.x |
| fill_placeholders の特殊キー一覧 | SPEC.md §5.2 |
| 印刷設定（余白・用紙サイズ） | SPEC.md §5.4 |
| GitHub Releases 自動更新の実装 | SPEC.md §6 / core/updater.py |
| 名簿データ暗号化同期 | SPEC.md §6 / core/data_sync.py / core/crypto.py |
| exe 自己更新バッチの実装コード | core/updater.py generate_update_batch() |
| build.spec のテンプレート | build.spec |
| テストケース一覧 | SPEC.md §7.8 |
