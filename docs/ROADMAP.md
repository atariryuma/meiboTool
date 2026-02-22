# 名簿帳票ツール 開発ロードマップ

> **最終更新**: 2026-02-22（不要ファイル・テスト整理 — 773 テスト全パス）
> このファイルは作業完了のたびに更新する。CLAUDE.md の「現在地」セクションも合わせて更新すること。

---

## フェーズ概要

| フェーズ | 目標 | 状態 |
|---------|------|------|
| Phase 0 | 開発基盤（環境・テスト・リント） | ✅ **完了** |
| Phase 1 | MVP: 名列表を1枚出力できる | ✅ **完了**（GUI 含む） |
| Phase 2 | 名札3種 + 台帳2種 + 調べ表 + 追加帳票 | ✅ **完了** |
| Phase 3 | 自動更新 + 暗号化同期 + 品質改善 | ✅ **完了** |
| Phase 4 | .lay レイアウトシステム + CSV + データ編集 | ✅ **完了** |
| Phase 5 | exe ビルド + 配布 + 最終調整 | ⏳ 残タスク |

---

## Phase 0 — 開発基盤 ✅

**目標**: バグを書く前に検出できる環境を整える

- [x] `venv` + `requirements.txt`（customtkinter / openpyxl / pandas / ruff / pytest / cryptography / chardet / Pillow）
- [x] `pyproject.toml`（`pythonpath`・`testpaths`・ruff ルール一元管理）
- [x] `tests/conftest.py`（`dummy_df` / `default_options` / `default_config` フィクスチャ）
- [x] `sys.path` ハック全廃
- [x] `.gitignore`（`.pytest_cache/` `.ruff_cache/` `.claude/` 追加）
- [x] Git 初期化
- [x] `docs/` 整備（SPEC.md / ROADMAP.md / TEMPLATE_GUIDE.md）
- [x] ruff クリーン・テスト全パス

---

## Phase 1 — MVP ✅

**目標**: ダミーデータを読み込んで「掲示用名列表.xlsx」を出力できる

- [x] `utils/wareki.py` — 西暦→和暦変換（令和/平成/昭和）
- [x] `utils/address.py` — 住所4フィールド結合・NaN除外
- [x] `utils/font_helper.py` — IPAmj明朝フォント適用
- [x] `utils/date_fmt.py` — 日付フォーマット（YY/MM/DD・Excel シリアル値対応）
- [x] `core/config.py` — config.json 読み書き・deep_merge・パス解決
- [x] `core/mapper.py` — C4th 全50カラム → 内部論理名マッピング + resolve_name_fields
- [x] `core/importer.py` — ヘッダー自動検出 + dtype=str 読込 + CSV 対応
- [x] `templates/template_registry.py` — テンプレート9種メタデータ + カテゴリ別表示
- [x] `templates/generators/generate_all.py` — 全テンプレート一括生成
- [x] `gui/app.py` — CustomTkinter メインウィンドウ（2カラム + CTkTabview プレビュー）
- [x] `gui/frames/import_frame.py` — ファイル選択・読込件数・同期ステータス表示
- [x] `gui/frames/class_select_panel.py` — 学年・組選択 + 特別支援学級表示
- [x] `gui/frames/select_frame.py` — テンプレート選択・担任名・学校名・特支配置設定
- [x] `gui/frames/output_frame.py` — 生成ボタン・進捗バー

---

## Phase 2 — 名札3種 + 台帳2種 + 調べ表 + 追加帳票 ✅

**目標**: 実際の授業開始に必要な帳票をすべて出力できる

### テンプレート生成スクリプト ✅

- [x] `gen_nafuda.py` — 名札3種（通常/装飾/1年生）+ ラベル3種（大2/小/特大）
- [x] `gen_from_legacy.py` — レガシーテンプレート変換（横名簿/縦一週間/男女一覧）

### ジェネレーター拡張 ✅

- [x] `GridGenerator._populate()` — ページネーション完全実装
- [x] `ListGenerator._find_template_row()` — 全名前系プレースホルダー検出
- [x] `IndividualGenerator._populate()` — fill 前にシート複製する方式に修正
- [x] `template_registry.py` — 9種メタデータ・`get_display_groups()` 追加

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

---

## Phase 4 — .lay レイアウトシステム + CSV + データ編集 ✅

**目標**: スズキ校務 .lay レイアウト27種の完全対応 + データ柔軟性の向上

### .lay バイナリパーサー ✅
- [x] `core/lay_parser.py` — EXCMIDataContainer01 形式パーサー
  - zlib 解凍 → TLV 再帰パース → LayoutObject / FontInfo / PaperLayout 構造体
  - `parse_lay_multi()` でマルチレイアウト .lay を一括パース
  - FIELD_ID_MAP（140+ フィールド ID → 内部論理名）
  - style タグ (0x03E9/0x03EA/0x03EB) の保持
  - raw_tags（path 付き生 TLV）でタグ情報を完全保持
  - ObjectType: LABEL / FIELD / LINE / IMAGE / TABLE / GROUP / MEIBO

### JSON シリアライザ ✅
- [x] `core/lay_serializer.py` — LayFile ↔ JSON 保存/読込
  - bold/italic/style/raw_tags のラウンドトリップ保証
  - base64 ペイロード保持

### レンダラー ✅
- [x] `core/lay_renderer.py` — PILBackend 描画エンジン
  - 縦書き/複数行テキスト/幅ベース自動折り返し
  - `fill_layout()` — MEIBO 展開 + FIELD データ差込
  - `render_layout_to_image()` — PIL Image 生成
  - GROUP 直接描画
  - style_1002==10 枠線描画

### レイアウトライブラリ管理 ✅
- [x] `core/layout_registry.py` — scan/import/delete/rename + スズキ校務エイリアス解決

### Windows GDI 印刷 ✅
- [x] `core/win_printer.py` — レイアウト直接印刷（縦書き・複数行・bold/italic 対応）

### レイアウトエディター ✅
- [x] `gui/editor/editor_window.py` — メインウィンドウ（Canvas + プロパティ + Undo/Redo）
- [x] `gui/editor/layout_canvas.py` — インタラクティブ Canvas（選択・移動・リサイズ）
- [x] `gui/editor/properties_panel.py` — オブジェクトプロパティ編集パネル
- [x] `gui/editor/toolbar.py` — ファイル操作・ズーム・オブジェクト追加/削除
- [x] `gui/editor/print_dialog.py` — プリンター選択・印刷実行
- [x] `gui/editor/print_preview_dialog.py` — 印刷プレビュー（PILBackend + ページ送り）
- [x] `gui/editor/layout_manager_dialog.py` — レイアウトライブラリ管理（Treeview 一覧）
- [x] `gui/editor/object_list.py` — オブジェクト一覧（TreeView + Canvas 連動）
- [x] `gui/frames/roster_print_panel.py` — レイアウト印刷パネル

### .lay → Excel 変換 ✅
- [x] `templates/generators/gen_from_lay.py` — .lay レイアウトから Excel テンプレートを自動生成

### CSV インポート + データ編集 ✅
- [x] `core/importer.py` — CSV インポート対応（文字コード自動判定 chardet）
- [x] `core/data_model.py` — EditableDataModel（GUI 上でデータ編集 + Undo/Redo）
- [x] `core/exporter.py` — Excel エクスポート

### 写真管理 ✅
- [x] `core/photo_manager.py` — 児童写真管理（学級編成用個票向け）

### 特別支援学級 ✅
- [x] `core/special_needs.py` — 特支学級判定・検出・統合・交流学級割り当て
- [x] `gui/dialogs/exchange_class_dialog.py` — 交流学級割り当てダイアログ

### その他 GUI ✅
- [x] `gui/preview_renderer.py` — openpyxl Worksheet → PIL Image レンダラー
- [x] `gui/dialogs/mapping_dialog.py` — カラムマッピング手動調整ダイアログ
- [x] `templates/template_scanner.py` — テンプレート自動検出（.xlsx スキャン → メタデータ推定）

---

## Phase 5 — exe ビルド + 配布 + 最終調整

**目標**: 全テンプレート対応 + 配布可能な exe

### ビルド・配布

- [x] `build.spec` 作成（`--onedir`・CustomTkinter データファイル同梱）
- [ ] 別 PC（Python なし環境）での exe 起動確認

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
| .lay パーサー | マルチレイアウトと単一レイアウトで座標単位が異なる | 0.1mm/unit (マルチ) vs 0.25mm/unit (単一) |
| .lay レンダラー | MEIBO 参照名とレイアウト名の不一致 | `_SUZUKI_REF_ALIASES` でエイリアス解決 |
| IPAmj明朝 | データフィールドは外字のみ IPAmj明朝、他はレイアウト元フォント | `font_helper.apply_font()` の混合モード |

---

## テスト状況

| 指標 | 値 |
|------|------|
| テスト総数 | **773** |
| テストファイル数 | 32 |
| リント | ruff クリーン |

---

## 参照先早見表

| やりたいこと | 参照先 |
|-------------|--------|
| C4th 全50カラムの正確な列名を確認 | SPEC.md §2.1 |
| テンプレートのレイアウト詳細（列幅・行高） | SPEC.md §4.x |
| fill_placeholders の特殊キー一覧 | SPEC.md §5.2 |
| 印刷設定（余白・用紙サイズ） | SPEC.md §5.4 |
| .lay バイナリ形式の詳細 | SPEC.md §8 |
| レイアウトのフィールド ID 一覧 | core/lay_parser.py `FIELD_ID_MAP` |
| GitHub Releases 自動更新の実装 | SPEC.md §6 / core/updater.py |
| 名簿データ暗号化同期 | SPEC.md §6 / core/data_sync.py / core/crypto.py |
| テンプレート作成方法 | docs/TEMPLATE_GUIDE.md |
| build.spec | build.spec |
