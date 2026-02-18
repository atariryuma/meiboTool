# 名簿帳票ツール 開発ロードマップ

> **最終更新**: 2026-02-19
> このファイルは作業完了のたびに更新する。CLAUDE.md の「現在地」セクションも合わせて更新すること。

---

## フェーズ概要

| フェーズ | 目標 | 状態 |
|---------|------|------|
| Phase 0 | 開発基盤（環境・テスト・リント） | ✅ **完了** |
| Phase 1 | MVP: 名列表を1枚出力できる | 🔄 進行中 |
| Phase 2 | 名札3種 + 台帳2種 + 調べ表 | ⏳ 未着手 |
| Phase 3 | 個票2種 + 自動更新 + ビルド | ⏳ 未着手 |

---

## Phase 0 — 開発基盤 ✅

**目標**: バグを書く前に検出できる環境を整える

- [x] `venv` + `requirements.txt`（customtkinter / openpyxl / pandas / ruff / pytest）
- [x] `pyproject.toml`（`pythonpath`・`testpaths`・ruff ルール一元管理）
- [x] `tests/conftest.py`（`dummy_df` / `default_options` / `default_config` フィクスチャ）
- [x] `sys.path` ハック全廃
- [x] `.gitignore`（`.pytest_cache/` `.ruff_cache/` `.claude/` 追加）
- [x] Git 初期化・3コミット済み
- [x] `docs/` 整備（SPEC.md / ROADMAP.md / images/ / .docx）
- [x] ruff クリーン（32件 → 0件）・テスト 21件全パス

---

## Phase 1 — MVP ✅/🔄

**目標**: ダミーデータを読み込んで「掲示用名列表.xlsx」を出力できる

### 完了済み ✅

- [x] `utils/wareki.py` — 西暦→和暦変換（令和/平成/昭和）
- [x] `utils/address.py` — 住所4フィールド結合・NaN除外
- [x] `utils/font_helper.py` — IPAmj明朝フォント適用
- [x] `core/config.py` — config.json 読み書き・パス解決
- [x] `core/mapper.py` — C4th 全50カラム → 内部論理名マッピング
- [x] `core/importer.py` — ヘッダー行自動検出 + dtype=str 読込
- [x] `core/generator.py` — BaseGenerator / GridGenerator / ListGenerator / IndividualGenerator 骨格
- [x] `templates/template_registry.py` — テンプレート9種メタデータ
- [x] `tests/generate_dummy.py` — ダミーC4th Excel生成スクリプト

### 残り作業 ❌

- [ ] **`templates/generators/gen_meireihyo.py`** ← **今すぐ着手**
  - `テンプレート/掲示用名列表.xlsx` を openpyxl でプログラム生成
  - 仕様: SPEC.md §4.4
  - A4縦・2列（左:No.1〜20、右:No.21〜40）・タイトル・薄ピンク背景
  - プレースホルダー: `{{出席番号}}` `{{氏名}}` `{{氏名かな}}` `{{学年}}` `{{組}}` `{{担任名}}`
- [ ] `tests/test_generator.py` — ListGenerator で名列表を出力し行数・シート名・値を検証
- [ ] `tests/fixtures/dummy_c4th.xlsx` 生成（`python meibo_tool/tests/generate_dummy.py`）
- [ ] **`gui/app.py`** — CustomTkinter メインウィンドウ（2カラムレイアウト）
  - 仕様: SPEC.md §3.1〜§3.4
- [ ] **`gui/frames/import_frame.py`** — ファイル選択・読込件数表示
- [ ] **`gui/frames/mandatory_panel.py`** — 組ドロップダウン・自動連番・確定ボタン
  - 仕様: SPEC.md §3.5（MandatoryInputPanel）
  - `App.mandatory_ok` フラグ管理
- [ ] **`gui/frames/select_frame.py`** — テンプレート選択ラジオボタン
- [ ] **`gui/frames/output_frame.py`** — 生成ボタン・進捗バー
- [ ] 統合確認: ダミーデータ読込 → 組設定 → 名列表.xlsx 出力まで一気通貫

---

## Phase 2 — 名札3種 + 台帳2種 + 調べ表

**目標**: 実際の授業開始に必要な帳票をすべて出力できる

### テンプレート生成スクリプト

- [ ] `gen_nafuda_plain.py` — 名札_通常（A4横・2列×5行・10枚/ページ）
  - 仕様: SPEC.md §4.1
- [ ] `gen_nafuda_decorated.py` — 名札_装飾あり（花柄ボーダー付き）
  - 仕様: SPEC.md §4.2
  - `テンプレート/assets/flower_border.png` が必要
- [ ] `gen_nafuda_1nen.py` — 名札_1年生用（縦長短冊・ふりがな縦書き）
  - 仕様: SPEC.md §4.3（`textRotation=255`）
- [ ] `gen_shuuryo_daicho.py` — 修了台帳（A4横・use_formal_name=true）
  - 仕様: SPEC.md §4.6
- [ ] `gen_sotsugyou_daicho.py` — 卒業台帳（A4横・証書番号付き）
  - 仕様: SPEC.md §4.7
- [ ] `gen_shirabehyo.py` — 調べ表（6列×10行グリッド）
  - 仕様: SPEC.md §4.5

### ジェネレーター拡張

- [ ] `GridGenerator._populate()` の完全実装
  - 番号付きプレースホルダー（`{{氏名_1}}` 〜 `{{氏名_10}}`）の展開
  - ページ自動追加ロジック
- [ ] `generate_all.py` — 全テンプレートを一括生成するエントリー

### テスト追加

- [ ] `test_generator.py` に GridGenerator テスト追加（名札10枚→2ページ確認）
- [ ] `test_generator.py` に ListGenerator 台帳テスト追加（use_formal_name=true 確認）

---

## Phase 3 — 個票2種 + 自動更新 + ビルド

**目標**: 全9テンプレート対応 + 配布可能な exe

### 個票テンプレート

- [ ] `gen_katei_chousahyo.py` — 家庭調査票（A4縦・1名/シート・最複雑）
  - 仕様: SPEC.md §4.8
- [ ] `gen_gakkyuu_kojihyo.py` — 学級編成用個票（A4縦・1名/シート）
  - 仕様: SPEC.md §4.9
- [ ] `IndividualGenerator._populate()` の完全実装（画像付きコピー確認）

### 自動更新

- [ ] `core/updater.py` — Google Drive 更新チェック
  - 仕様: SPEC.md §6.4〜§6.7
  - `check_for_updates()` / `download_from_gdrive()` / `generate_update_batch()`
- [ ] `gui/dialogs/update_dialog.py` — 更新確認ダイアログ
- [ ] `core/config.py` に `update.version_file_id` を設定

### GUI 完成

- [ ] `gui/dialogs/mapping_dialog.py` — カラムマッピング手動調整画面
- [ ] `gui/dialogs/add_column_dialog.py` — 任意カラム追加ダイアログ（担任名・証書番号等）
- [ ] ステータスバー（更新通知・生成完了メッセージ）

### ビルド・配布

- [ ] `build.spec` 作成（`--onedir`・CustomTkinter データファイル同梱）
  - 仕様: SPEC.md §7.1
- [ ] `pyinstaller build.spec` で `dist/名簿帳票ツール/` 生成確認
- [ ] 別 PC（Python なし環境）での起動確認
- [ ] Google Drive に配布フォルダ作成・`version.json` 配置

---

## 既知の落とし穴（実装時に必ず確認）

| 場所 | 落とし穴 | 対策 |
|------|---------|------|
| openpyxl | `cell.font.bold = True` → AttributeError | `cell.font = Font(bold=True)` で新インスタンス |
| openpyxl | `copy_worksheet` は画像をコピーしない | `source_ws._images` をループして `add_image()` |
| openpyxl | 結合セルの左上以外に書き込み → AttributeError | `MergedCell` チェックを入れる |
| openpyxl | `textRotation=255` のみが縦書き。0〜180 は回転角度 | §4.3 参照 |
| CustomTkinter | バックグラウンドスレッドから UI 操作 → クラッシュ | `self.after(0, callback)` で委譲 |
| C4th | 「保護者1　続柄」は全角スペース（U+3000） | `normalize_header()` で統一済み |
| PyInstaller | `--onefile` では CustomTkinter の .json/.otf が見つからない | `--onedir` のみ使用 |

---

## 参照先早見表

| やりたいこと | 参照先 |
|-------------|--------|
| C4th全50カラムの正確な列名を確認 | SPEC.md §2.1 |
| テンプレートのレイアウト詳細（列幅・行高） | SPEC.md §4.x |
| fill_placeholders の特殊キー一覧 | SPEC.md §5.2 |
| 印刷設定（余白・用紙サイズ） | SPEC.md §5.4 |
| Google Drive ダウンロードの実装コード | SPEC.md §6.5 |
| exe 自己更新バッチの実装コード | SPEC.md §6.6 |
| build.spec のテンプレート | SPEC.md §7.1 |
| テストケース一覧 | SPEC.md §7.8 |
