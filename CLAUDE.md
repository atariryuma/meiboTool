# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

> 詳細仕様は `docs/SPEC.md` を参照。各章に仕様・実装コード・落とし穴がセットで記載されている。

## プロジェクト概要

那覇市立小中学校向けの **名簿帳票自動生成 Windows デスクトップアプリ**。
C4th（EDUCOM 校務支援システム）の Excel エクスポートを読み込み、名札・ラベル・名簿など 9 種の Excel テンプレート帳票 + 27 種の .lay レイアウト帳票を自動生成する。Python 3.11 + CustomTkinter + openpyxl。

## コマンド

```bash
# 仮想環境セットアップ（初回のみ）
python -m venv venv
venv/Scripts/activate      # Windows Git Bash / Mac: source venv/bin/activate
pip install -r requirements.txt

# アプリ起動（プロジェクトルートから実行）
cd meibo_tool && python main.py

# テスト全実行（プロジェクトルートから。pyproject.toml で testpaths 設定済み）
venv/Scripts/python.exe -m pytest

# 単一テストファイル実行
venv/Scripts/python.exe -m pytest meibo_tool/tests/test_wareki.py -v

# リント（ゼロエラーになってからコミット）
venv/Scripts/python.exe -m ruff check meibo_tool/
venv/Scripts/python.exe -m ruff check meibo_tool/ --fix   # 自動修正

# テンプレート Excel 生成（初回および更新時）
cd meibo_tool && python -m templates.generators.generate_all

# exe ビルド
pyinstaller build.spec
# → dist/名簿帳票ツール/ フォルダが出力される
```

## アーキテクチャ

### データフロー

```text
C4th Excel
  → core/importer.py  （ヘッダー行自動検出 → pandas DataFrame に読込）
  → core/mapper.py    （C4th カラム名 → 内部論理名に変換）
  → gui/frames/       （クラス選択パネルで学年・組を選択）
  → core/generator.py （テンプレートにデータを差込 → xlsx 出力）
```

### generator.py の 3 種類のジェネレーター

| クラス | 用途 | テンプレート |
| -------- | ------ | ------------ |
| `GridGenerator` | 1 ページに N 名をグリッド配置 | 名札・名列表・調べ表・横名簿 等 |
| `ListGenerator` | データ行を児童数分コピー展開 | 台帳 2 種 |
| `IndividualGenerator` | 1 名/シートで `copy_worksheet` 複製 | 個票 2 種（未テンプレート） |

`create_generator()` がファクトリー関数。`TEMPLATES` レジストリの `type` キーで分岐する。

### プレースホルダー形式

テンプレート Excel 内は `{{論理名}}` 形式。`fill_placeholders()` が正規表現で一括置換する。
特殊キー: `{{年度}}` `{{年度和暦}}` `{{学校名}}` `{{担任名}}` `{{住所}}` (住所 4 フィールド自動結合)。

### GUI 構成

```text
App (CTk) — 2カラムレイアウト
├── ヘッダーバー（設定ボタン・レイアウトエディターボタン）
├── 左パネル（CTkScrollableFrame）
│   ├── ImportFrame      — ファイル選択・読込件数・同期ステータス
│   ├── ClassSelectPanel — 学年・組の選択（特別支援学級対応）
│   ├── SelectFrame      — テンプレート選択・担任名・学校名（on_template_change コールバック）
│   ├── OutputFrame      — 生成ボタン・進捗バー
│   └── RosterPrintPanel — レイアウト印刷パネル
└── 右パネル（_PreviewPanel — CTkTabview）
    ├── 「データ」タブ   — Treeview でデータ一覧表示（セル編集可能）
    └── 「プレビュー」タブ — PIL レンダリング画像で帳票プレビュー表示
```

### 自動更新

- **アプリ更新**: GitHub Releases API で最新バージョン確認 → zip ダウンロード → バッチファイルで更新
- **名簿データ同期**: LAN（共有フォルダ）/ Google Drive（AES-256-GCM 暗号化）の 2 モード対応
- **暗号化**: `core/crypto.py` — AES-256-GCM + PBKDF2-HMAC-SHA256 (600,000 回)、パスワードは Windows DPAPI で保護

## コーディング規約

- **リンター**: ruff。コード変更後は必ず `ruff check meibo_tool/` でゼロエラーを確認してからコミット
- **import 順序**: 標準ライブラリ → サードパーティ → 社内モジュール（ruff の I001 で自動チェック）
- **未使用ループ変数**: `_name` のようにアンダースコアプレフィックスを付ける
- **型注釈**: 関数シグネチャには型注釈を付ける（`def func(x: str) -> int:`）
- **テスト**: 新機能は同時に `tests/test_<モジュール名>.py` に追加
- **sys.path ハック禁止**: テストファイルに `sys.path.insert` を書かない（`pyproject.toml` の `pythonpath` で解決済み）
- **共通フィクスチャ**: `meibo_tool/tests/conftest.py` に追加（`dummy_df`・`default_options`・`default_config`）

## 絶対守るべきルール

1. **PyInstaller は `--onedir` のみ** — CustomTkinter の .json/.otf が `--onefile` では見つからずクラッシュする
2. **openpyxl スタイルはイミュータブル** — `cell.font.bold = True` は AttributeError。必ず `cell.font = Font(...)` で新インスタンスを作る
3. **`copy_worksheet` は画像をコピーしない** — `_data()` でバイナリ取得 → `BytesIO` → `Image()` で再構築（`copy_sheet_with_images()` 参照）
4. **結合セルは左上セルにのみ書き込む** — 他セルへの書き込みは `AttributeError: 'MergedCell' is read-only`
5. **CustomTkinter の UI 操作はメインスレッドのみ** — バックグラウンドスレッドからは `self.after(0, callback)` で委譲する。ジェネレーター作成（Tkinter ウィジェット参照）もメインスレッドで行う
6. **C4th ヘッダーの全角スペース（U+3000）** — 「保護者1　続柄」等は全角スペース区切り。`normalize_header()` で統一する
7. **IPAmj明朝フォント** — 正式名前列に IVS 付き異体字を含むため必須。`font_helper.apply_font()` で全セルに適用
8. **IndividualGenerator のシート複製は fill 前に行う** — fill 後にコピーするとプレースホルダーが消失する

## Excel テンプレート 9 種（カテゴリ別）

| カテゴリ | テンプレートファイル名 | タイプ | 状態 |
| -------- | ---------------------- | ------ | ---- |
| 名札・ラベル | ラベル_色付き.xlsx / サンプル_名札.xlsx / 名札_1年生用.xlsx | grid | ✅ |
| 名札・ラベル | ラベル_大2.xlsx / ラベル_小.xlsx / ラベル_特大.xlsx | grid | ✅ |
| 名簿・出欠表 | 横名簿.xlsx / 縦一週間.xlsx | grid | ✅ |
| 名簿・出欠表 | 男女一覧.xlsx | grid | ✅ sort_by=性別 |

## .lay レイアウト 27 種（デフォルト同梱）

初回起動時に `resources/default_layouts.lay` から自動インポートされる。
掲示用名列表・調べ表・修了台帳・卒業台帳・家庭調査票・学級編成用個票 等を含む。
英語名のレイアウト（takara_simei 等）はパーツ（部品）として他レイアウトから参照される。

## C4th データの重要な仕様

- **「組」「出席番号」は C4th に含まれない** — ファイル読込後にクラス選択パネルで教員が選択する（将来 C4th から出力予定）
- **特別支援学級** — C4th の「組」列に非数字の値（「なかよし」「ひまわり」等）が入る。`is_special_needs_class()` で判定。交流学級は初回インポート時にダイアログで割り当て → config に保存 → 以降自動統合
- **全カラム `dtype=str` で読み込む** — 生年月日等の型変換は後工程で行う
- **ヘッダー行自動検出** — ファイル先頭にメタ情報行がある可能性があるため、文字列セル 5 個以上の行を検索する

## config.json の主要キー

| キー | 説明 |
| ---- | ---- |
| `school_name` | 学校名（帳票タイトルに使用） |
| `school_type` | `"elementary"` / `"middle"` |
| `fiscal_year` | 年度（西暦・起動時に自動計算） |
| `template_dir` | テンプレートフォルダパス（`./テンプレート`） |
| `output_dir` | 出力フォルダパス（`./出力`） |
| `layout_dir` | レイアウトライブラリフォルダパス（`./レイアウト`） |
| `update.github_repo` | GitHub リポジトリ（`owner/repo` 形式） |
| `update.check_on_startup` | 起動時に更新チェックするか |
| `update.skip_version` | スキップ設定したバージョン |
| `data_source.mode` | `"manual"` / `"lan"` / `"gdrive"` |
| `data_source.lan_path` | LAN 共有フォルダの UNC パス |
| `data_source.gdrive_file_id` | Google Drive ファイル ID |
| `data_source.encryption_password` | DPAPI で保護された暗号化パスワード |
| `homeroom_teachers` | `{"1-1": "山田先生"}` 形式で担任名を保存 |
| `special_needs_assignments` | `{"1-なかよし-1": "1-1"}` 形式で特支児童の交流学級割り当てを保存 |
| `special_needs_placement` | `"appended"` (末尾追加) / `"integrated"` (出席番号順統合) |

## 現在地（セッション開始時に必ず確認）

> 詳細なフェーズ計画は `docs/ROADMAP.md` を参照。

### モジュール一覧

#### core/ — コアロジック（16 ファイル）

| ファイル | 内容 | テスト |
| -------- | ---- | ------ |
| `config.py` | config.json 読み書き・deep_merge・パス解決 | ✅ 13 |
| `crypto.py` | AES-256-GCM 暗号化/復号 + DPAPI パスワード保護 | ✅ 17 |
| `data_model.py` | EditableDataModel（GUI 上でデータ編集 + Undo/Redo） | ✅ 23 |
| `data_sync.py` | 名簿データ自動同期（LAN/GDrive/手動） | ✅ 15 |
| `exporter.py` | Excel エクスポート | ✅ 6 |
| `generator.py` | Grid/List/Individual ジェネレーター + 性別ソート | ✅ 29 |
| `importer.py` | ヘッダー自動検出付き Excel/CSV 読込 | ✅ 26 |
| `lay_parser.py` | .lay バイナリパーサー（TLV 再帰 + FontInfo + raw_tags） | ✅ 74 |
| `lay_renderer.py` | Canvas/PILBackend 描画 + fill_layout + MEIBO 展開 | ✅ 101 |
| `lay_serializer.py` | LayFile ↔ JSON 保存/読込（ラウンドトリップ保証） | ✅ 44 |
| `layout_registry.py` | レイアウトライブラリ管理（scan/import/delete/rename） | ✅ 43 |
| `mapper.py` | C4th カラム名マッピング + resolve_name_fields | ✅ 12 |
| `photo_manager.py` | 児童写真管理 | ✅ 46 |
| `special_needs.py` | 特別支援学級判定・検出・統合・交流学級割り当て | ✅ 36 |
| `updater.py` | GitHub Releases ベースのアプリ更新 | ✅ 33 |
| `win_printer.py` | Windows GDI 直接印刷（縦書き・bold/italic 対応） | ✅ 15 |

#### gui/ — ユーザーインターフェース（21 ファイル）

| ファイル | 内容 | テスト |
| -------- | ---- | ------ |
| `app.py` | メインウィンドウ（2カラム + CTkTabview + 特支自動統合） | ✅ 35 |
| `preview_renderer.py` | openpyxl Worksheet → PIL Image レンダラー | ✅ 20 |
| `frames/import_frame.py` | ファイル選択・同期ステータス表示 | — |
| `frames/class_select_panel.py` | 学年・組選択 + 特別支援学級 | — |
| `frames/select_frame.py` | テンプレート選択・担任名・学校名・特支配置 | — |
| `frames/output_frame.py` | 生成ボタン・進捗バー | — |
| `frames/roster_print_panel.py` | レイアウト印刷パネル | ✅ 22 |
| `dialogs/mapping_dialog.py` | カラムマッピング手動調整 | ✅ 13 |
| `dialogs/settings_dialog.py` | 管理者設定（同期モード設定） | — |
| `dialogs/exchange_class_dialog.py` | 交流学級割り当てダイアログ | — |
| `dialogs/update_dialog.py` | 更新確認ダイアログ | — |
| `editor/editor_window.py` | レイアウトエディター メインウィンドウ | — |
| `editor/layout_canvas.py` | インタラクティブ Canvas（選択・移動・リサイズ） | — |
| `editor/properties_panel.py` | オブジェクトプロパティ編集 | — |
| `editor/toolbar.py` | ファイル操作・ズーム・オブジェクト追加/削除 | — |
| `editor/print_dialog.py` | プリンター選択・印刷実行 | — |
| `editor/print_preview_dialog.py` | 印刷プレビュー（PILBackend + ページ送り） | — |
| `editor/layout_manager_dialog.py` | レイアウトライブラリ管理（Treeview 一覧） | — |
| `editor/object_list.py` | オブジェクト一覧（TreeView + Canvas 連動） | — |

#### utils/ — ユーティリティ（4 ファイル）

| ファイル | 内容 | テスト |
| -------- | ---- | ------ |
| `wareki.py` | 西暦→和暦変換 | ✅ 10 |
| `address.py` | 住所4フィールド結合 | ✅ 12 |
| `date_fmt.py` | 日付フォーマット（YY/MM/DD・Excel シリアル値対応） | ✅ 22 |
| `font_helper.py` | IPAmj明朝フォント適用 | ✅ 11 |

#### templates/ — テンプレート管理（7 ファイル）

| ファイル | 内容 | テスト |
| -------- | ---- | ------ |
| `template_registry.py` | テンプレート 9 種メタデータ + カテゴリ別表示 | ✅ 28 |
| `template_scanner.py` | テンプレート自動検出（.xlsx スキャン → メタデータ推定） | ✅ 29 |
| `generators/gen_nafuda.py` | 名札テンプレート生成 | ✅ 16 |
| `generators/gen_from_lay.py` | .lay → Excel テンプレート変換 | ✅ 27 |
| `generators/gen_from_legacy.py` | レガシーテンプレート変換 | — |
| `generators/generate_all.py` | 全テンプレート一括生成 | — |

#### その他

| ファイル | 内容 |
| -------- | ---- |
| `.github/workflows/build-release.yml` | CI/CD（テスト→ビルド→Release） |
| `tests/conftest.py` | 共通フィクスチャ（dummy_df・default_options・default_config） |

### 未実装 ❌

1. 別 PC（Python なし環境）での exe 起動確認

### 開発環境の状態

- テスト: **773 件全パス**（`venv/Scripts/python.exe -m pytest`）
- テストファイル: 32 ファイル
- リント: ruff クリーン（`venv/Scripts/python.exe -m ruff check meibo_tool/`）
- Git: master ブランチ

## 詳細仕様の参照先

| 内容 | 参照先 |
| ---- | ------ |
| C4th 全 50 カラム定義・EXACT_MAP | SPEC.md 第2章 |
| GUI 各パネルの詳細仕様 | SPEC.md 第3章 |
| Excel テンプレート 9 種・ジェネレーター | SPEC.md 第4章 |
| 特別支援学級 | SPEC.md 第5章 |
| 自動更新（GitHub Releases + 名簿データ同期） | SPEC.md 第6章 |
| ビルド・テスト構成 | SPEC.md 第7章 |
| .lay レイアウトシステム（パーサー・レンダラー・エディター） | SPEC.md 第8章 |
| レイアウト帳票のフィールド ID 一覧 | core/lay_parser.py `FIELD_ID_MAP` |
| テンプレート作成ガイド（Excel） | docs/TEMPLATE_GUIDE.md |
| 開発ロードマップ（フェーズ別進捗） | docs/ROADMAP.md |
