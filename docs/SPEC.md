# 名簿帳票ツール 開発仕様書

> **バージョン**: 4.0（2026-02-22 現状同期版）
> **対象**: Claude Code + VSCode によるコーディングエージェント実装
> **構成方針**: 各章が「仕様 → 実装コード → 落とし穴」をワンセットで含む

---

## 第1章 プロジェクト概要

### 1.1 背景と目的

那覇市立小中学校では、校務支援システム（C4th / EDUCOM）からエクスポートした Excel 名簿データを基に、名札・名列表・台帳・個票などの帳票を手作業で作成している。本プロジェクトでは、名簿データの読み込みからテンプレートへの差し込み、Excel ファイル出力までを自動化する Windows デスクトップアプリケーションを開発する。

加えて、スズキ校務（旧校務支援システム）の .lay バイナリレイアウト形式を完全サポートし、27 種のレイアウト帳票を直接印刷・プレビュー・編集できる。

### 1.2 前提条件と制約

- **IPAmj明朝フォント要件**: 児童生徒の氏名に外字（異体字）が含まれるケースがあり、58,712 文字を収録する IPAmj明朝で正確に表示・印刷する。C4th の「正式名前」列に IVS 付き文字が含まれる。データフィールドでは外字のみ IPAmj明朝、それ以外はレイアウト元フォントを使用する。
- **動作環境**: Windows 10/11、Microsoft Excel 2016 以降、Python ランタイム不要（exe 化配布）。
- **ネットワーク**: 初回起動・自動更新時のみ GitHub Releases / Google Drive への HTTPS 接続。帳票生成自体はオフライン動作。

### 1.3 技術スタック

| 項目 | 技術 | バージョン | 用途 |
| ---- | ---- | --------- | ---- |
| 言語 | Python | 3.11+ | メインロジック |
| GUI | CustomTkinter | 5.x | モダンな UI 構築 |
| Excel 操作 | openpyxl | 3.1+ | テンプレート読み込み・データ書き込み・xlsx 出力 |
| データ処理 | pandas | 2.x | 名簿データ操作 |
| 画像 | Pillow | 9.0+ | レイアウト描画・プレビュー・openpyxl の画像操作 |
| 暗号化 | cryptography | 41+ | AES-256-GCM 暗号化/復号 |
| 文字コード | chardet | 5.x | CSV エンコーディング自動判定 |
| exe 化 | PyInstaller | 6.x | パッケージング |
| HTTP 通信 | requests | 2.x | GitHub Releases / Google Drive 通信 |
| テスト | pytest | 8.x | 自動テスト（773 ケース） |
| リント | ruff | 0.x | コードスタイル検査 |

### 1.4 ソースコード構成

```
meibo_tool/
  ├── main.py                           # エントリーポイント
  ├── core/
  │   ├── config.py                     # config.json 読み書き・deep_merge・パス解決
  │   ├── crypto.py                     # AES-256-GCM 暗号化/復号 + DPAPI
  │   ├── data_model.py                 # EditableDataModel（GUI 上でデータ編集 + Undo/Redo）
  │   ├── data_sync.py                  # 名簿データ同期（manual/lan/gdrive 3モード）
  │   ├── exporter.py                   # Excel エクスポート
  │   ├── generator.py                  # テンプレートへのデータ差込（Grid/List/Individual）
  │   ├── importer.py                   # C4th Excel/CSV データ読込（ヘッダー自動検出）
  │   ├── lay_parser.py                 # .lay バイナリパーサー（TLV 再帰パース）
  │   ├── lay_renderer.py              # Canvas/PILBackend 描画 + fill_layout
  │   ├── lay_serializer.py            # LayFile ↔ JSON 保存/読込
  │   ├── layout_registry.py           # レイアウトライブラリ管理（scan/import/delete/rename）
  │   ├── mapper.py                     # C4th カラム名マッピング + resolve_name_fields
  │   ├── photo_manager.py             # 児童写真管理
  │   ├── special_needs.py             # 特別支援学級の判定・統合・交流学級割り当て
  │   ├── updater.py                   # GitHub Releases ベースのアプリ更新
  │   └── win_printer.py              # Windows GDI 直接印刷（縦書き・bold/italic 対応）
  ├── gui/
  │   ├── app.py                       # メインウィンドウ（2カラム + CTkTabview）
  │   ├── preview_renderer.py          # openpyxl Worksheet → PIL Image レンダラー
  │   ├── dialogs/
  │   │   ├── exchange_class_dialog.py # 交流学級割り当てダイアログ
  │   │   ├── mapping_dialog.py        # カラムマッピング手動調整
  │   │   ├── settings_dialog.py       # 管理者設定（同期モード等）
  │   │   └── update_dialog.py         # 更新確認ダイアログ
  │   ├── editor/
  │   │   ├── editor_window.py         # レイアウトエディター メインウィンドウ
  │   │   ├── layout_canvas.py         # インタラクティブ Canvas（選択・移動・リサイズ）
  │   │   ├── layout_manager_dialog.py # レイアウトライブラリ管理（Treeview 一覧）
  │   │   ├── object_list.py           # オブジェクト一覧（TreeView + Canvas 連動）
  │   │   ├── print_dialog.py          # プリンター選択・印刷実行
  │   │   ├── print_preview_dialog.py  # 印刷プレビュー（PILBackend + ページ送り）
  │   │   ├── properties_panel.py      # オブジェクトプロパティ編集パネル
  │   │   └── toolbar.py              # ファイル操作・ズーム・オブジェクト追加/削除
  │   └── frames/
  │       ├── class_select_panel.py    # 学年・組選択 + 特別支援学級
  │       ├── import_frame.py          # ファイル選択・同期ステータス
  │       ├── output_frame.py          # 生成ボタン・進捗バー
  │       ├── roster_print_panel.py    # レイアウト印刷パネル
  │       └── select_frame.py         # テンプレート選択・担任名・学校名・特支配置
  ├── templates/
  │   ├── template_registry.py         # テンプレート9種メタデータ + カテゴリ別表示
  │   ├── template_scanner.py          # テンプレート自動検出（.xlsx スキャン）
  │   └── generators/
  │       ├── convert_legacy_xlsx.py   # レガシー XLSX 変換
  │       ├── gen_from_lay.py          # .lay → Excel テンプレート変換
  │       ├── gen_from_legacy.py       # レガシーテンプレート変換
  │       ├── gen_nafuda.py            # 名札テンプレート生成（6種）
  │       ├── gen_sample_nafuda.py     # サンプル名札生成
  │       └── generate_all.py         # 全テンプレート一括生成
  ├── utils/
  │   ├── address.py                   # 住所4フィールド結合・NaN除外
  │   ├── date_fmt.py                  # 日付フォーマット（YY/MM/DD・Excel シリアル値対応）
  │   ├── font_helper.py              # IPAmj明朝フォント適用
  │   └── wareki.py                   # 西暦→和暦変換
  └── tests/                           # 32 テストファイル（773 テスト）
      ├── conftest.py                  # 共通フィクスチャ（dummy_df, default_options, default_config）
      ├── test_address.py
      ├── test_app_logic.py
      ├── test_config.py
      ├── test_crypto.py
      ├── test_data_model.py
      ├── test_data_sync.py
      ├── test_date_fmt.py
      ├── test_exporter.py
      ├── test_font_helper.py
      ├── test_gen_from_lay.py
      ├── test_gen_nafuda.py
      ├── test_generator.py
      ├── test_importer.py
      ├── test_integration.py
      ├── test_lay_parser.py
      ├── test_lay_renderer.py
      ├── test_lay_serializer.py
      ├── test_layout_registry.py
      ├── test_list_individual_gen.py
      ├── test_mapper.py
      ├── test_mapper_resolve.py
      ├── test_mapping_dialog.py
      ├── test_photo_manager.py
      ├── test_pil_backend.py
      ├── test_preview_renderer.py
      ├── test_roster_print_panel.py
      ├── test_special_needs.py
      ├── test_template_registry.py
      ├── test_template_scanner.py
      ├── test_updater.py
      ├── test_wareki.py
      └── test_win_printer.py
```

### 1.5 配布フォルダ構成（--onedir 方式）

```
名簿帳票ツール/                    ← フォルダ全体を共有
  ├── 名簿帳票ツール.exe            # アプリ本体
  ├── _internal/                    # PyInstaller 内部ファイル（触らない）
  │   ├── customtkinter/
  │   └── ... (DLL群)
  ├── config.json                   # 設定ファイル
  ├── テンプレート/                  # Excel テンプレート群（9種）
  ├── レイアウト/                    # .lay JSON レイアウト群（27種）
  ├── 写真/                          # 児童写真フォルダ
  └── 出力/                          # 生成ファイル出力先
```

### 1.6 処理フロー全体像

```
① exe ダブルクリック → アプリ起動
② [バックグラウンド] updater.py が GitHub Releases の最新バージョンをチェック
③ 更新がある場合 → ダイアログで通知 → 同意でダウンロード＆更新
④ [バックグラウンド] data_sync.py が名簿データの同期を試行
⑤ 「名簿読込」ボタンで C4th エクスポート Excel/CSV を選択
⑥ importer.py がヘッダー自動検出 → mapper.py でカラムマッピング
⑦ ClassSelectPanel で学年・組を選択
⑧ テンプレート一覧から帳票を選択（Excel テンプレート 9 種 + .lay レイアウト 27 種）
⑨ 「Excel 生成」ボタン → generator.py が xlsx 出力
   または「印刷」ボタン → win_printer.py / lay_renderer.py が直接印刷
⑩ 出力フォルダが開き、教員が確認
```

### 1.7 config.json 仕様

```json
{
  "app_version": "1.0.0",
  "school_name": "",
  "school_type": "elementary",
  "template_dir": "./テンプレート",
  "output_dir": "./出力",
  "layout_dir": "./レイアウト",
  "photo_dir": "./写真",
  "default_font": "IPAmj明朝",
  "fiscal_year": 2025,
  "graduation_cert_start_number": 1,
  "homeroom_teachers": {
    "1-1": "山田先生"
  },
  "last_loaded_file": "",
  "update": {
    "github_repo": "owner/meibo-tool",
    "check_on_startup": true,
    "current_app_version": "1.0.0",
    "last_check_time": "",
    "skip_version": ""
  },
  "data_source": {
    "mode": "manual",
    "lan_path": "",
    "gdrive_file_id": "",
    "encryption_password": "",
    "last_sync_hash": "",
    "last_sync_time": "",
    "cache_file": ""
  },
  "special_needs_assignments": {},
  "special_needs_placement": "appended"
}
```

| キー | 型 | 説明 |
| ---- | -- | ---- |
| `school_name` | str | 学校名（帳票タイトル `{{学校名}}` に使用） |
| `school_type` | str | `"elementary"` / `"middle"` |
| `fiscal_year` | int | 年度（西暦。起動時に `_current_fiscal_year()` で自動計算） |
| `template_dir` | str | Excel テンプレートフォルダパス |
| `output_dir` | str | 出力フォルダパス |
| `layout_dir` | str | .lay JSON レイアウトライブラリフォルダパス |
| `photo_dir` | str | 児童写真フォルダパス |
| `homeroom_teachers` | dict | `{"1-1": "山田先生"}` 形式で担任名を保存 |
| `update.github_repo` | str | GitHub リポジトリ（`owner/repo` 形式） |
| `data_source.mode` | str | `"manual"` / `"lan"` / `"gdrive"` |
| `data_source.lan_path` | str | LAN 共有フォルダの UNC パス |
| `data_source.gdrive_file_id` | str | Google Drive ファイル ID |
| `data_source.encryption_password` | str | DPAPI で保護された暗号化パスワード |
| `special_needs_assignments` | dict | `{"1-なかよし-1": "1-1"}` 形式で特支児童の交流学級割り当て |
| `special_needs_placement` | str | `"appended"` (末尾追加) / `"integrated"` (出席番号順統合) |

### 1.8 動作環境要件

| 項目 | 要件 |
| ---- | ---- |
| OS | Windows 10 / 11（64bit） |
| Excel | Microsoft Excel 2016 以降 |
| フォント | IPAmj明朝インストール済み |
| ネットワーク | HTTPS 接続可能（オフラインでも帳票生成は可能） |
| ディスク | exe: 50〜100MB、テンプレート+レイアウト: 10MB |
| Python | 不要（exe 化配布） |

---

## 第2章 データモデル

### 2.1 C4th エクスポート Excel — 確定済みヘッダー（全50カラム）

以下は C4th（EDUCOM）からエクスポートされる実際のヘッダー（2026年2月確認済み）。

| # | C4th ヘッダー（実際の列名） | 内部論理名 | 用途・備考 |
| -- | ------------------------ | --------- | --------- |
| 1 | 生徒コード | 生徒コード | C4th 内部 ID。データ突合に使用 |
| 2 | 学年 | 学年 | 1〜6（小）、1〜3（中）。フィルタリングの主キー |
| 3 | 名前 | 氏名 | 通常表示用。名札・名列表等で使用 |
| 4 | ふりがな | 氏名かな | ひらがな表記 |
| 5 | 正式名前 | 正式氏名 | IPAmj明朝用（IVS 付き異体字含む）。公式帳票で使用 |
| 6 | 正式名前ふりがな | 正式氏名かな | 正式ふりがな |
| 7 | 性別 | 性別 | 修了台帳、学級編成用個票で使用 |
| 8 | 生年月日 | 生年月日 | 台帳・個票・家庭調査票で使用。日付型 |
| 9 | 外国籍 | 外国籍 | 学級編成用個票で参考情報 |
| 10 | 郵便番号 | 郵便番号 | 家庭調査票で使用 |
| 11 | 都道府県 | 都道府県 | 住所構成要素(1) |
| 12 | 市区町村 | 市区町村 | 住所構成要素(2) |
| 13 | 町番地 | 町番地 | 住所構成要素(3) |
| 14 | アパート/マンション名 | 建物名 | 住所構成要素(4) |
| 15 | 電話番号1 | 電話番号1 | 主電話番号 |
| 16-18 | 電話番号2/3, FAX番号 | 電話番号2/3, FAX番号 | 副連絡先 |
| 19 | 出身校 | 出身校 | 卒業台帳等で参考 |
| 20-22 | 出身校住所 / 在籍開始日 / 終了日 | — | 帳票では直接使用しない |
| 23 | 入学日 | 入学日 | 参考情報 |
| 24-29 | 転入前学校 / 住所 / 在籍日 / 転入日 / 事由 | 転入日 等 | 修了台帳の「学校をまった日」に使用可能 |
| 30-35 | 編入前学校 / 住所 / 在籍日 / 編入日 / 事由 | — | 参考情報 |
| 36 | 保護者1　続柄 | 保護者続柄 | 家庭調査票で使用（全角スペース区切り） |
| 37 | 保護者1　名前 | 保護者名 | 家庭調査票・台帳で使用 |
| 38 | 保護者1　名前ふりがな | 保護者名かな | 家庭調査票で使用 |
| 39 | 保護者1　正式名前 | 保護者正式名 | IPAmj明朝用。公式帳票の保護者名に使用 |
| 40 | 保護者1　正式名前ふりがな | 保護者正式名かな | 公式帳票のふりがなに使用 |
| 41 | 保護者1　郵便番号 | 保護者郵便番号 | 保護者住所が児童と異なる場合 |
| 42-44 | 保護者1　都道府県/市区町村/町番地 | 保護者住所 | 結合して使用 |
| 45 | 保護者1　アパート/マンション名 | 保護者建物名 | 保護者住所の一部 |
| 46-49 | 保護者1　電話番号1/2/3, FAX番号 | 保護者電話1/2/3, 保護者FAX | 保護者連絡先 |
| 50 | 保護者1　緊急連絡先 | 緊急連絡先 | 家庭調査票の緊急連絡先欄に使用 |

> **落とし穴**: C4th ヘッダーの「保護者1　続柄」は全角スペース（U+3000）で区切られている。`normalize_header()` で統一処理する。

### 2.2 カラムマッピング

`core/mapper.py` に以下の定数が定義されている:

- **EXACT_MAP** (`dict[str, str]`): C4th 確定ヘッダー → 内部論理名の完全一致マップ（50+ エントリ）
- **COLUMN_ALIASES** (`dict[str, list[str]]`): 他校や旧バージョンでの表記揺れ対応（氏名/出席番号等）
- **_FALLBACK_COLUMNS** (`list[tuple[str, str]]`): 正式名フォールバックペア（正式氏名→氏名 等）

主要関数:

| 関数 | シグネチャ | 説明 |
| ---- | --------- | ---- |
| `normalize_header` | `(s: str) -> str` | ヘッダー名を正規化（トリム、全角スペース統一） |
| `map_columns` | `(df) -> tuple[DataFrame, list[str]]` | DataFrame 全カラムをマッピング。戻り値は (変換後 df, 未マッピングカラム一覧) |
| `ensure_fallback_columns` | `(df) -> None` | 正式名系カラムが空なら通常名にフォールバック（in-place） |
| `resolve_name_fields` | `(data_row: dict, use_formal: bool) -> dict[str, str]` | テンプレートの use_formal_name フラグに基づき表示氏名を選択 |

### 2.3 「名前」vs「正式名前」の使い分け

| 帳票 | 使用するフィールド | 理由 |
| ---- | ---------------- | ---- |
| 名札（全3種） | 氏名（名前） | 名札は日常使用のため通常表示で十分 |
| 掲示用名列表 | 氏名（名前） | 教室掲示用 |
| 調べ表 | 氏名（名前） | 日常使用 |
| 修了台帳 | 正式氏名（正式名前） | 公式書類のため正式表記が必須 |
| 卒業台帳 | 正式氏名（正式名前） | 公式書類のため正式表記が必須 |
| 家庭調査票 | 正式氏名（正式名前） | 公式書類のため正式表記が必須 |
| 学級編成用個票 | 正式氏名（正式名前） | 引継ぎ書類のため正式表記 |

### 2.4 住所結合処理

`utils/address.py` の `build_address(row)` が都道府県 + 市区町村 + 町番地 + 建物名を結合する。NaN/空値はスキップされる。

### 2.5 日付フォーマット

`utils/date_fmt.py` の `format_date()` が日付を YY/MM/DD 形式に変換する。対応形式:
- `YYYY-MM-DD`, `YYYY/MM/DD`
- Excel シリアル値（数値）
- `DATE_KEYS`: frozenset — 日付として処理すべきフィールド名のセット

### 2.6 C4th に含まれないカラム

| 必須カラム | C4th に含まれるか | 補完方法 |
| --------- | --------------- | ------- |
| 学年 | ○ 含まれる | C4th から自動読込 |
| 組 | × 含まれない | ClassSelectPanel で教員が選択 |
| 出席番号 | × 含まれない | ClassSelectPanel で自動連番 or 手動設定 |

### 2.7 EditableDataModel（core/data_model.py）

GUI 上でデータを編集可能にするラッパークラス:

| メソッド | シグネチャ | 説明 |
| -------- | --------- | ---- |
| `get_value` | `(row: int, col: str) -> str` | セル値取得 |
| `set_value` | `(row: int, col: str, value: str) -> None` | セル値設定（Undo 履歴に記録） |
| `undo` | `() -> _EditOp \| None` | 直前の編集を元に戻す |
| `redo` | `() -> _EditOp \| None` | 元に戻した編集をやり直す |
| `is_modified` | `() -> bool` | 未保存の変更があるか |

---

## 第3章 GUI 設計

### 3.1 メインウィンドウ

| 項目 | 仕様 |
| ---- | ---- |
| ウィンドウタイトル | 名簿帳票ツール v1.0 |
| 初期サイズ | 900 x 700 px |
| 最小サイズ | 800 x 600 px |
| テーマ | CustomTkinter "blue" テーマ、ライトモード |
| フォント | メイリオ 11pt（UI 全体） |

### 3.2 レイアウト構造

```
App (CTk) — 2カラムレイアウト
├── ヘッダーバー（設定ボタン・レイアウトエディターボタン）
├── 左パネル（CTkScrollableFrame）
│   ├── ImportFrame      — ファイル選択・読込件数・同期ステータス
│   ├── ClassSelectPanel — 学年・組の選択（特別支援学級対応）
│   ├── SelectFrame      — テンプレート選択（カテゴリ別）・担任名・学校名・特支配置
│   ├── OutputFrame      — 生成ボタン・進捗バー
│   └── RosterPrintPanel — レイアウト印刷パネル
└── 右パネル（_PreviewPanel — CTkTabview）
    ├── 「データ」タブ   — Treeview でデータ一覧表示（セル編集可能）
    └── 「プレビュー」タブ — PIL レンダリング画像で帳票プレビュー表示
```

### 3.3 各フレームの責務

| フレーム | ファイル | 主な責務 |
| -------- | ------- | ------- |
| ImportFrame | `gui/frames/import_frame.py` | Excel/CSV ファイル選択、データ同期ステータス表示 |
| ClassSelectPanel | `gui/frames/class_select_panel.py` | 学年・組ドロップダウン、特別支援学級の交流学級割り当て |
| SelectFrame | `gui/frames/select_frame.py` | テンプレート選択（CATEGORY_ORDER でグループ化）、担任名・学校名入力、特支配置モード切替 |
| OutputFrame | `gui/frames/output_frame.py` | 「Excel 生成」ボタン、進捗バー。メインスレッドでジェネレーター作成 |
| RosterPrintPanel | `gui/frames/roster_print_panel.py` | .lay レイアウトの印刷・プレビュー |

### 3.4 ダイアログ

| ダイアログ | ファイル | 用途 |
| ---------- | ------- | ---- |
| MappingDialog | `gui/dialogs/mapping_dialog.py` | C4th カラム ↔ 内部名の手動マッピング調整 |
| SettingsDialog | `gui/dialogs/settings_dialog.py` | 同期モード設定・暗号化パスワード・LAN パス |
| ExchangeClassDialog | `gui/dialogs/exchange_class_dialog.py` | 特支児童の交流学級割り当て |
| UpdateDialog | `gui/dialogs/update_dialog.py` | アプリ更新確認（進捗バー付き） |

### 3.5 レイアウトエディター（gui/editor/）

独立ウィンドウ（CTkToplevel）で .lay レイアウトの視覚的編集を行う:

| ファイル | 責務 |
| ------- | ---- |
| `editor_window.py` | メインウィンドウ（Canvas + プロパティ + Undo/Redo） |
| `layout_canvas.py` | インタラクティブ Canvas（オブジェクト選択・移動・リサイズ） |
| `properties_panel.py` | 選択オブジェクトのプロパティ編集 |
| `toolbar.py` | ファイル操作・ズーム・オブジェクト追加/削除 |
| `print_dialog.py` | プリンター選択・用紙設定・印刷実行 |
| `print_preview_dialog.py` | PILBackend レンダリング + ページ送り |
| `layout_manager_dialog.py` | レイアウトライブラリ管理（Treeview 一覧・インポート・削除・リネーム） |
| `object_list.py` | オブジェクト一覧（TreeView と Canvas の双方向連動） |

> **落とし穴**: CustomTkinter の UI 操作はメインスレッドのみ。バックグラウンドスレッドからは `self.after(0, callback)` で委譲する。

---

## 第4章 テンプレート仕様（Excel 9 種）

### 4.0 差込処理の 3 タイプ

| タイプ | 説明 | 対象テンプレート |
| ------ | ---- | --------------- |
| grid（グリッド配置型） | 1 ページに N 名を `{{氏名_1}}` `{{氏名_2}}` 形式で配置 | 名札3種 + ラベル3種 + 横名簿 + 縦一週間 + 男女一覧 |
| list（リスト展開型） | テンプレート行を児童数分コピー展開 | 修了台帳、卒業台帳 |
| individual（個票複製型） | 1 ページ 1 児童のシートを `copy_worksheet` で複製 | 家庭調査票、学級編成用個票 |

ファクトリー関数 `create_generator()` が `TEMPLATES` レジストリの `type` キーで分岐する。

### 4.1 テンプレート一覧（template_registry.py）

| テンプレート名 | ファイル | タイプ | 向き | 1P最大 | カテゴリ |
| -------------- | ------- | ------ | --- | ------ | ------- |
| ラベル_色付き | ラベル_色付き.xlsx | grid | 縦 | 40名 | 名札・ラベル |
| サンプル_名札 | サンプル_名札.xlsx | grid | 縦 | 8名 | 名札・ラベル |
| 名札_1年生用 | 名札_1年生用.xlsx | grid | 縦 | 8名 | 名札・ラベル |
| ラベル_大2 | ラベル_大2.xlsx | grid | 縦 | 40名 | 名札・ラベル |
| ラベル_小 | ラベル_小.xlsx | grid | 縦 | 20名 | 名札・ラベル |
| ラベル_特大 | ラベル_特大.xlsx | grid | 縦 | 40名 | 名札・ラベル |
| 横名簿 | 横名簿.xlsx | grid | 横 | 1名 | 名簿・出欠表 |
| 縦一週間 | 縦一週間.xlsx | grid | 縦 | 1名 | 名簿・出欠表 |
| 男女一覧 | 男女一覧.xlsx | grid | 縦 | 6名 | 名簿・出欠表 |

> **注**: 台帳・個票は .lay レイアウトシステムに移行済み。Excel テンプレート版は生成スクリプトはあるが非デフォルト。

### 4.2 プレースホルダー形式

テンプレート Excel 内は `{{論理名}}` 形式。`fill_placeholders()` が正規表現で一括置換する。

**特殊キー:**
| キー | 値の出所 |
| ---- | ------- |
| `{{年度}}` | config の fiscal_year（西暦） |
| `{{年度和暦}}` | `wareki.to_wareki()` で自動計算 |
| `{{学校名}}` | config の school_name |
| `{{担任名}}` | GUI で入力 |
| `{{住所}}` | `address.build_address()` で 4 フィールド結合 |

**名前表示モード（name_display）:**

| モード | `{{氏名かな}}` | `{{氏名}}` |
| ------ | ------------- | --------- |
| furigana（デフォルト） | ふりがな値 | 漢字氏名 |
| kanji | 空白 | 漢字氏名 |
| kana | 空白 | ふりがな値 |

### 4.3 ジェネレータークラス（core/generator.py）

```
BaseGenerator(ABC)
  ├── __init__(template_path, output_path, data: DataFrame, options: dict)
  ├── generate() -> str          # テンプレート読込 → _populate → フォント適用 → 保存
  ├── _populate()                 # 抽象メソッド（サブクラスで実装）
  └── _apply_font_all()          # 全シートに IPAmj明朝適用

GridGenerator(BaseGenerator)     # {{氏名_1}} 形式の番号付きプレースホルダーを処理
ListGenerator(BaseGenerator)     # テンプレート行を検出→児童数分コピー展開
IndividualGenerator(BaseGenerator) # fill 前にシート複製（プレースホルダー消失防止）
```

主要ヘルパー関数:

| 関数 | 説明 |
| ---- | ---- |
| `fill_placeholders(ws, data_row, options)` | ワークシート内の全 `{{...}}` を置換 |
| `copy_sheet_with_images(wb, source_ws, new_title)` | 画像も含めたシート複製 |
| `setup_print(ws, orientation, paper_size)` | 印刷設定（余白・fitToPage） |
| `create_generator(template_name, ...)` | ファクトリー関数 |

> **落とし穴**:
> - openpyxl スタイルはイミュータブル。`cell.font.bold = True` は不可。`cell.font = Font(...)` で新インスタンスを作る。
> - `copy_worksheet` は画像をコピーしない。`_data()` → `BytesIO` → `Image()` で再構築する。
> - 結合セルは左上セルにのみ書き込む。他セルは `MergedCell` で read-only。
> - IndividualGenerator のシート複製は fill 前に行う。fill 後だとプレースホルダーが消失する。

### 4.4 印刷設定

```python
def setup_print(ws, orientation='portrait', paper_size=9):
    ws.page_setup.paperSize = paper_size  # 9=A4
    ws.page_setup.orientation = orientation
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0  # 0=自動
    ws.page_margins = PageMargins(
        left=0.39, right=0.39,    # 10mm
        top=0.39, bottom=0.39,
        header=0.2, footer=0.2,
    )
    ws.print_options.horizontalCentered = True
```

### 4.5 Excel 寸法チートシート

| 項目 | openpyxl 単位 | 1mm の値 | 備考 |
| ---- | ------------- | ------- | ---- |
| 列幅 | 文字数（半角基準） | 約 0.35 | フォントに依存 |
| 行高 | ポイント（pt） | 約 2.83 | 1pt = 1/72 inch |
| 余白 | インチ | 約 0.039 | PageMargins |
| 画像幅/高 | ピクセル（px） | 約 3.78 | 96dpi 基準 |

---

## 第5章 特別支援学級（core/special_needs.py）

### 5.1 概要

C4th の「組」列に非数字の値（「なかよし」「ひまわり」等）が入る場合、特別支援学級と判定する。

### 5.2 主要関数

| 関数 | 説明 |
| ---- | ---- |
| `is_special_needs_class(klass)` | クラス名が特支かどうか判定（非数字 = 特支） |
| `detect_special_needs_students(df)` | DataFrame から特支児童を検出 |
| `merge_special_needs_students(df, assignments, placement)` | 特支児童を通常学級に統合 |

### 5.3 交流学級割り当て

初回インポート時に `ExchangeClassDialog` で教員が割り当て → config の `special_needs_assignments` に保存。以降は自動統合。

配置モード（`special_needs_placement`）:
- `"appended"`: 通常学級の末尾に追加
- `"integrated"`: 出席番号順に統合

---

## 第6章 自動更新（アプリ + 名簿データ同期）

### 6.1 概要

2つの「自動更新」機能がある:

1. **アプリ自動更新** — GitHub Releases に新バージョンを公開 → 各校のアプリが起動時に検出
2. **名簿データ同期** — 管理者が名簿データを更新 → 教員のアプリが起動時に自動取得

### 6.2 アプリ自動更新（core/updater.py）

```
check_for_update(config) -> UpdateInfo | None
  GitHub Releases API で最新バージョンを確認
is_newer(remote, local) -> bool
  semver 比較（v-prefix 対応）
download_release_asset(url, dest, progress_cb)
  Release asset をストリーミングダウンロード
generate_update_batch(zip_path, current_dir)
  --onedir フォルダ更新バッチ生成→実行→sys.exit(0)
```

更新フロー:
```
起動時 [バックグラウンドスレッド]:
  GET api.github.com/repos/{owner}/{repo}/releases/latest (timeout 5s)
  → semver 比較 → 新版あり → メインスレッドへ通知 → UpdateDialog 表示
    → 「今すぐ更新」: zip ダウンロード → バッチファイル生成 → アプリ終了
    → 「後で」: 次回起動時に再確認
    → 「スキップ」: config に skip_version を保存
  → ネットワークエラー → サイレントにスキップ
```

> **落とし穴**: バッチの `ren` コマンドはスペース含むパスで失敗するため `move` を使用。`chcp 65001` は日本語パスの文字化け対策で必須。

### 6.3 名簿データ同期（core/data_sync.py）

3モード対応:

| モード | 説明 |
| ------ | ---- |
| manual | 従来の手動ファイル選択。同期処理なし |
| lan | 共有フォルダの Excel をハッシュ比較 → 変更ありなら cache にコピー → 自動インポート |
| gdrive | Google Drive からダウンロード → AES-256-GCM 復号 → cache に保存 → インポート |

スレッド安全性: `sync()` はバックグラウンドスレッドから呼ばれるため、config を直接変更しない。変更内容は `SyncResult.config_updates` に格納し、メインスレッドで適用する。

### 6.4 暗号化（core/crypto.py）

- **アルゴリズム**: AES-256-GCM（`cryptography` の AESGCM）
- **鍵導出**: PBKDF2-HMAC-SHA256、600,000 イテレーション（OWASP 2023 推奨）
- **ファイル形式**: `[4B magic: MBE1][16B salt][12B nonce][ciphertext + GCM tag]`
- **config 内パスワード保護**: Windows DPAPI（`ctypes` 経由、pywin32 不要）。非 Windows は base64 フォールバック。

### 6.5 CI/CD（.github/workflows/build-release.yml）

```
トリガー: v* タグの push
→ Windows-latest / Python 3.11
→ pip install → ruff check → pytest → pyinstaller build.spec
→ zip 化 → softprops/action-gh-release@v2 で Release 作成
```

---

## 第7章 ビルド・配布

### 7.1 PyInstaller ビルド（--onedir 必須）

> **最重要**: CustomTkinter は .json/.otf などの非 Python データファイルを含むため `--onefile` では正常動作しない。必ず `--onedir` を使用する。

ビルドコマンド: `pyinstaller build.spec`

出力: `dist/名簿帳票ツール/` フォルダ。配布時はフォルダ全体を共有する。

### 7.2 テスト構成

```
tests/
  ├── conftest.py              # 共通フィクスチャ（dummy_df, default_options, default_config）
  ├── test_wareki.py           # 和暦変換（10テスト）
  ├── test_address.py          # 住所結合（12テスト）
  ├── test_date_fmt.py         # 日付フォーマット（22テスト）
  ├── test_font_helper.py      # フォント適用（11テスト）
  ├── test_config.py           # config 管理（13テスト）
  ├── test_mapper.py           # カラムマッピング（6テスト）
  ├── test_mapper_resolve.py   # 名前フィールド解決（6テスト）
  ├── test_importer.py         # Excel/CSV 読込（26テスト）
  ├── test_generator.py        # GridGenerator/fill_placeholders（23テスト）
  ├── test_list_individual_gen.py # List/Individual ジェネレーター（6テスト）
  ├── test_gen_nafuda.py       # 名札テンプレート生成（16テスト）
  ├── test_gen_from_lay.py     # .lay → Excel 変換（27テスト）
  ├── test_template_registry.py # テンプレートメタデータ（28テスト）
  ├── test_template_scanner.py # テンプレート自動検出（29テスト）
  ├── test_crypto.py           # AES-256-GCM（17テスト）
  ├── test_data_sync.py        # データ同期（15テスト）
  ├── test_data_model.py       # EditableDataModel（23テスト）
  ├── test_exporter.py         # Excel エクスポート（6テスト）
  ├── test_updater.py          # 自動更新（33テスト）
  ├── test_special_needs.py    # 特支学級（36テスト）
  ├── test_photo_manager.py    # 写真管理（46テスト）
  ├── test_lay_parser.py       # .lay パーサー（74テスト）
  ├── test_lay_serializer.py   # JSON シリアライザ（44テスト）
  ├── test_lay_renderer.py     # レイアウト描画（101テスト）
  ├── test_pil_backend.py      # PIL バックエンド（22テスト）
  ├── test_layout_registry.py  # レイアウトライブラリ（43テスト）
  ├── test_win_printer.py      # Windows GDI 印刷（15テスト）
  ├── test_app_logic.py        # App 統合（35テスト）
  ├── test_preview_renderer.py # プレビュー描画（20テスト）
  ├── test_mapping_dialog.py   # マッピングダイアログ（13テスト）
  ├── test_roster_print_panel.py # 印刷パネル（22テスト）
  └── test_integration.py      # 統合テスト（6テスト）
```

**共通フィクスチャ（conftest.py）:**
- `dummy_df()` — 5名分のテスト DataFrame
- `default_options()` — ジェネレーターオプション dict
- `default_config()` — config dict

---

## 第8章 .lay レイアウトシステム

### 8.1 概要

スズキ校務（旧校務支援システム）の .lay バイナリレイアウト形式をパース・描画・編集・印刷する機能。27 種のレイアウト帳票をサポートする。

### 8.2 .lay バイナリ形式（core/lay_parser.py）

**ファイル形式**: EXCMIDataContainer01

```
[ヘッダー "EXCMIDataContainer01"]
[zlib 圧縮データ]
  → 解凍後: TLV (Tag-Length-Value) の再帰構造
```

**主要データ構造:**

```python
ObjectType(IntEnum):
  LINE = 1, GROUP = 2, LABEL = 3, FIELD = 4,
  TABLE = 5, MEIBO = 6, IMAGE = 7

@dataclass FontInfo:
  name: str, size_pt: float, bold: bool, italic: bool, vertical: bool

@dataclass LayoutObject:
  obj_type: ObjectType
  rect: Rect | None           # LABEL/FIELD/TABLE/MEIBO/IMAGE/GROUP 用
  line_start/end: Point | None # LINE 用
  text: str                    # LABEL のテキスト
  field_id: int                # FIELD の参照 ID
  font: FontInfo
  table_columns: list[TableColumn]  # TABLE の列定義
  meibo: MeiboArea | None      # MEIBO の展開情報
  style_1001/1002/1003: int | None  # style タグ
  raw_tags: list[RawTag]       # 生 TLV 完全保持

@dataclass LayFile:
  title: str, version: int
  page_width/height: int       # 座標単位
  objects: list[LayoutObject]
  paper: PaperLayout | None
  raw_tags: list[RawTag]
```

**座標系:**

| モード | 単位 | A4 幅 | 例 |
| ------ | --- | ----- | -- |
| 単一レイアウト | 0.25mm/unit | 840 units = 210mm | `unit_mm=0.25` |
| マルチレイアウト | 0.1mm/unit | 2100 units = 210mm | `unit_mm=0.1` |

**用紙サイズ検出（PaperLayout）:**
- `mode=0`: フルページ（A4/A3/B4 を page_width/height から検出）
- `mode=1`: ラベルサイズ（item_width × item_height で複数面付け）
- `DOC_FLAG2=2`: 横向きオーバーライド

### 8.3 FIELD_ID_MAP

.lay レイアウト内のフィールド ID → C4th 内部論理名のマッピング。60+ エントリ。

主要マッピング:

| ID 範囲 | 内容 |
| ------- | ---- |
| 100-106 | 生徒コード、学年、組、在籍、組表示、出席番号 |
| 107-109 | 性別、氏名、氏名かな |
| 110-113 | 年度和暦、学校名、学校住所、担任名 |
| 134, 137 | 年度西暦、転出日 |
| 400 | 写真 |
| 601-683 | 在校兄弟、保護者名、都道府県、市区町村、町番地、緊急連絡先 等 |
| 1506-1542 | 欠席、ピアノ、要保護、家庭環境、配慮児童、アレルギー、引継事項 等 |

詳細は `core/lay_parser.py` の `FIELD_ID_MAP` を参照。

### 8.4 JSON シリアライザ（core/lay_serializer.py）

LayFile ↔ JSON の保存/読込。以下をラウンドトリップ保証:
- bold/italic/vertical フォント属性
- style_1001/1002/1003
- raw_tags（payload は base64 エンコード）

主要関数:
- `save_layout(lay: LayFile, path: str) -> None`
- `load_layout(path: str) -> LayFile`
- `layfile_to_dict(lay: LayFile) -> dict`
- `dict_to_layfile(d: dict) -> LayFile`

### 8.5 レンダラー（core/lay_renderer.py）

2つの描画バックエンド:

| バックエンド | 用途 | 出力 |
| ------------ | ---- | ---- |
| CanvasBackend | GUI プレビュー・レイアウトエディター | tk.Canvas オブジェクト |
| PILBackend | 印刷プレビュー・画像生成 | PIL Image |

**主要関数:**

| 関数 | 説明 |
| ---- | ---- |
| `render_layout_to_image(lay, dpi, ...)` | LayFile → PIL Image |
| `fill_layout(lay, data_row, options)` | FIELD オブジェクトにデータを差し込み |
| `fill_meibo_layout(lay, data_rows, options, registry)` | MEIBO 展開（児童数分の繰り返し）+ FIELD 差込 |
| `has_meibo(lay)` | レイアウトに MEIBO オブジェクトがあるか |
| `get_page_arrangement(lay)` | 用紙レイアウトの面付け計算 |
| `tile_layouts(layouts, cols, rows, ...)` | 複数レイアウトを用紙に面付け |

**描画対応:**
- LABEL: テキスト描画（縦書き、自動折り返し、style_1002==10 で枠線）
- FIELD: データフィールド描画（外字は IPAmj明朝、他はレイアウト元フォント）
- LINE: 直線描画
- GROUP: 矩形描画（4辺の LINE に展開せず直接描画）
- TABLE: テーブル描画（ヘッダー + 列定義）
- MEIBO: 繰り返し領域展開（サブレイアウト参照解決）
- IMAGE: 埋め込み画像描画

### 8.6 レイアウトライブラリ管理（core/layout_registry.py）

`レイアウト/` フォルダ内の JSON ファイルを管理。

主要関数:

| 関数 | 説明 |
| ---- | ---- |
| `scan_layouts(layout_dir)` | フォルダ内の .json をスキャンして一覧取得 |
| `import_lay_file(lay_path, layout_dir)` | .lay ファイルをパース → JSON で保存 |
| `delete_layout(name, layout_dir)` | レイアウト削除 |
| `rename_layout(old, new, layout_dir)` | レイアウトリネーム（raw_tags 保持） |

`_SUZUKI_REF_ALIASES`: スズキ校務の MEIBO 参照名とレイアウト名の不一致をエイリアスで解決する辞書。

### 8.7 Windows GDI 印刷（core/win_printer.py）

Windows GDI API（ctypes 経由）でレイアウトを直接印刷。縦書き・複数行・bold/italic 対応。

### 8.8 デフォルトレイアウト 27 種

初回起動時に `resources/default_layouts.lay` から自動インポート。カテゴリ:

| カテゴリ | レイアウト例 | 数 |
| ------- | ---------- | -- |
| 名札・ラベル | 机用名札（装飾あり/なし）、1年生用 | 3 |
| 掲示用 | 掲示用名列表（通常/ひらがな/ふりがな付き） | 3 |
| 調べ表 | 調べ表（通常/ひらがな） | 2 |
| 台帳 | 修了台帳、卒業台帳 | 2 |
| 個票 | 家庭調査票、学級編成用個票 | 2 |
| 名簿 | 学年名簿(50音)、学級名簿 連絡先一覧、複数学級名列表 等 | 6+ |
| その他 | はがき宛名、写真ラベル、写真名簿、職員用教室掲示 等 | 9+ |

---

## 付録

### A. コマンドリファレンス

```bash
# 仮想環境セットアップ（初回のみ）
python -m venv venv
venv/Scripts/activate

# 依存インストール
pip install -r requirements.txt

# アプリ起動
cd meibo_tool && python main.py

# テスト全実行（プロジェクトルートから）
venv/Scripts/python.exe -m pytest

# 単一テストファイル実行
venv/Scripts/python.exe -m pytest meibo_tool/tests/test_wareki.py -v

# リント
venv/Scripts/python.exe -m ruff check meibo_tool/
venv/Scripts/python.exe -m ruff check meibo_tool/ --fix

# テンプレート Excel 生成
cd meibo_tool && python -m templates.generators.generate_all

# exe ビルド
pyinstaller build.spec
```

### B. 落とし穴一覧

| 場所 | 落とし穴 | 対策 |
| ---- | ------- | ---- |
| openpyxl | `cell.font.bold = True` → AttributeError | `cell.font = Font(bold=True)` で新インスタンス |
| openpyxl | `copy_worksheet` は画像をコピーしない | `_data()` → `BytesIO` → `Image()` で再構築 |
| openpyxl | 結合セルの左上以外に書き込み → AttributeError | `MergedCell` チェック |
| openpyxl | `textRotation=255` のみが縦書き | 0〜180 は回転角度 |
| CustomTkinter | バックグラウンドスレッドから UI 操作 → クラッシュ | `self.after(0, callback)` で委譲 |
| generator | fill 後にシートコピーするとプレースホルダー消失 | 全シート複製後に fill |
| C4th | 「保護者1　続柄」は全角スペース（U+3000） | `normalize_header()` で統一 |
| PyInstaller | `--onefile` では .json/.otf が見つからない | `--onedir` のみ使用 |
| バッチファイル | `ren` はスペース含むパスで失敗 | `move` コマンドを使用 |
| バッチファイル | 日本語パスが文字化け | `chcp 65001` 必須 |
| .lay | マルチ/単一で座標単位が異なる | 0.1mm vs 0.25mm |
| .lay | MEIBO 参照名とレイアウト名の不一致 | `_SUZUKI_REF_ALIASES` |
| .lay | GROUP を LINE に展開すると情報が失われる | `ObjectType.GROUP` として保持 |

### C. 将来拡張

- 別 PC（Python なし環境）での exe 起動確認
- 保護者 2 以降への対応
- 他帳票の追加（通知表、出席簿、保健関連等）
- 複数校一括処理（教育委員会向け）
- C4th API 直接連携（将来の API 公開時）
