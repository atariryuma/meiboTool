"""全テンプレート Excel を一括生成するエントリースクリプト

使用方法（meibo_tool/ ディレクトリから実行）:
    python -m templates.generators.generate_all

出力先は config.json の template_dir（デフォルト: プロジェクトルートの テンプレート/）。
"""

from __future__ import annotations

import sys
from pathlib import Path

from core.config import get_template_dir, load_config
from templates.generators import (
    gen_daicho,
    gen_gakkyuu_kojihyo,
    gen_katei_chousahyo,
    gen_meireihyo,
    gen_nafuda,
    gen_shirabehyo,
)
from templates.generators.gen_from_legacy import generate as _gen_from_legacy

# 元の .xls ファイルパス（プロジェクトルートに配置）
_XLS_PATH = str(Path(__file__).resolve().parents[3] / '様式R2年度名簿 型.xls')


def main() -> None:
    config = load_config()
    template_dir = get_template_dir(config)

    # ── 掲示用名列表（3モード共通テンプレート 1 種）─────────────────────────
    generators = [
        ('掲示用名列表.xlsx', gen_meireihyo.generate),
        # ── 1年生用名札（Phase 2 の独自実装を継続使用）──────────────────────
        ('名札_1年生用.xlsx',  lambda p: gen_nafuda.generate(p, mode='1年生')),
        # ── 調べ表（Phase 2） ────────────────────────────────────────────
        ('調べ表.xlsx', gen_shirabehyo.generate),
        # ── 台帳 2 種（Phase 2） ─────────────────────────────────────────
        ('修了台帳.xlsx', lambda p: gen_daicho.generate(p, mode='shuuryo')),
        ('卒業台帳.xlsx',  lambda p: gen_daicho.generate(p, mode='sotsugyou')),
        # ── 個票 2 種（Phase 4） ────────────────────────────────────────
        ('家庭調査票.xlsx',    gen_katei_chousahyo.generate),
        ('学級編成用個票.xlsx', gen_gakkyuu_kojihyo.generate),
    ]

    print(f'出力先: {template_dir}')
    for filename, gen_func in generators:
        out_path = str(Path(template_dir) / filename)
        try:
            gen_func(out_path)
        except Exception as e:
            print(f'ERROR [{filename}]: {e}', file=sys.stderr)

    # ── xls 由来テンプレート（名札_通常 / 名札_装飾あり / ラベル各種 / 横名簿 / 縦一週間）
    try:
        _gen_from_legacy(template_dir, _XLS_PATH)
    except Exception as e:
        print(f'ERROR [gen_from_legacy]: {e}', file=sys.stderr)

    print('完了。')


if __name__ == '__main__':
    main()
