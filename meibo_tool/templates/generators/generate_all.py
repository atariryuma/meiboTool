"""全テンプレート Excel を一括生成するエントリースクリプト

使用方法（meibo_tool/ ディレクトリから実行）:
    python -m templates.generators.generate_all

出力先は config.json の template_dir（デフォルト: プロジェクトルートの テンプレート/）。
"""

from __future__ import annotations

import sys
from pathlib import Path

from core.config import get_template_dir, load_config
from templates.generators import gen_meireihyo


def main() -> None:
    config = load_config()
    template_dir = get_template_dir(config)

    generators = [
        ('掲示用名列表.xlsx', gen_meireihyo.generate),
        # 以降のテンプレートは Phase 2 で追加
        # ('名札_通常.xlsx',     gen_nafuda_plain.generate),
        # ('名札_装飾あり.xlsx', gen_nafuda_decorated.generate),
        # ('名札_1年生用.xlsx',  gen_nafuda_1nen.generate),
        # ('調べ表.xlsx',        gen_shirabehyo.generate),
        # ('修了台帳.xlsx',      gen_shuuryo_daicho.generate),
        # ('卒業台帳.xlsx',      gen_sotsugyou_daicho.generate),
        # ('家庭調査票.xlsx',    gen_katei_chousahyo.generate),
        # ('学級編成用個票.xlsx', gen_gakkyuu_kojihyo.generate),
    ]

    print(f'出力先: {template_dir}')
    for filename, gen_func in generators:
        out_path = str(Path(template_dir) / filename)
        try:
            gen_func(out_path)
        except Exception as e:
            print(f'ERROR [{filename}]: {e}', file=sys.stderr)

    print('完了。')


if __name__ == '__main__':
    main()
