"""テスト用ダミー C4th データ生成スクリプト

SPEC.md 第7章 §7.10 参照。
実行: python meibo_tool/tests/generate_dummy.py
→ meibo_tool/tests/fixtures/dummy_c4th.xlsx が生成される
"""

import pandas as pd
import random
import os

LAST_NAMES = ['山田', '田中', '鈴木', '佐藤', '高橋', '伊藤', '渡辺', '中村', '小林', '加藤',
              '新垣', '比嘉', '宮城', '金城', '玉城']
FIRST_NAMES_M = ['太郎', '一郎', '健太', '翔', '大輔', '拓也', '修', '隆', '誠', '光']
FIRST_NAMES_F = ['花子', '美咲', '陽菜', '結衣', 'さくら', '凛', '葵', '彩', '楓', '愛']
FIRST_KANA_M = ['たろう', 'いちろう', 'けんた', 'しょう', 'だいすけ',
                'たくや', 'おさむ', 'たかし', 'まこと', 'ひかる']
FIRST_KANA_F = ['はなこ', 'みさき', 'ひな', 'ゆい', 'さくら',
                'りん', 'あおい', 'あや', 'かえで', 'あい']
AREAS = ['天久', '古島', '真地', '小禄', '壺川']
GUARDIANS_M = ['一郎', '健一', '浩二', '雄一', '正則']
GUARDIANS_F = ['幸子', '美穂', '典子', '恵子', '和子']


def generate_dummy(n: int = 35, grade: int = 1, seed: int = 42) -> pd.DataFrame:
    random.seed(seed)
    rows = []
    for i in range(n):
        sex = random.choice(['男', '女'])
        sei = random.choice(LAST_NAMES)
        sei_k = 'やまだ'  # 簡略化

        if sex == '男':
            idx = random.randint(0, len(FIRST_NAMES_M) - 1)
            mei, mei_k = FIRST_NAMES_M[idx], FIRST_KANA_M[idx]
            g_mei = random.choice(GUARDIANS_F)  # 保護者は母親が多い
        else:
            idx = random.randint(0, len(FIRST_NAMES_F) - 1)
            mei, mei_k = FIRST_NAMES_F[idx], FIRST_KANA_F[idx]
            g_mei = random.choice(GUARDIANS_F)

        area = random.choice(AREAS)
        chome = f'{random.randint(1,3)}-{random.randint(1,20)}-{random.randint(1,30)}'
        birth_month = random.randint(4, 12) if grade == 1 else random.randint(1, 12)

        rows.append({
            '生徒コード': f'S{grade:02d}{i+1:03d}',
            '学年': str(grade),
            '名前': f'{sei} {mei}',
            'ふりがな': f'{sei_k} {mei_k}',
            '正式名前': f'{sei} {mei}',
            '正式名前ふりがな': f'{sei_k} {mei_k}',
            '性別': sex,
            '生年月日': f'20{19-grade:02d}-{birth_month:02d}-{random.randint(1,28):02d}',
            '外国籍': '',
            '郵便番号': f'900-{random.randint(1000,9999)}',
            '都道府県': '沖縄県',
            '市区町村': '那覇市',
            '町番地': f'{area}{chome}',
            'アパート/マンション名': '',
            '電話番号1': f'098-{random.randint(100,999)}-{random.randint(1000,9999)}',
            '電話番号2': '',
            '電話番号3': '',
            'FAX番号': '',
            '出身校': '',
            '入学日': f'20{24-grade:02d}-04-01',
            '転入日': '',
            f'保護者1\u3000続柄': '母',
            f'保護者1\u3000名前': f'{sei} {g_mei}',
            f'保護者1\u3000名前ふりがな': f'{sei_k} かずこ',
            f'保護者1\u3000正式名前': f'{sei} {g_mei}',
            f'保護者1\u3000正式名前ふりがな': f'{sei_k} かずこ',
            f'保護者1\u3000郵便番号': '',
            f'保護者1\u3000都道府県': '',
            f'保護者1\u3000市区町村': '',
            f'保護者1\u3000町番地': '',
            f'保護者1\u3000アパート/マンション名': '',
            f'保護者1\u3000電話番号1': f'090-{random.randint(1000,9999)}-{random.randint(1000,9999)}',
            f'保護者1\u3000電話番号2': '',
            f'保護者1\u3000電話番号3': '',
            f'保護者1\u3000FAX番号': '',
            f'保護者1\u3000緊急連絡先': f'090-{random.randint(1000,9999)}-{random.randint(1000,9999)}',
        })
    return pd.DataFrame(rows)


if __name__ == '__main__':
    out_dir = os.path.join(os.path.dirname(__file__), 'fixtures')
    os.makedirs(out_dir, exist_ok=True)
    df = generate_dummy(n=35, grade=1)
    out_path = os.path.join(out_dir, 'dummy_c4th.xlsx')
    df.to_excel(out_path, index=False)
    print(f'生成完了: {out_path}  ({len(df)} 名)')
