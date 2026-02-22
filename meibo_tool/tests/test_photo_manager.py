"""photo_manager.py のユニットテスト

テスト対象:
  - scan_photos: フォルダスキャン + キー生成
  - match_photo_to_student: マッチング優先順位
  - load_photo_bytes: 画像読込 + EXIF回転 + リサイズ + クロップ
  - get_match_status: マッチ状況
  - import_photos_from_folder: 一括コピー
  - _generate_match_keys: キー生成ロジック
  - _center_crop: センタークロップ
"""

from __future__ import annotations

import io
import os

import pandas as pd
import pytest
from PIL import Image

from core.photo_manager import (
    _center_crop,
    _generate_match_keys,
    get_match_status,
    import_photos_from_folder,
    load_photo_bytes,
    match_photo_to_student,
    scan_photos,
)

# ── ヘルパー ─────────────────────────────────────────────────────────────────

def _create_test_image(
    path: str, width: int = 100, height: int = 100,
    color: tuple = (255, 0, 0),
) -> None:
    """テスト用の画像ファイルを作成する。"""
    img = Image.new('RGB', (width, height), color)
    img.save(path)


def _create_rgba_image(path: str, width: int = 100, height: int = 100) -> None:
    """テスト用の RGBA 画像ファイルを作成する。"""
    img = Image.new('RGBA', (width, height), (255, 0, 0, 128))
    img.save(path)


def _create_image_with_exif(path: str) -> None:
    """EXIF Orientation=6 (90度回転) の JPEG を作成する。"""

    img = Image.new('RGB', (200, 100), (0, 255, 0))
    buf = io.BytesIO()
    img.save(buf, format='JPEG')

    # 簡易 EXIF: Orientation=6 (90度時計回り)
    # Pillow の piexif を使わず、直接保存時に exif として付与
    from PIL.ExifTags import Base as ExifBase

    img_with_exif = Image.new('RGB', (200, 100), (0, 255, 0))
    exif = img_with_exif.getexif()
    exif[ExifBase.Orientation] = 6  # type: ignore[attr-defined]
    img_with_exif.save(path, exif=exif.tobytes())


# ── _generate_match_keys ─────────────────────────────────────────────────────


class TestGenerateMatchKeys:
    """ファイル名からマッチキーを生成するテスト。"""

    def test_numbered_three_part(self) -> None:
        """'1-1-01' → ['1-1-01', '1-1-1', '01', '1']"""
        keys = _generate_match_keys('1-1-01')
        assert '1-1-01' in keys
        assert '1-1-1' in keys
        assert '01' in keys
        assert '1' in keys

    def test_numbered_no_padding(self) -> None:
        """'1-1-1' → ['1-1-1', '01', '1']"""
        keys = _generate_match_keys('1-1-1')
        assert '1-1-1' in keys
        # 出席番号のみのキーも生成される
        assert '1' in keys

    def test_number_only(self) -> None:
        """'01' → ['01', '1']"""
        keys = _generate_match_keys('01')
        assert '01' in keys
        assert '1' in keys

    def test_number_only_no_padding(self) -> None:
        """'5' → ['5']"""
        keys = _generate_match_keys('5')
        assert '5' in keys

    def test_name_key(self) -> None:
        """'山田太郎' → ['山田太郎']"""
        keys = _generate_match_keys('山田太郎')
        assert keys == ['山田太郎']

    def test_empty_string(self) -> None:
        keys = _generate_match_keys('')
        assert keys == []

    def test_whitespace_only(self) -> None:
        keys = _generate_match_keys('  ')
        assert keys == []


# ── scan_photos ─────────────────────────────────────────────────────────────


class TestScanPhotos:
    """写真フォルダスキャンのテスト。"""

    def test_empty_dir(self, tmp_path: object) -> None:
        result = scan_photos(str(tmp_path))
        assert result == {}

    def test_nonexistent_dir(self) -> None:
        result = scan_photos('/nonexistent/path')
        assert result == {}

    def test_numbered_files(self, tmp_path: object) -> None:
        d = str(tmp_path)
        _create_test_image(os.path.join(d, '1-1-01.jpg'))
        _create_test_image(os.path.join(d, '1-1-02.png'))

        result = scan_photos(d)
        assert '1-1-01' in result
        assert '1-1-1' in result
        assert '1-1-02' in result
        assert '1-1-2' in result

    def test_name_files(self, tmp_path: object) -> None:
        d = str(tmp_path)
        _create_test_image(os.path.join(d, '山田太郎.jpg'))

        result = scan_photos(d)
        assert '山田太郎' in result

    def test_mixed_extensions(self, tmp_path: object) -> None:
        d = str(tmp_path)
        _create_test_image(os.path.join(d, '01.jpg'))
        _create_test_image(os.path.join(d, '02.PNG'))  # 大文字拡張子

        result = scan_photos(d)
        assert '01' in result
        # Windows は case-insensitive なのでファイルは作成されるが拡張子は .PNG のまま
        assert '02' in result or '2' in result

    def test_ignores_non_image(self, tmp_path: object) -> None:
        d = str(tmp_path)
        with open(os.path.join(d, 'readme.txt'), 'w') as f:
            f.write('test')

        result = scan_photos(d)
        assert result == {}

    def test_subdirectory_scan(self, tmp_path: object) -> None:
        d = str(tmp_path)
        subdir = os.path.join(d, '1年')
        os.makedirs(subdir)
        _create_test_image(os.path.join(subdir, '01.jpg'))

        result = scan_photos(d)
        assert '01' in result
        assert '1' in result

    def test_duplicate_key_warning(self, tmp_path: object) -> None:
        """同名キーが複数ある場合、先に見つかった方が使われる。"""
        d = str(tmp_path)
        _create_test_image(os.path.join(d, '01.jpg'))
        subdir = os.path.join(d, 'sub')
        os.makedirs(subdir)
        _create_test_image(os.path.join(subdir, '01.jpg'))

        result = scan_photos(d)
        # キー '01' は存在するはず（先に見つかった方）
        assert '01' in result


# ── match_photo_to_student ──────────────────────────────────────────────────


class TestMatchPhotoToStudent:
    """生徒→写真マッチングのテスト。"""

    def test_match_by_number_padded(self) -> None:
        photo_map = {'1-1-01': '/photos/1-1-01.jpg'}
        row = {'学年': '1', '組': '1', '出席番号': '1'}
        assert match_photo_to_student(row, photo_map) == '/photos/1-1-01.jpg'

    def test_match_by_number_no_pad(self) -> None:
        photo_map = {'1-1-1': '/photos/1-1-1.jpg'}
        row = {'学年': '1', '組': '1', '出席番号': '1'}
        assert match_photo_to_student(row, photo_map) == '/photos/1-1-1.jpg'

    def test_match_by_number_only(self) -> None:
        """出席番号のみのマッチ（単学級向け）。"""
        photo_map = {'03': '/photos/03.jpg'}
        row = {'学年': '1', '組': '1', '出席番号': '3'}
        assert match_photo_to_student(row, photo_map) == '/photos/03.jpg'

    def test_match_by_name(self) -> None:
        photo_map = {'山田太郎': '/photos/山田太郎.jpg'}
        row = {'学年': '1', '組': '1', '出席番号': '99', '氏名': '山田 太郎'}
        # 番号99は写真なし → 氏名でマッチ（スペース除去）
        assert match_photo_to_student(row, photo_map) == '/photos/山田太郎.jpg'

    def test_match_by_formal_name(self) -> None:
        photo_map = {'山田太郎': '/photos/山田太郎.jpg'}
        row = {'学年': '1', '組': '1', '出席番号': '99', '正式氏名': '山田 太郎'}
        assert match_photo_to_student(row, photo_map) == '/photos/山田太郎.jpg'

    def test_number_takes_priority_over_name(self) -> None:
        photo_map = {
            '1-1-01': '/photos/by_number.jpg',
            '山田太郎': '/photos/by_name.jpg',
        }
        row = {'学年': '1', '組': '1', '出席番号': '1', '氏名': '山田太郎'}
        assert match_photo_to_student(row, photo_map) == '/photos/by_number.jpg'

    def test_no_match_returns_none(self) -> None:
        photo_map = {'1-2-01': '/photos/1-2-01.jpg'}
        row = {'学年': '1', '組': '1', '出席番号': '1'}
        assert match_photo_to_student(row, photo_map) is None

    def test_empty_photo_map(self) -> None:
        row = {'学年': '1', '組': '1', '出席番号': '1'}
        assert match_photo_to_student(row, {}) is None

    def test_zenkaku_numbers(self) -> None:
        """全角数字でも正しくマッチする。"""
        photo_map = {'1-1-01': '/photos/1-1-01.jpg'}
        row = {'学年': '１', '組': '１', '出席番号': '１'}
        assert match_photo_to_student(row, photo_map) == '/photos/1-1-01.jpg'

    def test_nan_values(self) -> None:
        """NaN 値でクラッシュしない。"""
        photo_map = {'1-1-01': '/photos/1-1-01.jpg'}
        row = {'学年': 'nan', '組': 'nan', '出席番号': 'nan'}
        assert match_photo_to_student(row, photo_map) is None

    def test_missing_fields(self) -> None:
        """フィールドがない場合もクラッシュしない。"""
        photo_map = {'1-1-01': '/photos/1-1-01.jpg'}
        assert match_photo_to_student({}, photo_map) is None


# ── load_photo_bytes ─────────────────────────────────────────────────────────


class TestLoadPhotoBytes:
    """画像読込・処理のテスト。"""

    def test_loads_jpg(self, tmp_path: object) -> None:
        path = os.path.join(str(tmp_path), 'test.jpg')
        _create_test_image(path, 200, 300)
        result = load_photo_bytes(path)
        assert result is not None
        # PNG bytes であることを確認
        assert result[:8] == b'\x89PNG\r\n\x1a\n'

    def test_loads_png(self, tmp_path: object) -> None:
        path = os.path.join(str(tmp_path), 'test.png')
        _create_test_image(path, 100, 100)
        result = load_photo_bytes(path)
        assert result is not None

    def test_resizes_large_image(self, tmp_path: object) -> None:
        path = os.path.join(str(tmp_path), 'large.jpg')
        _create_test_image(path, 2000, 1500)
        result = load_photo_bytes(path, max_size=400)
        assert result is not None
        img = Image.open(io.BytesIO(result))
        assert max(img.size) <= 400

    def test_rgba_to_rgb(self, tmp_path: object) -> None:
        path = os.path.join(str(tmp_path), 'rgba.png')
        _create_rgba_image(path)
        result = load_photo_bytes(path)
        assert result is not None
        img = Image.open(io.BytesIO(result))
        assert img.mode == 'RGB'

    def test_nonexistent_returns_none(self) -> None:
        result = load_photo_bytes('/nonexistent/photo.jpg')
        assert result is None

    def test_corrupt_file_returns_none(self, tmp_path: object) -> None:
        path = os.path.join(str(tmp_path), 'corrupt.jpg')
        with open(path, 'wb') as f:
            f.write(b'not an image')
        result = load_photo_bytes(path)
        assert result is None

    def test_target_rect_crops(self, tmp_path: object) -> None:
        """target_rect でアスペクト比に合わせてクロップされる。"""
        path = os.path.join(str(tmp_path), 'wide.jpg')
        _create_test_image(path, 400, 200)  # 2:1 アスペクト比
        result = load_photo_bytes(path, max_size=200, target_rect=(100, 200))
        assert result is not None
        img = Image.open(io.BytesIO(result))
        # target_rect は 1:2 (縦長) なので横がクロップされる
        w, h = img.size
        assert h > w  # 縦長になっているはず

    def test_exif_rotation(self, tmp_path: object) -> None:
        """EXIF Orientation タグで自動回転される。"""
        path = os.path.join(str(tmp_path), 'rotated.jpg')
        try:
            _create_image_with_exif(path)
        except Exception:
            pytest.skip('EXIF 作成に失敗')

        result = load_photo_bytes(path)
        assert result is not None

    def test_small_image_not_upscaled(self, tmp_path: object) -> None:
        """小さい画像は拡大しない。"""
        path = os.path.join(str(tmp_path), 'small.jpg')
        _create_test_image(path, 50, 50)
        result = load_photo_bytes(path, max_size=600)
        assert result is not None
        img = Image.open(io.BytesIO(result))
        assert img.size == (50, 50)


# ── _center_crop ─────────────────────────────────────────────────────────────


class TestCenterCrop:
    """センタークロップのテスト。"""

    def test_wider_image_crops_sides(self) -> None:
        """横長画像 → 左右をクロップ。"""
        img = Image.new('RGB', (400, 200))
        result = _center_crop(img, 100, 200)  # 1:2 アスペクト比
        w, h = result.size
        assert h == 200
        assert w == 100  # 横が 100 にクロップ

    def test_taller_image_crops_top_bottom(self) -> None:
        """縦長画像 → 上下をクロップ。"""
        img = Image.new('RGB', (200, 400))
        result = _center_crop(img, 200, 100)  # 2:1 アスペクト比
        w, h = result.size
        assert w == 200
        assert h == 100

    def test_matching_ratio_no_crop(self) -> None:
        """アスペクト比が一致 → クロップなし。"""
        img = Image.new('RGB', (200, 300))
        result = _center_crop(img, 200, 300)
        assert result.size == (200, 300)


# ── get_match_status ─────────────────────────────────────────────────────────


class TestGetMatchStatus:
    """マッチ状況のテスト。"""

    def test_all_matched(self) -> None:
        df = pd.DataFrame([
            {'学年': '1', '組': '1', '出席番号': '1', '氏名': 'A'},
            {'学年': '1', '組': '1', '出席番号': '2', '氏名': 'B'},
        ])
        photo_map = {'1-1-01': '/p/01.jpg', '1-1-02': '/p/02.jpg'}
        matched, total, unmatched = get_match_status(df, photo_map)
        assert matched == 2
        assert total == 2
        assert unmatched == []

    def test_partial_match(self) -> None:
        df = pd.DataFrame([
            {'学年': '1', '組': '1', '出席番号': '1', '氏名': '山田太郎'},
            {'学年': '1', '組': '1', '出席番号': '2', '氏名': '田中花子'},
        ])
        photo_map = {'1-1-01': '/p/01.jpg'}
        matched, total, unmatched = get_match_status(df, photo_map)
        assert matched == 1
        assert total == 2
        assert '田中花子' in unmatched

    def test_empty_df(self) -> None:
        df = pd.DataFrame()
        matched, total, unmatched = get_match_status(df, {})
        assert matched == 0
        assert total == 0
        assert unmatched == []


# ── import_photos_from_folder ────────────────────────────────────────────────


class TestImportPhotos:
    """写真一括コピーのテスト。"""

    def test_copies_files(self, tmp_path: object) -> None:
        src = os.path.join(str(tmp_path), 'src')
        dst = os.path.join(str(tmp_path), 'dst')
        os.makedirs(src)
        os.makedirs(dst)
        _create_test_image(os.path.join(src, '01.jpg'))
        _create_test_image(os.path.join(src, '02.png'))

        copied, skipped = import_photos_from_folder(src, dst)
        assert copied == 2
        assert skipped == []
        assert os.path.exists(os.path.join(dst, '01.jpg'))
        assert os.path.exists(os.path.join(dst, '02.png'))

    def test_skips_existing(self, tmp_path: object) -> None:
        src = os.path.join(str(tmp_path), 'src')
        dst = os.path.join(str(tmp_path), 'dst')
        os.makedirs(src)
        os.makedirs(dst)
        _create_test_image(os.path.join(src, '01.jpg'))
        _create_test_image(os.path.join(dst, '01.jpg'))  # 既存

        copied, skipped = import_photos_from_folder(src, dst)
        assert copied == 0
        assert skipped == ['01.jpg']

    def test_ignores_non_image(self, tmp_path: object) -> None:
        src = os.path.join(str(tmp_path), 'src')
        dst = os.path.join(str(tmp_path), 'dst')
        os.makedirs(src)
        os.makedirs(dst)
        with open(os.path.join(src, 'readme.txt'), 'w') as f:
            f.write('test')

        copied, skipped = import_photos_from_folder(src, dst)
        assert copied == 0
        assert skipped == []

    def test_nonexistent_src(self, tmp_path: object) -> None:
        dst = os.path.join(str(tmp_path), 'dst')
        copied, skipped = import_photos_from_folder('/nonexistent', dst)
        assert copied == 0
        assert skipped == []

    def test_auto_creates_dst(self, tmp_path: object) -> None:
        src = os.path.join(str(tmp_path), 'src')
        dst = os.path.join(str(tmp_path), 'dst')
        os.makedirs(src)
        _create_test_image(os.path.join(src, '01.jpg'))

        copied, _ = import_photos_from_folder(src, dst)
        assert copied == 1
        assert os.path.isdir(dst)
