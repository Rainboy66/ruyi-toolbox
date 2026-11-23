"""Static JPG/PNG/WebP conversion with explicit limits and safe output."""
import tempfile
import warnings
from io import BytesIO
from dataclasses import dataclass
from pathlib import Path

from PIL import Image, ImageOps, UnidentifiedImageError
from file_output import publish_file

FORMATS = {'JPEG': '.jpg', 'PNG': '.png', 'WEBP': '.webp'}
MAX_PIXELS = 50_000_000


@dataclass(frozen=True)
class ImageResult:
    output_path: Path
    original_bytes: int
    output_bytes: int
    output_size: tuple[int, int]
    note: str


def load_image(path: Path) -> Image.Image:
    try:
        with warnings.catch_warnings():
            warnings.simplefilter('error', Image.DecompressionBombWarning)
            with Image.open(path) as opened:
                if opened.format not in FORMATS:
                    raise ValueError('当前仅支持 JPG、PNG、WebP 静态图片。')
                if getattr(opened, 'n_frames', 1) != 1:
                    raise ValueError('这是动画图片；当前版本不支持，避免丢失动画。')
                if opened.width * opened.height > MAX_PIXELS:
                    raise ValueError('图片超过 5000 万像素，请先缩小尺寸。')
                opened.load()
                return ImageOps.exif_transpose(opened)
    except (UnidentifiedImageError, Image.DecompressionBombError, Image.DecompressionBombWarning) as error:
        raise ValueError('图片无效、格式不受支持或像素过大。') from error


def convert_image(source: Path, output_dir: Path, output_format: str = 'ORIGINAL',
                  quality: int = 82, max_side: int | None = None, *,
                  lossless: bool = False, target_kb: int | None = None,
                  only_smaller: bool = True) -> ImageResult:
    source = Path(source).resolve(strict=True)
    output_dir = Path(output_dir).resolve(strict=True)
    try:
        with Image.open(source) as original:
            original_format = original.format
    except (UnidentifiedImageError, Image.DecompressionBombError) as error:
        raise ValueError('图片无效、格式不受支持或像素过大。') from error
    if output_format == 'ORIGINAL':
        output_format = original_format
    if lossless and output_format == 'JPEG':
        output_format = 'PNG'
    if output_format not in FORMATS:
        raise ValueError('请选择 JPG、PNG 或 WebP。')
    if not 1 <= quality <= 100:
        raise ValueError('画质应在 1 到 100 之间。')
    if target_kb is not None and not 1 <= target_kb <= 102400:
        raise ValueError('目标大小应在 1 到 102400 KB 之间。')
    if max_side is not None and not 100 <= max_side <= 12000:
        raise ValueError('最长边应在 100 到 12000 像素之间。')
    original_bytes = source.stat().st_size
    image = load_image(source)
    try:
        original_size = image.size
        icc = image.info.get('icc_profile')
        alpha = 'A' in image.getbands() or 'transparency' in image.info
        image = image.convert('RGBA' if alpha else 'RGB')
        if max_side is not None:
            image.thumbnail((max_side, max_side), Image.Resampling.LANCZOS)
        notes = []
        if lossless:
            notes.append('无损编码' + ('；JPEG 改存 PNG' if original_format == 'JPEG' and output_format == 'PNG' else ''))
        if output_format == 'JPEG' and alpha:
            background = Image.new('RGB', image.size, 'white')
            background.paste(image, mask=image.getchannel('A'))
            image = background
            notes.append('透明区域已填白')
        # Pillow can implicitly copy EXIF from image.info for some encoders.
        image.info.clear()
        options = {'icc_profile': icc} if icc else {}
        if output_format == 'JPEG':
            options.update(quality=quality, optimize=True, progressive=True)
        elif output_format == 'WEBP':
            options.update(quality=quality, method=4, lossless=lossless, exact=True)
        else:
            options.update(optimize=True, compress_level=9)
        def encode(level):
            settings = dict(options)
            if output_format != 'PNG':
                settings['quality'] = level
            with BytesIO() as buffer:
                image.save(buffer, format=output_format, **settings)
                return buffer.getvalue()

        data = encode(quality)
        # Keep the requested dimensions. A size limit never silently resizes.
        if target_kb is not None:
            limit = target_kb * 1024
            if len(data) > limit:
                if output_format == 'PNG' or lossless:
                    raise ValueError('无损输出无法达到目标大小；请减小像素，或选择 JPG/WebP 有损编码。')
                minimum = encode(1)
                if len(minimum) > limit:
                    raise ValueError('保持当前像素无法达到目标大小；请减小最长边或提高目标 KB。')
                data, chosen, low, high = minimum, 1, 2, quality - 1
                while low <= high:
                    middle = (low + high) // 2
                    candidate = encode(middle)
                    if len(candidate) <= limit:
                        data, chosen, low = candidate, middle, middle + 1
                    else:
                        high = middle - 1
                notes.append(f'目标已达成，编码画质 {chosen}')
            else:
                notes.append('目标大小已达成')
        with tempfile.TemporaryDirectory(prefix='ruyi-image-') as temp:
            completed = Path(temp) / ('image' + FORMATS[output_format])
            if (only_smaller and output_format == original_format and image.size == original_size
                    and len(data) >= original_bytes):
                # Explicitly report retained metadata instead of claiming EXIF removal.
                completed = source
                notes = ['未变小，保留原文件副本（含原元数据）']
            else:
                completed.write_bytes(data)
                if len(data) > original_bytes:
                    notes.append('格式或尺寸已变化，结果体积增加')
            target = publish_file(completed, output_dir, source.stem, FORMATS[output_format])
        return ImageResult(target, original_bytes, target.stat().st_size, image.size, '；'.join(notes))
    finally:
        image.close()
