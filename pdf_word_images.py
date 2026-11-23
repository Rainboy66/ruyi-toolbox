"""Encoding only: never resize, change crop geometry or Word drawing extents."""
from dataclasses import dataclass
from io import BytesIO

# Emergency rollback, including original-stream reuse, to the former RGB PNG.
OPTIMIZE_WORD_IMAGES = True
JPEG_QUALITY = 88


@dataclass(frozen=True)
class SourceImage:
    data: bytes
    full_bounds: bool = False
    untransformed: bool = False
    pixels_match: bool = False


def choose_image_encoding(image, original=None):
    """Return (extension, Pillow parameters); mode/source are encoder controls.

    Alpha-bearing images stay PNG even when every alpha sample is opaque.
    Classification examines all pixels, without resampling or quantising output.
    Near-binary coloured marks stay RGB rather than silently losing their colour.
    """
    if not OPTIMIZE_WORD_IMAGES:
        return 'png', {'format': 'PNG'}
    png = {'format': 'PNG', 'optimize': True, 'compress_level': 9}
    if 'A' in image.getbands() or 'transparency' in image.info:
        return 'png', png
    if original and original.full_bounds and original.untransformed and original.pixels_match:
        return 'jpg', {'source': original.data}
    if min(image.size) < 200 or (original and len(original.data) < 30 * 1024):
        return 'png', png
    import numpy as np
    rgb = np.asarray(image.convert('RGB'))
    extremes = (rgb.max(axis=2) <= 8) | (rgb.min(axis=2) >= 247)
    if np.count_nonzero(extremes) >= extremes.size * .98:
        # Retain antialiasing in 8-bit L; 1-bit only for exactly binary pixels.
        if np.array_equal(rgb[:, :, 0], rgb[:, :, 1]) and np.array_equal(rgb[:, :, 1], rgb[:, :, 2]):
            exact = np.all((rgb[:, :, 0] == 0) | (rgb[:, :, 0] == 255))
            return 'png', dict(png, mode='1' if exact else 'L')
        return 'png', png
    colors = image.convert('RGB').getcolors(maxcolors=4095)
    if colors is not None:
        # Fixed 5-bit RGB buckets; no palette conversion of actual output.
        buckets = rgb.astype(np.uint16) >> 3
        keys = (buckets[:, :, 0] << 10) | (buckets[:, :, 1] << 5) | buckets[:, :, 2]
        counts = np.bincount(keys.ravel(), minlength=32768)
        if np.partition(counts, -32)[-32:].sum() >= keys.size * .95:
            return 'png', png
    return 'jpg', {'format': 'JPEG', 'quality': JPEG_QUALITY,
                   'subsampling': 2, 'optimize': True}


def encode_word_image(image, original=None):
    if OPTIMIZE_WORD_IMAGES and original is None:
        # The previous PNG bytes define the 30 KB rule for raster crops which
        # have no standalone source stream (formula, chart or composite page).
        baseline = BytesIO()
        image.save(baseline, format='PNG')
        original = SourceImage(baseline.getvalue())
    extension, options = choose_image_encoding(image, original)
    options = dict(options)
    raw = options.pop('source', None)
    if raw is not None:
        return BytesIO(raw)
    mode = options.pop('mode', None)
    converted = image.convert(mode) if mode else None
    # PIL cannot save e.g. CMYK or palette images as JPEG without conversion.
    if extension == 'jpg' and image.mode not in ('L', 'RGB'):
        converted = image.convert('RGB')
    result = BytesIO()
    try:
        (converted if converted is not None else image).save(result, **options)
    finally:
        if converted is not None:
            converted.close()
    result.seek(0)
    return result


def matching_source_jpeg(pdf_page, box, cropped, page_height):
    """Conservatively reuse top-level JPEGs; other objects use the rendered crop.

    Exact pixel comparison also protects clipping, transparency, colour decode,
    interpolation and overlaid content, not exposed fully by PDFium extraction.
    A source whose pixel size differs from the existing crop is never reused.
    """
    from PIL import Image
    from pypdfium2 import raw
    try:
        if pdf_page.get_rotation():
            return None
        for obj in pdf_page.get_objects(max_depth=0, filter=[raw.FPDF_PAGEOBJ_IMAGE]):
            left, bottom, right, top = obj.get_bounds()
            bounds = (left, page_height - top, right, page_height - bottom)
            if any(abs(a-b) > .5 for a, b in zip(bounds, box)):
                continue
            matrix = obj.get_matrix()
            if abs(matrix.b) > .001 or abs(matrix.c) > .001 or matrix.a <= 0 or matrix.d <= 0:
                continue
            if obj.get_px_size() != cropped.size or obj.get_filters() != ['DCTDecode']:
                continue
            data = bytes(obj.get_data())
            with Image.open(BytesIO(data)) as source:
                if source.format != 'JPEG' or source.size != cropped.size or source.mode not in ('L', 'RGB'):
                    continue
                decoded = source.convert('RGB')
                try:
                    if decoded.tobytes() == cropped.convert('RGB').tobytes():
                        return SourceImage(data, True, True, True)
                finally:
                    decoded.close()
    except (AttributeError, RuntimeError, OSError, ValueError):
        # Optional optimisation must not fail a formerly convertible document.
        return None
    return None
