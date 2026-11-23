"""Local portrait matting and background replacement using Apache-2.0 MODNet.

Model: ZHKKKe/MODNet, ONNX conversion distributed by Xenova/modnet.
Its output is an alpha matte.  Keep its absolute alpha values and remove the
old background from translucent pixels before compositing a new backdrop.
"""
import hashlib
import os
from pathlib import Path
import sys
import tempfile
import threading
import time
from urllib.request import urlopen

import numpy as np
from PIL import Image, ImageChops, ImageFilter, ImageOps
from file_output import publish_file
from id_photo_backgrounds import COLORS, make_background, resolve_background
from office_engine import ConversionCancelled, cleanup_workspace

MAX_PIXELS = 60_000_000
MODEL_URL = ('https://huggingface.co/Xenova/modnet/resolve/'
             'e231692d2977d62c7be22830f92d9f8a2f0737f7/onnx/model.onnx')
MODEL_SHA256 = '07c308cf0fc7e6e8b2065a12ed7fc07e1de8febb7dc7839d7b7f15dd66584df9'
MODEL_BYTES = 25_888_640
_SESSION = None
_SESSION_LOCK = threading.Lock()
EDGE_MODES = ('自动精细', '去除余边', '自然发丝')


def _hash_file(path):
    digest = hashlib.sha256()
    with Path(path).open('rb') as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b''):
            digest.update(block)
    return digest.hexdigest()


def _checked_model(path):
    return (path.is_file() and path.stat().st_size == MODEL_BYTES
            and _hash_file(path) == MODEL_SHA256)


def _model_path(cancel=None):
    if getattr(sys, 'frozen', False):
        bundled = Path(sys._MEIPASS) / 'models' / 'modnet.onnx'
        if _checked_model(bundled):
            return bundled
        raise ValueError('内置人像模型校验失败，请重新解压或安装如意完整版。')
    override = os.environ.get('RUYI_MODEL_DIR')
    local = os.environ.get('LOCALAPPDATA')
    model_root = Path(override) if override else Path(local or Path.home()) / 'Ruyi' / 'models'
    model_root.mkdir(parents=True, exist_ok=True)
    model_path = model_root / 'modnet.onnx'
    if _checked_model(model_path):
        return model_path
    partial = model_path.with_suffix('.onnx.part')
    try:
        with urlopen(MODEL_URL, timeout=30) as response, partial.open('wb') as output:
            while True:
                if cancel and cancel.is_set():
                    raise ConversionCancelled('已停止模型下载。')
                chunk = response.read(1024 * 1024)
                if not chunk:
                    break
                output.write(chunk)
        if not _checked_model(partial):
            raise ValueError('下载的人像模型校验未通过。')
        partial.replace(model_path)
        return model_path
    except Exception as error:
        partial.unlink(missing_ok=True)
        if isinstance(error, ConversionCancelled):
            raise
        raise ValueError('人像模型下载或校验失败（约 26 MB），请检查网络后重试。原始照片不会上传。') from error


def _session(cancel=None):
    global _SESSION
    if _SESSION is not None:
        return _SESSION
    try:
        import onnxruntime as ort
    except ImportError as error:
        raise ValueError('证件照 AI 组件未安装，请安装如意完整版本。') from error
    with _SESSION_LOCK:
        if _SESSION is not None:
            return _SESSION
        model_path = _model_path(cancel)
        try:
            options = ort.SessionOptions()
            options.intra_op_num_threads = min(4, os.cpu_count() or 1)
            _SESSION = ort.InferenceSession(str(model_path), sess_options=options,
                                           providers=['CPUExecutionProvider'])
            return _SESSION
        except Exception as error:
            raise ValueError('人像模型无法启动，请重新安装如意完整版。') from error


def _inference_size(size, short_edge=768):
    width, height = size
    scale = min(short_edge / min(width, height), 1536 / max(width, height))
    return (max(32, int(width * scale) // 32 * 32),
            max(32, int(height * scale) // 32 * 32))


def _human_mask(image, session):
    resized = image.convert('RGB').resize(_inference_size(image.size), Image.Resampling.LANCZOS)
    values = np.asarray(resized, dtype=np.float32) / 127.5 - 1.0
    resized.close()
    tensor = np.ascontiguousarray(values.transpose((2, 0, 1))[None, ...], dtype=np.float32)
    output = np.squeeze(session.run(None, {session.get_inputs()[0].name: tensor})[0])
    if output.ndim != 2 or not np.all(np.isfinite(output)):
        raise ValueError('人像模型返回无效结果，请更换照片后重试。')
    alpha = np.clip(output, 0.0, 1.0).astype(np.float32)
    if float(np.max(alpha)) < .2:
        raise ValueError('没有识别到清晰人像，请使用光线均匀的正面照片。')
    matte = Image.fromarray(alpha)
    try:
        return matte.resize(image.size, Image.Resampling.BILINEAR)
    finally:
        matte.close()


def _box_mean(array, radius):
    """Separable box average, requiring no extra image-processing dependency."""
    if radius < 1:
        return array.copy()
    span = radius * 2 + 1
    padded = np.pad(array, ((radius, radius), (0, 0)), mode='edge')
    integral = np.pad(np.cumsum(padded, axis=0, dtype=np.float64), ((1, 0), (0, 0)))
    vertical = (integral[span:] - integral[:-span]) / span
    del integral, padded
    padded = np.pad(vertical, ((0, 0), (radius, radius)), mode='edge')
    integral = np.pad(np.cumsum(padded, axis=1, dtype=np.float64), ((0, 0), (1, 0)))
    return ((integral[:, span:] - integral[:, :-span]) / span).astype(np.float32)


def _refine_alpha(rgb, alpha, edge_mode):
    """Guide the alpha boundary by original-resolution photographic edges."""
    guide = rgb[..., 0] * .299 + rgb[..., 1] * .587 + rgb[..., 2] * .114
    radius = max(1, round(min(alpha.shape) / 600))
    mean_i, mean_p = _box_mean(guide, radius), _box_mean(alpha, radius)
    covariance = _box_mean(guide * alpha, radius) - mean_i * mean_p
    variance = _box_mean(guide * guide, radius) - mean_i * mean_i
    slope = covariance / (variance + .001)
    offset = mean_p - slope * mean_i
    refined = _box_mean(slope, radius) * guide + _box_mean(offset, radius)
    uncertain = (alpha > .003) & (alpha < .997)
    # Fine strands have little spatial support. Preserve some of the model's
    # original absolute opacity instead of letting a local average erase them.
    guided = .65 * refined + .35 * alpha if edge_mode == '自动精细' else refined
    alpha = np.where(uncertain, guided, alpha)
    if edge_mode == '去除余边':
        alpha = (alpha - .012) / .988
    alpha = np.clip(alpha, 0.0, 1.0).astype(np.float32)
    alpha[alpha < .004] = 0.0
    alpha[alpha > .997] = 1.0
    return alpha


def _old_background(rgb, alpha):
    """Estimate the nearby old backdrop using confident background pixels."""
    weights = np.clip((.06 - alpha) / .06, 0.0, 1.0)
    border = np.zeros(alpha.shape, dtype=bool)
    border[:max(1, alpha.shape[0] // 20), :] = True
    border[-max(1, alpha.shape[0] // 20):, :] = True
    border[:, :max(1, alpha.shape[1] // 20)] = True
    border[:, -max(1, alpha.shape[1] // 20):] = True
    samples = rgb[border & (weights > .9)]
    if len(samples) < 20:
        samples = rgb[weights > .9]
    if len(samples) < 20:
        return None, None
    global_color = np.median(samples, axis=0)
    homogeneous = np.quantile(np.max(np.abs(samples - global_color), axis=1), .8) < .10
    radius = max(8, round(min(alpha.shape) / 40))
    denominator = _box_mean(weights, radius)
    background = np.empty_like(rgb)
    for channel in range(3):
        local = _box_mean(rgb[..., channel] * weights, radius) / np.maximum(denominator, .0001)
        background[..., channel] = np.where(denominator > .015, local, global_color[channel])
    confidence = np.ones(alpha.shape, dtype=np.float32) if homogeneous else np.clip(denominator / .1, 0, 1)
    return background, confidence


def _refine_uniform_boundary(rgb, alpha, background, confidence, edge_mode):
    """Correct mixed edge pixels that the model predicted as fully opaque.

    Only a narrow contour of a uniform backdrop is eligible. The pixel must
    lie close to the colour line connecting a reliable inner foreground and
    the old background, so a pale shirt is not mistaken for the white backdrop.
    """
    if float(confidence.min()) < .999:
        return alpha
    radius = max(2, min(4, round(min(alpha.shape) / 400)))
    with Image.fromarray((alpha > .5).astype(np.uint8)*255) as hard:
        with hard.filter(ImageFilter.MaxFilter(radius*2+1)) as dilated:
            outside = np.asarray(dilated) > 0
        with hard.filter(ImageFilter.MinFilter(radius*2+1)) as eroded:
            inside = np.asarray(eroded) > 0
    band = outside & ~inside & (alpha > .03)
    with Image.fromarray((alpha > .985).astype(np.uint8)*255) as opaque:
        with opaque.filter(ImageFilter.MinFilter(radius*2+1)) as eroded:
            core = np.asarray(eroded, dtype=np.float32)/255.0
    weight = _box_mean(core, radius*3)
    foreground = np.empty_like(rgb)
    for channel in range(3):
        foreground[..., channel] = _box_mean(rgb[..., channel]*core, radius*3) / np.maximum(weight, .0001)
    direction = foreground-background
    separation = np.sum(direction*direction, axis=2)
    projected = np.clip(np.sum((rgb-background)*direction, axis=2)
                        / np.maximum(separation, .0001), 0, 1)
    reconstruction = background+projected[..., None]*direction
    residual = np.max(np.abs(rgb-reconstruction), axis=2)
    eligible = band & (weight > .03) & (separation > .10) & (residual < .055)
    corrected = np.minimum(alpha, projected+.012)
    strength = .95 if edge_mode == '去除余边' else .5
    return np.where(eligible, alpha+strength*(corrected-alpha), alpha).astype(np.float32)


def _recover_foreground(rgb, alpha, background):
    """Unmix old backdrop without magnifying noise in translucent flyaway hair.

    At low alpha, solving (C - (1-a)B) / a directly amplifies tiny JPEG or matte
    errors. A nearby reliable foreground supplies a small regularisation term.
    It vanishes on opaque pixels and does not apply any whole-portrait colour
    matching or lightening.
    """
    a = alpha[..., None]
    numerator = rgb - (1-a) * background
    direct = np.clip(numerator / np.maximum(a, .02), 0, 1)
    weights = np.where(alpha >= .35, alpha * alpha, 0).astype(np.float32)
    radius = max(3, min(10, round(min(alpha.shape) / 200)))
    support = _box_mean(weights, radius)
    prior = np.empty_like(rgb)
    for channel in range(3):
        local = _box_mean(direct[..., channel] * weights, radius)
        prior[..., channel] = np.where(support > .01,
                                      local / np.maximum(support, .0001),
                                      direct[..., channel])
    # Only the uncertain translucent fringe needs regularisation. Retain the
    # existing exact solution for stronger hair, clothing and skin opacity.
    regulariser = np.square(.06 * (1-alpha)) * (alpha < .25) * (support > .01)
    recovered = (a * numerator + regulariser[..., None] * prior) / np.maximum(
        a*a + regulariser[..., None], .0004)
    return np.clip(recovered, 0, 1)


def _refine_auto_boundary(rgb, alpha, background, confidence):
    """Refine each contour pixel only when its photographed colour supports it.

    The network's opacity is not an erosion mask. A solid contour with a nearby
    reliable foreground can be tested against C=aF+(1-a)B. A thin strand has
    less support and receives a bounded correction; an opaque pale shirt or
    skin that is not a reliable mixture is left alone. Background confidence,
    colour separation, fit and local support all contribute continuously.
    """
    radius = max(2, min(5, round(min(alpha.shape) / 350)))
    with Image.fromarray((alpha > .5).astype(np.uint8) * 255) as hard:
        with hard.filter(ImageFilter.MaxFilter(radius * 2 + 1)) as dilated:
            outside = np.asarray(dilated) > 0
        with hard.filter(ImageFilter.MinFilter(radius * 2 + 1)) as eroded:
            inside = np.asarray(eroded) > 0
    band = outside & ~inside & (alpha > .003)
    # Eroding only the *sampling mask* excludes mixed RGB from the foreground
    # estimate. It never erodes the exported silhouette itself.
    with Image.fromarray((alpha > .985).astype(np.uint8) * 255) as opaque:
        with opaque.filter(ImageFilter.MinFilter(radius * 2 + 1)) as eroded:
            core = np.asarray(eroded, dtype=np.float32) / 255.0
    support = _box_mean(core, radius * 3)
    foreground = np.empty_like(rgb)
    foreground_variance = np.zeros_like(alpha)
    for channel in range(3):
        foreground[..., channel] = _box_mean(rgb[..., channel] * core, radius * 3) / np.maximum(support, .0001)
        second_moment = _box_mean(rgb[..., channel] ** 2 * core, radius * 3) / np.maximum(support, .0001)
        foreground_variance += np.maximum(second_moment - foreground[..., channel] ** 2, 0)
    direction = foreground - background
    separation = np.sum(direction * direction, axis=2)
    projection = np.clip(np.sum((rgb - background) * direction, axis=2) /
                         np.maximum(separation, .0001), 0, 1)
    residual = np.max(np.abs(rgb - (background + projection[..., None] * direction)), axis=2)
    reliability = (confidence * np.clip((support - .02) / .16, 0, 1)
                   * np.clip((separation - .06) / .20, 0, 1)
                   * np.clip((.06 - residual) / .04, 0, 1)
                   # Adjacent pale clothing and a dark lapel are not one F.
                   # Do not solve their average as a translucent dark fringe.
                   * np.clip((.025 - foreground_variance) / .02, 0, 1))
    occupancy = _box_mean((alpha > .5).astype(np.float32), radius * 2)
    thin = (occupancy < .48) & (alpha < .95)
    # Thin hair is not a uniformly contracted contour. In ambiguous mixtures,
    # bound its opacity reduction and protect the positive low-alpha signal.
    candidate = np.where(thin, np.maximum(projection, alpha * .65), projection)
    candidate = np.minimum(candidate, alpha + .12)
    strength = np.where(thin, .65, .98) * reliability * band
    refined = alpha + strength * (candidate - alpha)
    # A weak model haze whose camera RGB is indistinguishable from a reliable
    # old backdrop is not a strand. Real dark low-alpha hair has colour contrast.
    backdrop_distance = np.max(np.abs(rgb - background), axis=2)
    haze = (alpha < .08) & (backdrop_distance < .012) & (confidence > .95)
    refined = np.where(haze, refined * np.clip(backdrop_distance / .012, 0, 1), refined)
    return np.clip(refined, 0, 1).astype(np.float32)


def _composite(image, matte, color, edge_mode):
    # Keep RGB tuples compatible with engine callers from earlier versions.
    if isinstance(color, tuple):
        if len(color) != 3 or any(type(value) is not int or not 0 <= value <= 255 for value in color):
            raise ValueError('背景 RGB 颜色应为三个 0–255 的整数。')
        background_spec = '#%02X%02X%02X' % color
    else:
        background_spec = resolve_background(color)
    if edge_mode not in EDGE_MODES:
        raise ValueError('边缘处理应为“自动精细”、“去除余边”或“自然发丝”。')
    # Bound edge processing at 2048 px; exported dimensions stay unchanged.
    work_size = image.size
    if max(work_size) > 2048:
        scale = 2048 / max(work_size)
        work_size = tuple(max(1, round(value * scale)) for value in work_size)
    work_image = image.resize(work_size, Image.Resampling.LANCZOS)
    work_matte = matte.resize(work_size, Image.Resampling.BILINEAR)
    rgb = np.asarray(work_image, dtype=np.float32) / 255.0
    alpha = np.clip(np.asarray(work_matte, dtype=np.float32), 0, 1)
    work_image.close()
    work_matte.close()
    alpha = _refine_alpha(rgb, alpha, edge_mode)
    background, confidence = _old_background(rgb, alpha)
    if background is not None:
        if edge_mode == '自动精细':
            alpha = _refine_auto_boundary(rgb, alpha, background, confidence)
        else:
            alpha = _refine_uniform_boundary(rgb, alpha, background, confidence, edge_mode)
        # C = a F + (1-a) B: recover F before replacing B.
        recovered = _recover_foreground(rgb, alpha, background)
        strength = .7 if edge_mode == '自然发丝' else 1.0
        mix = confidence * strength * (alpha > .004) * (alpha < .997)
        rgb += mix[..., None] * (recovered-rgb)
    # Resize premultiplied foreground and alpha with the same filter. Resizing
    # straight RGB independently can smear discarded old-background colours
    # into the translucent edge when a high-resolution photo is restored.
    foreground = Image.fromarray(np.rint(np.clip(rgb * alpha[..., None], 0, 1) * 255).astype(np.uint8))
    alpha_image = Image.fromarray(np.rint(alpha * 255).astype(np.uint8))
    if work_size != image.size:
        small_foreground, small_alpha = foreground, alpha_image
        foreground = foreground.resize(image.size, Image.Resampling.BILINEAR)
        alpha_image = alpha_image.resize(image.size, Image.Resampling.BILINEAR)
        small_foreground.close()
        small_alpha.close()
    backdrop = make_background(image.size, background_spec)
    try:
        # PIL works in bounded byte images here, avoiding several 60 MP float
        # arrays when exporting a large original-resolution photograph.
        backdrop.paste((0, 0, 0), mask=alpha_image)
        result = ImageChops.add(foreground, backdrop)
        # Premultiplication is rounded to bytes above; opaque camera pixels
        # remain exact even on photos larger than the edge-processing limit.
        with alpha_image.point(lambda value: 255 if value >= 254 else 0) as opaque:
            result.paste(image, mask=opaque)
        return result
    finally:
        backdrop.close()
        foreground.close()
        alpha_image.close()


def replace_background(source: Path, output_dir: Path, color: str,
                       cancel: threading.Event | None = None, *,
                       edge_mode: str = '自动精细') -> tuple[Path, int, float]:
    source = Path(source).resolve(strict=True)
    output_dir = Path(output_dir).resolve(strict=True)
    background_spec = resolve_background(color)
    if edge_mode not in EDGE_MODES:
        raise ValueError('边缘处理应为“自动精细”、“去除余边”或“自然发丝”。')
    started = time.monotonic()
    if cancel and cancel.is_set():
        raise ConversionCancelled('已停止。')
    try:
        with Image.open(source) as opened:
            if opened.width * opened.height > MAX_PIXELS:
                raise ValueError('照片像素过大，当前上限为 6000 万像素。')
            original = ImageOps.exif_transpose(opened).convert('RGB')
    except (OSError, Image.DecompressionBombError) as error:
        raise ValueError('无法读取照片；请确认文件未损坏且未超过像素上限。') from error
    try:
        if cancel and cancel.is_set():
            raise ConversionCancelled('已停止。')
        matte = _human_mask(original, _session(cancel))
        try:
            result = _composite(original, matte, background_spec, edge_mode)
        finally:
            matte.close()
        try:
            if cancel and cancel.is_set():
                raise ConversionCancelled('已停止。')
            workspace = Path(tempfile.mkdtemp(prefix='ruyi-idphoto-'))
            try:
                generated = workspace / 'id-photo.jpg'
                result.save(generated, format='JPEG', quality=98, subsampling=0, optimize=True)
                background_name = background_spec.name
                if background_spec.code == 'CUSTOM':
                    background_name += '_' + background_spec.start[1:]
                target = publish_file(generated, output_dir, source.stem + '_证件照_' + background_name, '.jpg')
                return target, target.stat().st_size, time.monotonic() - started
            finally:
                cleanup_workspace(workspace)
        finally:
            result.close()
    finally:
        original.close()
