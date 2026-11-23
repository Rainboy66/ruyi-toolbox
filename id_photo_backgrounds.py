"""Ordered photo background palette supplied by the user; no universal ID rules."""
from dataclasses import dataclass
import re


@dataclass(frozen=True)
class Background:
    code: str
    name: str
    start: str
    end: str | None = None
    usage: str = ''

    @property
    def gradient(self):
        return self.end is not None


BACKGROUNDS = (
    Background('BG-CN-W01', '标准白', '#FFFFFF', usage='身份证、护照、驾照、签证参考'),
    Background('BG-CN-B01', '标准蓝', '#438EDB', usage='简历、工作证、报名表参考'),
    Background('BG-CN-B02', '亮蓝', '#00A0E9', usage='学信网、部分考试采集参考'),
    Background('BG-CN-B03', '深蓝', '#005BAC', usage='商务证件照、职业资格证参考'),
    Background('BG-CN-B04', '藏青', '#1F3864', usage='商务形象照参考'),
    Background('BG-CN-B05', '浅蓝', '#A8CFEF', usage='儿童照、留学材料参考'),
    Background('BG-CN-R01', '正红', '#FF0000', usage='屏幕显示、结婚证电子版参考'),
    Background('BG-CN-R02', '证件红', '#D9001B', usage='打印用红参考'),
    Background('BG-CN-G01', '浅灰', '#D9D9D9', usage='商务照、英美签证参考'),
    Background('BG-CN-G02', '中灰', '#C8C8C8', usage='简历、艺术照参考'),
    Background('BG-CN-G03', '深灰', '#9A9A9A', usage='商务、深色西装参考'),
    Background('BG-CN-V01', '经典蓝白渐变', '#A8CFEF', '#2E6FB7', '上浅下深，照相馆风格'),
)
BACKGROUND_NAMES = tuple(item.name for item in BACKGROUNDS)
# Preserve old callers. New UI uses the ordered palette above.
COLORS = {'红底': '#FF0000', '蓝底': '#438EDB', '白底': '#FFFFFF'}
ALIASES = {'红底': '正红', '蓝底': '标准蓝', '白底': '标准白'}


def resolve_background(value):
    if isinstance(value, Background):
        # Validate even internal specifications before allocating image memory.
        if not re.fullmatch(r'#[0-9A-Fa-f]{6}', value.start) or (
                value.end is not None and not re.fullmatch(r'#[0-9A-Fa-f]{6}', value.end)):
            raise ValueError('背景颜色格式应为 #RRGGBB。')
        return value
    if not isinstance(value, str):
        raise ValueError('请选择背景预设或 #RRGGBB 格式的自定义颜色。')
    value = value.strip()
    value = ALIASES.get(value, value)
    for preset in BACKGROUNDS:
        if value in (preset.code, preset.name):
            return preset
    if re.fullmatch(r'#[0-9A-Fa-f]{6}', value):
        return Background('CUSTOM', '自定义', value.upper())
    raise ValueError('背景颜色格式无效，请选择预设或输入 #RRGGBB。')


def make_background(size, value):
    """Render exact solid colors or a top-to-bottom sRGB gradient."""
    from PIL import Image
    background = resolve_background(value)
    if not background.gradient:
        return Image.new('RGB', size, background.start)
    import numpy as np
    width, height = size
    start = np.array([int(background.start[i:i+2], 16) for i in (1, 3, 5)], dtype=np.float32)
    end = np.array([int(background.end[i:i+2], 16) for i in (1, 3, 5)], dtype=np.float32)
    weights = np.linspace(0, 1, height, dtype=np.float32)[:, None]
    colors = np.rint(start[None, :] * (1-weights) + end[None, :] * weights).astype(np.uint8)
    strip = Image.fromarray(colors[:, None, :])
    try:
        return strip.resize((width, height), Image.Resampling.NEAREST)
    finally:
        strip.close()
