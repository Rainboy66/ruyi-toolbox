"""Rebuild PDF pages as independently editable PowerPoint objects.

PDF has coordinates rather than document semantics.  Keep those coordinates
instead of flattening paragraphs, and retain image-only pages without OCR.
No font or image is downloaded and the source PDF is opened read-only.
"""
from dataclasses import dataclass
from io import BytesIO
import math
import re
from statistics import median


@dataclass
class PptxLayoutStats:
    text_boxes: int = 0
    images: int = 0
    shapes: int = 0
    scanned_pages: int = 0
    raster_fragments: int = 0
    tables: int = 0
    cells: int = 0


def _emu(value):
    return round(float(value) * 12700)


def _rgb(value):
    """Convert the PDF's gray, RGB or CMYK color to an RGB tuple."""
    if isinstance(value, (int, float)):
        value = (value, value, value)
    if not isinstance(value, (tuple, list)):
        return (0, 0, 0)
    if len(value) == 1:
        value = value * 3
    elif len(value) == 4:
        c, m, y, k = (float(v) for v in value)
        value = (1 - min(1, c + k), 1 - min(1, m + k), 1 - min(1, y + k))
    if len(value) != 3:
        return (0, 0, 0)
    return tuple(max(0, min(255, round(float(v) * 255))) for v in value)


def _font_name(value):
    value = re.sub(r'^[A-Z]{6}\+', '', value or 'Arial')
    normalized = re.sub(r'[\s,_-]', '', value).lower()
    families = {
        'simsun': '宋体', 'simsunregular': '宋体', 'simhei': '黑体',
        'kaiti': '楷体', 'kaitigb2312': '楷体', 'fangsong': '仿宋',
        'fangsonggb2312': '仿宋', 'microsoftyahei': '微软雅黑',
        'microsoftyaheibold': '微软雅黑', 'microsoftyaheilight': '微软雅黑 Light',
        'arialmt': 'Arial', 'arialboldmt': 'Arial', 'arialitalicmt': 'Arial',
        'arialbolditalicmt': 'Arial', 'helvetica': 'Arial',
        'helveticabold': 'Arial', 'helveticaoblique': 'Arial',
        'timesroman': 'Times New Roman', 'timesbold': 'Times New Roman',
        'timesnewromanpsmt': 'Times New Roman',
        'timesnewromanpsboldmt': 'Times New Roman',
        'timesnewromanpsitalicmt': 'Times New Roman',
        'timesnewromanpsbolditalicmt': 'Times New Roman',
        'courier': 'Courier New', 'couriernewpsmt': 'Courier New',
        'symbolmt': 'Symbol', 'zapfdingbats': 'Wingdings',
        'segoeui': 'Segoe UI', 'segoeuibold': 'Segoe UI',
        'segoeuiitalic': 'Segoe UI', 'segoeuibolditalic': 'Segoe UI',
        'segoeuisemibold': 'Segoe UI Semibold', 'segoeuilight': 'Segoe UI Light',
        'calibri': 'Calibri', 'calibribold': 'Calibri', 'calibriitalic': 'Calibri',
        'aptos': 'Aptos', 'aptosbold': 'Aptos', 'aptositalic': 'Aptos',
    }
    return families.get(normalized, value)


def _angle(char):
    matrix = char.get('matrix', (1, 0, 0, 1, 0, 0))
    return round((-math.degrees(math.atan2(matrix[1], matrix[0]))) % 360, 2)


def _baseline(char, height):
    matrix = char.get('matrix')
    return height - matrix[5] if matrix else char['bottom'] - char.get('size', 10) * .18


def _annotate_chars(page, pdf_page):
    """Keep synthetic fill-and-stroke bold, which font names do not express."""
    from pypdfium2 import raw
    bold_regions = []
    for obj in pdf_page.get_objects(filter=[raw.FPDF_PAGEOBJ_TEXT]):
        try:
            mode = raw.FPDFTextObj_GetTextRenderMode(obj)
            if mode in (1, 2, 5, 6) or obj.get_font().get_weight() >= 600:
                x0, y0, x1, y1 = obj.get_bounds()
                bold_regions.append((x0, page.height - y1, x1, page.height - y0))
        except (AttributeError, RuntimeError, ValueError):
            continue
    for char in page.chars:
        x, y = (char['x0'] + char['x1']) / 2, (char['top'] + char['bottom']) / 2
        char['_ppt_bold'] = any(a - .3 <= x <= c + .3 and b - .3 <= y <= d + .3
                                for a, b, c, d in bold_regions)


def _annotate_order(page):
    """Preserve painter order before pdfplumber groups objects by type."""
    from collections import defaultdict
    counts = defaultdict(int)
    layer = 0
    for index, obj in enumerate(page.iter_layout_objects(page.layout._objs)):
        kind = obj['object_type']
        if kind == 'anno':
            continue
        target = page.objects[kind][counts[kind]]
        target['_ppt_order'] = index
        if kind != 'char':
            layer += 1
        target['_ppt_layer'] = layer
        counts[kind] += 1


def _annotate_vector_alpha(page, pdf_page):
    """pdfplumber omits graphics-state alpha; retain it from PDFium paths."""
    import ctypes
    from pypdfium2 import raw
    paths = []
    for obj in pdf_page.get_objects(filter=[raw.FPDF_PAGEOBJ_PATH]):
        a, b, c, d = obj.get_bounds()
        paths.append((obj, (a, page.height-d, c, page.height-b)))
    used = set()
    for item in sorted([*page.rects, *page.lines, *page.curves], key=lambda v: v.get('_ppt_order', 0)):
        box = (item['x0'], item['top'], item['x1'], item['bottom'])
        candidates = [(sum(abs(a-b) for a, b in zip(box, bounds)), i, obj)
                      for i, (obj, bounds) in enumerate(paths) if i not in used]
        if not candidates:
            continue
        error, index, obj = min(candidates, key=lambda value: value[:2])
        if error > 4:
            continue
        used.add(index)
        for name, reader in (('_ppt_fill_alpha', raw.FPDFPageObj_GetFillColor),
                             ('_ppt_stroke_alpha', raw.FPDFPageObj_GetStrokeColor)):
            values = [ctypes.c_uint() for _ in range(4)]
            if reader(obj, *(ctypes.byref(v) for v in values)):
                item[name] = values[3].value


def _segments(page, chars=None):
    """Cluster by baseline, not top edge, so smaller dates keep their position."""
    lines = []
    rotated = []
    for char in page.chars if chars is None else chars:
        text = char.get('text', '')
        if not text or text == '\x00':
            continue
        if _angle(char) not in (0, 360):
            rotated.append([char])
            continue
        baseline = _baseline(char, page.height)
        line = next((line for line in lines
                     if abs(line[0] - baseline) <= .7), None)
        if line is None:
            lines.append((baseline, [char]))
        else:
            line[1].append(char)
    for _baseline_value, chars in sorted(lines, key=lambda item: item[0]):
        segment = []
        for char in sorted(chars, key=lambda item: item['x0']):
            if segment:
                previous = segment[-1]
                gap = char['x0'] - previous['x1']
                # Separate distant columns/dates. Retain genuine word spaces
                # within a box using exact character spacing below.
                if (gap > max(8, min(char.get('size', 10), previous.get('size', 10)) * .8) or
                        char.get('_ppt_layer') != previous.get('_ppt_layer')):
                    yield segment
                    segment = []
            segment.append(char)
        if segment:
            yield segment
    yield from rotated


def _element(name, **attrs):
    from pptx.oxml.xmlchemy import OxmlElement
    element = OxmlElement(name)
    for key, value in attrs.items():
        element.set(key, str(value))
    return element


def _write_text(slide, chars, transform):
    from pptx.dml.color import RGBColor
    from pptx.util import Pt
    ox, oy, scale = transform
    left, top = min(c['x0'] for c in chars), min(c['top'] for c in chars)
    right, bottom = max(c['x1'] for c in chars), max(c['bottom'] for c in chars)
    max_size = max(float(c.get('size', 10)) for c in chars)
    rotation = _angle(chars[0])
    # All regular fragments share a PDF baseline. Small inline dates receive
    # their own baseline shift instead of being top-aligned to the heading.
    baseline = _baseline(chars[0], chars[0].get('_page_height', 0))
    y = baseline - max_size * .93 if not rotation else top
    shape = slide.shapes.add_textbox(_emu(ox + left * scale), _emu(oy + y * scale),
                                    _emu(max(right - left + max_size * 2, 5) * scale),
                                    _emu(max(max_size * 1.65, bottom - top) * scale))
    shape.name = 'PDF文字 ' + ''.join(c['text'] for c in chars)[:24]
    frame = shape.text_frame
    frame.clear()
    frame.margin_left = frame.margin_right = 0
    frame.margin_top = frame.margin_bottom = 0
    frame.word_wrap = False
    body = frame._txBody.bodyPr
    body.set('anchor', 't')
    body.set('wrap', 'none')
    body.set('rtlCol', '0')
    body.set('horzOverflow', 'overflow')
    body.set('vertOverflow', 'overflow')
    for child in list(body):
        if child.tag.endswith(('spAutoFit', 'normAutofit')):
            body.remove(child)
    body.append(_element('a:noAutofit'))
    p = frame.paragraphs[0]
    p.space_before = p.space_after = Pt(0)
    p.line_spacing = Pt(max_size * scale)
    ppr = p._p.get_or_add_pPr()
    ppr.set('fontAlgn', 'base')
    ppr.set('marL', '0')
    ppr.set('indent', '0')
    ppr.set('eaLnBrk', '0')
    ppr.set('latinLnBrk', '0')
    for index, char in enumerate(chars):
        run = p.add_run()
        run.text = char['text'].replace('\x00', '')
        font = run.font
        font.name = _font_name(char.get('fontname'))
        font.size = Pt(max(.5, min(4000, float(char.get('size', 10)) * scale)))
        source_font = char.get('fontname', '').lower()
        font.bold = bool(char.get('_ppt_bold')) or any(k in source_font for k in ('bold', 'black', 'demi'))
        font.italic = any(k in source_font for k in ('italic', 'oblique'))
        font.color.rgb = RGBColor(*_rgb(char.get('non_stroking_color', 0)))
        rpr = run._r.get_or_add_rPr()
        rpr.set('kern', '0')
        # DrawingML's tracking is hundredths of a point and can represent
        # the source's negative tracking without inserting visible spaces.
        if index + 1 < len(chars):
            gap = chars[index + 1]['x0'] - char['x1']
            matrix = char.get('matrix', (1, 0, 0, 1, 0, 0))
            horizontal = math.hypot(matrix[0], matrix[1])
            vertical = math.hypot(matrix[2], matrix[3])
            # Some PDFs stretch text horizontally (the CV uses 104%).
            # Compensate the advances instead of allowing long lines to
            # become progressively shorter than the source.
            stretch = ((char['x1'] - char['x0']) * (1 - vertical / horizontal)
                       if horizontal > 1e-8 else 0)
            rpr.set('spc', str(max(-400000, min(400000, round((gap + stretch) * scale * 100)))))
        if char.get('size', max_size) != max_size:
            # All runs share baseline, as in the PDF.
            rpr.set('baseline', '0')
        for kind in ('a:ea', 'a:cs'):
            node = rpr.find('{http://schemas.openxmlformats.org/drawingml/2006/main}' + kind[2:])
            if node is None:
                node = _element(kind)
                rpr.append(node)
            node.set('typeface', font.name)
    end = p._p.get_or_add_endParaRPr()
    end.set('sz', str(round(max_size * scale * 100)))
    if rotation:
        # Rotated source glyphs are independent so their original placement
        # remains stable and they can still be edited or moved separately.
        if rotation in (90, 270):
            shape.width, shape.height = _emu(max(bottom - top, max_size) * scale), _emu(max(right - left, max_size) * scale)
            shape.left = _emu(ox + (left + (right - left - shape.width / 12700 / scale) / 2) * scale)
            shape.top = _emu(oy + (top + (bottom - top - shape.height / 12700 / scale) / 2) * scale)
        shape.rotation = rotation
    return shape


def _paint_shape(shape, obj, scale):
    from pptx.dml.color import RGBColor
    from pptx.util import Pt
    if obj.get('fill'):
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(*_rgb(obj.get('non_stroking_color', 0)))
        if obj.get('_ppt_fill_alpha', 255) < 255:
            shape.fill.fore_color._xFill.get_or_change_to_srgbClr().append(_element('a:alpha', val=round(obj['_ppt_fill_alpha']/255*100000)))
    else:
        shape.fill.background()
    if obj.get('stroke'):
        shape.line.color.rgb = RGBColor(*_rgb(obj.get('stroking_color', 0)))
        shape.line.width = Pt(max(.05, float(obj.get('linewidth') or .5) * scale))
        if obj.get('_ppt_stroke_alpha', 255) < 255:
            shape.line.color._xFill.get_or_change_to_srgbClr().append(_element('a:alpha', val=round(obj['_ppt_stroke_alpha']/255*100000)))
        dash = obj.get('dash')
        if dash and dash[0]:
            line = shape._element.spPr.get_or_add_ln()
            custom = _element('a:custDash')
            lengths = list(dash[0])
            if len(lengths) % 2:
                lengths *= 2
            unit = max(.05, float(obj.get('linewidth') or .5))
            for stroke, gap in zip(lengths[::2], lengths[1::2]):
                custom.append(_element('a:ds', d=max(1, round(stroke / unit * 100000)),
                                       sp=max(1, round(gap / unit * 100000))))
            line.append(custom)
    else:
        shape.line.fill.background()
    # add_shape() includes a theme effectRef by default.  Its inherited
    # shadow makes a PDF's thin rules look doubled in actual PowerPoint.
    # Every visual property here comes from the PDF, so disable the theme
    # style and explicitly retain an empty effect list.
    for child in list(shape._element):
        if child.tag.endswith('}style'):
            shape._element.remove(child)
    props = shape._element.spPr
    for child in list(props):
        if child.tag.endswith(('}effectLst', '}effectDag')):
            props.remove(child)
    props.append(_element('a:effectLst'))


def _add_vector(slide, obj, transform):
    """Use native custom geometry, including cubic Bézier paths."""
    from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
    ox, oy, scale = transform
    left, top, right, bottom = obj['x0'], obj['top'], obj['x1'], obj['bottom']
    if not obj.get('stroke') and not obj.get('fill'):
        return None
    width, height = max(.01, right - left), max(.01, bottom - top)
    shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE,
                                  _emu(ox + left * scale), _emu(oy + top * scale),
                                  _emu(width * scale), _emu(height * scale))
    shape.name = 'PDF图形 ' + obj.get('object_type', 'path')
    if obj.get('object_type') != 'rect':
        props = shape._element.spPr
        for old in list(props):
            if old.tag.endswith(('prstGeom', 'custGeom')):
                props.remove(old)
        geometry = _element('a:custGeom')
        for node in ('a:avLst', 'a:gdLst', 'a:ahLst', 'a:cxnLst'):
            geometry.append(_element(node))
        geometry.append(_element('a:rect', l='0', t='0', r='r', b='b'))
        paths = _element('a:pathLst')
        path = _element('a:path', w=_emu(width), h=_emu(height),
                        fill='norm' if obj.get('fill') else 'none',
                        stroke='1' if obj.get('stroke') else '0')
        commands = obj.get('path') or [('m', obj['pts'][0])] + [('l', p) for p in obj['pts'][1:]]
        for command in commands:
            kind, points = command[0], command[1:]
            if kind == 'h':
                path.append(_element('a:close'))
                continue
            tag = {'m': 'a:moveTo', 'l': 'a:lnTo', 'c': 'a:cubicBezTo', 'q': 'a:quadBezTo'}.get(kind)
            if tag is None:
                continue
            node = _element(tag)
            for x, y in points:
                node.append(_element('a:pt', x=_emu(x - left), y=_emu(y - top)))
            path.append(node)
        paths.append(path)
        geometry.append(paths)
        props.insert(1, geometry)
    _paint_shape(shape, obj, scale)
    return shape


class _PageRaster:
    def __init__(self, page, width, height):
        self.page, self.width, self.height = page, width, height
        self.image = None

    def crop(self, box):
        if self.image is None:
            bitmap = self.page.render(scale=3)
            try:
                pil = bitmap.to_pil()
                try:
                    self.image = pil.convert('RGB')
                finally:
                    pil.close()
            finally:
                bitmap.close()
        sx, sy = self.image.width / self.width, self.image.height / self.height
        coords = (max(0, int(math.floor(box[0] * sx))), max(0, int(math.floor(box[1] * sy))),
                  min(self.image.width, int(math.ceil(box[2] * sx))),
                  min(self.image.height, int(math.ceil(box[3] * sy))))
        image = self.image.crop(coords)
        try:
            stream = BytesIO()
            image.save(stream, format='PNG')
            stream.seek(0)
            return stream
        finally:
            image.close()

    def close(self):
        if self.image is not None:
            self.image.close()


def _picture(slide, stream, box, transform, label):
    ox, oy, scale = transform
    shape = slide.shapes.add_picture(stream, _emu(ox + box[0] * scale), _emu(oy + box[1] * scale),
                                     width=_emu((box[2] - box[0]) * scale),
                                     height=_emu((box[3] - box[1]) * scale))
    shape.name = label
    return shape


def _image_stream(obj):
    """Render the independent image object including alpha/rotation masks."""
    bitmap = obj.get_bitmap(render=True)
    try:
        pil = bitmap.to_pil()
        try:
            stream = BytesIO()
            pil.save(stream, format='PNG')
            stream.seek(0)
            return stream
        finally:
            pil.close()
    finally:
        bitmap.close()


def rebuild_pdf_pptx(source, generated, pages, check_cancel, progress=None):
    import pdfplumber
    import pypdfium2 as pdfium
    from pypdfium2 import raw
    from pptx import Presentation
    stats = PptxLayoutStats()
    deck = Presentation()
    document = pdfium.PdfDocument(source)
    try:
        with pdfplumber.open(source) as parsed:
            if len(parsed.pages) != pages:
                raise ValueError('PDF 页数在处理期间发生变化，请重试。')
            first = parsed.pages[0]
            factor = min(1, 56 * 72 / max(first.width, first.height))
            width, height = first.width * factor, first.height * factor
            deck.slide_width, deck.slide_height = _emu(width), _emu(height)
            for index, page in enumerate(parsed.pages):
                check_cancel()
                fit = min(width / page.width, height / page.height)
                transform = ((width - page.width * fit) / 2, (height - page.height * fit) / 2, fit)
                slide = deck.slides.add_slide(deck.slide_layouts[6])
                pdf_page = document[index]
                raster = _PageRaster(pdf_page, page.width, page.height)
                try:
                    if not any(c.get('text', '').strip() for c in page.chars):
                        box = (0, 0, page.width, page.height)
                        _picture(slide, raster.crop(box), box, transform, 'PDF扫描页 原版外观')
                        stats.images += 1
                        stats.scanned_pages += 1
                    else:
                        from pdf_pptx_tables import find_tables, contains_vector, write_table
                        from pdf_pptx_paths import compound_vectors
                        _annotate_order(page)
                        _annotate_vector_alpha(page, pdf_page)
                        _annotate_chars(page, pdf_page)
                        tables, rules, table_chars = find_tables(page)
                        drawings = []
                        # Preserve fills/lines as native shapes; thin filled
                        # rectangles are the separators used in many CV PDFs.
                        for obj in compound_vectors(page, pdf_page):
                            check_cancel()
                            if any(contains_vector(obj, table) for table in tables):
                                continue
                            vector = _add_vector(slide, obj, transform)
                            if vector is not None:
                                stats.shapes += 1
                                drawings.append((obj.get('_ppt_order', 0), vector._element))
                        images = list(pdf_page.get_objects(filter=[raw.FPDF_PAGEOBJ_IMAGE]))
                        used_images = set()
                        for image in page.images:
                            check_cancel()
                            box = (max(0, image['x0']), max(0, image['top']),
                                   min(page.width, image['x1']), min(page.height, image['bottom']))
                            if box[2] <= box[0] or box[3] <= box[1]:
                                continue
                            selected = None
                            for image_index, obj in enumerate(images):
                                if image_index in used_images:
                                    continue
                                a, b, c, d = obj.get_bounds()
                                error = abs(a - box[0]) + abs(c - box[2]) + abs(page.height - d - box[1]) + abs(page.height - b - box[3])
                                if error < 2:
                                    selected = obj
                                    used_images.add(image_index)
                                    break
                            try:
                                stream = _image_stream(selected) if selected is not None else raster.crop(box)
                            except (RuntimeError, ValueError):
                                stream = raster.crop(box)
                            picture = _picture(slide, stream, box, transform, 'PDF图片 可独立移动')
                            drawings.append((image.get('_ppt_order', 0), picture._element))
                            stats.images += 1
                        for table in tables:
                            native = write_table(slide, page, table, rules, transform, _write_text, _element, _emu)
                            order = max((c.get('_ppt_order', 0) for cell in table.cells for c in cell.chars), default=0)
                            drawings.append((order, native._element))
                            stats.tables += 1
                            stats.cells += len(table.cells)
                        for chars in _segments(page, [c for c in page.chars if id(c) not in table_chars]):
                            check_cancel()
                            text = ''.join(c['text'] for c in chars)
                            if re.search(r'\(cid:\d+\)|[\ufffd\ue000-\uf8ff]', text):
                                # A PDF can omit the character-to-Unicode map.
                                # Keep its visible content rather than silently
                                # outputting an invented or missing character.
                                box = (min(c['x0'] for c in chars), min(c['top'] for c in chars),
                                       max(c['x1'] for c in chars), max(c['bottom'] for c in chars))
                                shape = _picture(slide, raster.crop(box), box, transform, 'PDF未映射符号 原版外观')
                                stats.images += 1
                                stats.raster_fragments += 1
                            else:
                                for char in chars:
                                    char['_page_height'] = page.height
                                shape = _write_text(slide, chars, transform)
                                stats.text_boxes += 1
                            drawings.append((max(c.get('_ppt_order', 0) for c in chars), shape._element))
                        for _order, element in sorted(drawings, key=lambda item: item[0]):
                            slide.shapes._spTree.remove(element)
                            slide.shapes._spTree.append(element)
                finally:
                    raster.close()
                    pdf_page.close()
                if progress:
                    progress(index + 1, pages)
            check_cancel()
            deck.save(generated)
    finally:
        document.close()
    return stats
