"""Rebuild editable PDF tables/text, free page pictures and vector artwork.

PDF has no table/paragraph semantics. This module uses visible ruling geometry,
not the plain-text extraction order, and keeps unsupported scan pages as images.
Coordinates stay relative to each source page; pictures never inherit a table
or legacy text-frame container that restricts Word's dragging operation.
"""
from dataclasses import dataclass
from bisect import insort
from io import BytesIO
from math import hypot
from statistics import median


@dataclass
class LayoutStats:
    tables: int = 0
    cells: int = 0
    images: int = 0
    scanned_pages: int = 0
    vector_shapes: int = 0


def _inside(char, box):
    x = (char['x0'] + char['x1']) / 2
    y = (char['top'] + char['bottom']) / 2
    return box[0] - .05 <= x < box[2] + .05 and box[1] - .05 <= y < box[3] + .05


def _cluster(values, tolerance=.7):
    groups = []
    for value in sorted(values):
        if groups and abs(value - median(groups[-1])) <= tolerance:
            groups[-1].append(value)
        else:
            groups.append([value])
    return [median(group) for group in groups]


def find_ruled_tables(page):
    # White page/background rectangles have edges in pdfplumber, but are not
    # visible rules. Treating those as borders swallows titles and footnotes.
    def visible_rule(obj):
        color=obj.get('stroking_color')
        # Invisible white rules are often form-option layout boxes. They must
        # not turn an ordinary checkbox list into stacked native tables.
        if obj.get('object_type')=='line' or obj.get('stroke',False):
            if isinstance(color,(int,float)) and color>=.99:return False
            if isinstance(color,(list,tuple)) and len(color) in (1,3) and min(color)>=.99:return False
        if obj.get('object_type') != 'rect' or obj.get('stroke', False):
            return True
        # Some producers paint narrow black rectangles instead of stroked
        # lines. Keep these rules, while excluding large shading backgrounds.
        if not obj.get('fill') or min(obj.get('width', 0), obj.get('height', 0)) > 2:
            return False
        color = obj.get('non_stroking_color')
        if isinstance(color, (int, float)):
            return color < .8
        if isinstance(color, (list, tuple)) and len(color) in (1, 3):
            return sum(color) / len(color) < .8
        if isinstance(color, (list, tuple)) and len(color) == 4:
            return color[3] > .2 or sum(color[:3]) / 3 > .2
        return True
    visible = page.filter(visible_rule)
    candidates = visible.find_tables({
        'vertical_strategy': 'lines', 'horizontal_strategy': 'lines',
        'snap_tolerance': 2, 'join_tolerance': 2, 'intersection_tolerance': 3,
    })
    return [table for table in candidates if len(table.cells) >= 2]


def _baseline(char):
    """Use the PDF text origin, not font-specific bounding-box ascenders."""
    matrix = char.get('matrix')
    if matrix and len(matrix) == 6 and abs(matrix[1]) < abs(matrix[0]) * .05:
        return -matrix[5]
    return char['bottom'] - max(1, char.get('size', 10)) * .2


def _lines(chars):
    lines = []
    # Cache each line's exact median rather than re-sorting every line's
    # glyphs for every new glyph. Keep the original nearest-line/tie policy.
    cached = []
    def middle(values):
        i=len(values)//2
        return values[i] if len(values)%2 else (values[i-1]+values[i])/2
    for char in sorted(chars, key=lambda c: (_baseline(c), c['x0'])):
        if not char.get('text') or char.get('text') == '\x00':
            continue
        size = max(1, char.get('size', char['bottom'] - char['top']))
        # Chinese, Latin digits and a large separator glyph can share the same
        # baseline while their PDF bounding-box tops differ by several points.
        # Looking only at recently created lines also lost interleaved columns.
        base=_baseline(char)
        candidates = [(abs(entry[2]-base), index) for index,entry in enumerate(cached)]
        distance, index = min(candidates, key=lambda item:item[0]) if candidates else (float('inf'), None)
        if index is not None and distance > max(.7,min(size,cached[index][3])*.10):
            index = None
        if index is None:
            lines.append([char])
            cached.append(([base],[char.get('size',10)],base,char.get('size',10)))
        else:
            lines[index].append(char)
            bases,sizes,_,_=cached[index]
            insort(bases,base);insort(sizes,char.get('size',10))
            cached[index]=(bases,sizes,middle(bases),middle(sizes))
    for line in lines:
        line.sort(key=lambda c: c['x0'])
    return sorted(lines, key=lambda line: (min(c['top'] for c in line), min(c['x0'] for c in line)))


def _outside_segments(chars):
    for line in _lines(chars):
        segment = []
        for char in line:
            # Separate right-aligned dates/columns before Word can reflow them
            # behind a long, horizontally compressed title.
            if segment and char['x0'] - segment[-1]['x1'] > max(6, char.get('size', 10) * .7):
                yield segment
                segment = []
            segment.append(char)
        if segment:
            yield segment


def _font_name(name):
    name = name.split('+', 1)[-1]
    return {'SimSun': '宋体', 'SimHei': '黑体', 'KaiTi': '楷体',
            'KaiTi_GB2312': '楷体', 'FangSong': '仿宋', 'FangSong_GB2312': '仿宋',
            'MicrosoftYaHei': '微软雅黑', 'MicrosoftYaHeiUI': '微软雅黑',
            'ArialMT': 'Arial', 'TimesNewRomanPSMT': 'Times New Roman',
            'TimesNewRomanPS-BoldMT': 'Times New Roman',
            'TimesNewRomanPS-ItalicMT': 'Times New Roman',
            'TimesNewRomanPS-BoldItalicMT': 'Times New Roman',
            'Arial-BoldMT': 'Arial', 'Arial-ItalicMT': 'Arial',
            'Helvetica': 'Arial', 'Helvetica-Bold': 'Arial'}.get(name, name)


def _rgb_hex(color):
    if isinstance(color, (int, float)):
        color = (color,) * 3
    if not isinstance(color, (tuple, list)):
        return '000000'
    if len(color) == 1:
        color = color * 3
    elif len(color) == 4:
        c, m, y, k = color
        color = tuple((1 - component) * (1 - k) for component in (c, m, y))
    if len(color) != 3:
        return '000000'
    return ''.join(f'{max(0, min(255, round(float(value) * 255))):02X}' for value in color)


def _annotate_bold(page, pdf_page):
    """Recognize synthetic bold made by fill-and-stroke PDF text rendering."""
    from pypdfium2 import raw
    regions = []
    for obj in pdf_page.get_objects(filter=[raw.FPDF_PAGEOBJ_TEXT]):
        try:
            mode = raw.FPDFTextObj_GetTextRenderMode(obj)
            bold = mode in (1, 2, 5, 6) or obj.get_font().get_weight() >= 600
            if bold:
                left, bottom, right, top = obj.get_bounds()
                regions.append((left, page.height - top, right, page.height - bottom))
        except (AttributeError, RuntimeError, ValueError):
            # Older PDFs may omit font metadata. Font-name matching still works.
            continue
    for char in page.chars:
        char['_ruyi_bold'] = any(_inside(char, box) for box in regions)


def _write_line(paragraph, chars, scale, available_width=None):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Pt, RGBColor
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.keep_together = True
    paragraph.paragraph_format.widow_control = False
    props = paragraph._p.get_or_add_pPr()
    # PDF already supplies the spacing. Word's automatic CJK/Latin and
    # CJK/number spacing can otherwise push a one-line title onto two lines.
    props.append(_element('w:autoSpaceDE', val=0))
    props.append(_element('w:autoSpaceDN', val=0))
    size = max(c.get('size', 10) for c in chars) * scale
    source_width = (max(c['x1'] for c in chars) - min(c['x0'] for c in chars)) * scale
    fit = min(1, available_width / source_width) if available_width and source_width > 0 else 1
    # A small safety allowance keeps Word's font metrics from wrapping a cell
    # whose text exactly fills the PDF glyph bounds.
    fit = max(.5, fit)
    paragraph.paragraph_format.line_spacing = Pt(size * 1.12)
    previous = None
    for char_index, char in enumerate(chars):
        text = char['text']
        if previous:
            gap = char['x0'] - previous['x1']
            if gap > max(1.6, char.get('size', 10) * .25) and not previous['text'].endswith(' ') and not text.startswith(' '):
                paragraph.add_run(' ' * max(1, min(16, round(gap / max(2, char.get('size', 10) * .45)))))
        run = paragraph.add_run(text)
        font_name = _font_name(char.get('fontname', 'SimSun'))
        run.font.name = font_name
        run.font.size = Pt(max(4, char.get('size', 10) * scale * fit))
        run.font.bold = char.get('_ruyi_bold', False) or any(token in char.get('fontname', '').lower() for token in ('bold', 'black'))
        color = char.get('non_stroking_color')
        if color is not None:
            run.font.color.rgb = RGBColor.from_string(_rgb_hex(color))
        run.font.italic = any(token in char.get('fontname', '').lower() for token in ('italic', 'oblique'))
        fonts = run._element.get_or_add_rPr().find(qn('w:rFonts'))
        fonts.set(qn('w:eastAsia'), font_name)
        fonts.set(qn('w:cs'), font_name)
        # Preserve the PDF text matrix horizontal scale (e.g. the user's
        # compressed Chinese report title). Character gaps alone cannot do so.
        matrix = char.get('matrix')
        if matrix and len(matrix) == 6:
            vertical = hypot(matrix[2], matrix[3])
            if vertical > 0:
                width_ratio = max(.1, min(6, hypot(matrix[0], matrix[1]) / vertical))
                if abs(width_ratio - 1) > .002:
                    run._element.get_or_add_rPr().append(_element('w:w', val=round(width_ratio * 100)))
        # Preserve PDF character tracking, especially Chinese form titles.
        next_index = char_index + 1
        if next_index < len(chars):
            gap = chars[next_index]['x0'] - char['x1']
            if -.2 <= gap <= char.get('size', 10) * .2:
                tracking = OxmlElement('w:spacing')
                tracking.set(qn('w:val'), str(round(gap * scale * 20)))
                run._element.get_or_add_rPr().append(tracking)
        previous = char


def _element(tag, **attributes):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    node = OxmlElement(tag)
    for name, value in attributes.items():
        node.set(qn('w:' + name), str(value))
    return node


def _small_anchor(paragraph):
    from docx.shared import Pt
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = Pt(1)
    paragraph.add_run().font.size = Pt(1)


def _add_text_frame(document, chars, scale):
    from docx.shared import Pt
    paragraph = document.add_paragraph()
    left = min(c['x0'] for c in chars) * scale
    top = min(c['top'] for c in chars) * scale
    width = (max(c['x1'] for c in chars) - min(c['x0'] for c in chars)) * scale
    height = max(c.get('size', 10) for c in chars) * scale * 1.3
    paragraph._p.get_or_add_pPr().append(_element(
        'w:framePr', w=round((width + max(6, width * .06)) * 20), h=round(height * 20),
        x=round(left * 20), y=round(top * 20), hAnchor='page', vAnchor='page',
        hRule='atLeast', wrap='none', hSpace=0, vSpace=0))
    _write_line(paragraph, chars, scale)
    return paragraph


def _anchor_picture(paragraph, image_stream, box, scale):
    """Use an unconstrained page anchor in an ordinary body paragraph.

    Word restricts dragging when a page-relative anchor is owned by a table
    cell or an old-style text frame. layoutInCell=0 alone does not override it.
    """
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Pt
    picture = paragraph.add_run().add_picture(image_stream, width=Pt((box[2] - box[0]) * scale),
                                              height=Pt((box[3] - box[1]) * scale))
    inline = picture._inline
    anchor = OxmlElement('wp:anchor')
    for key, value in {'distT': '0', 'distB': '0', 'distL': '0', 'distR': '0',
                       'simplePos': '0', 'relativeHeight': '251658240', 'behindDoc': '0',
                       'locked': '0', 'layoutInCell': '0', 'allowOverlap': '1'}.items():
        anchor.set(key, value)
    simple = OxmlElement('wp:simplePos'); simple.set('x', '0'); simple.set('y', '0'); anchor.append(simple)
    for axis, position in (('H', box[0]), ('V', box[1])):
        node = OxmlElement('wp:position' + axis); node.set('relativeFrom', 'page')
        offset = OxmlElement('wp:posOffset'); offset.text = str(int(Pt(position * scale))); node.append(offset)
        anchor.append(node)
    for tag in ('wp:extent', 'wp:effectExtent'):
        node = inline.find(qn(tag))
        if node is not None:
            anchor.append(node)
    anchor.append(OxmlElement('wp:wrapNone'))
    for tag in ('wp:docPr', 'wp:cNvGraphicFramePr', 'a:graphic'):
        node = inline.find(qn(tag))
        if node is not None:
            anchor.append(node)
    inline.getparent().replace(inline, anchor)


def _page_picture(document, image_stream, box, scale):
    paragraph = document.add_paragraph()
    _small_anchor(paragraph)
    _anchor_picture(paragraph, image_stream, box, scale)
    return paragraph


def _shape_path(obj, left, top, factor=1000):
    """Translate PDF move/line/cubic/close paths into editable VML geometry."""
    parts = []
    for operation in obj.get('path', []):
        command, *points = operation
        if command == 'h':
            parts.append('x')
            continue
        if command not in ('m', 'l', 'c'):
            return None
        numbers = []
        for point in points:
            if not isinstance(point, (tuple, list)) or len(point) != 2:
                return None
            numbers.extend((round((point[0] - left) * factor), round((point[1] - top) * factor)))
        parts.append(command + ' ' + ','.join(map(str, numbers)))
    return ' '.join(parts) + ' e' if parts else None


def _add_vectors(document, page, tables, scale, pictures, stats, *, covered_regions=()):
    """Keep page separators/backgrounds and editable PDF vector shapes.

    Table ruling geometry is already represented by actual Word cell borders.
    Other page objects must not disappear just because they are not tables.
    """
    from lxml import etree
    from docx.oxml import OxmlElement
    vns, ons = 'urn:schemas-microsoft-com:vml', 'urn:schemas-microsoft-com:office:office'
    for obj in list(page.rects) + list(page.lines) + list(page.curves):
        kind = obj.get('object_type')
        fill, stroke = bool(obj.get('fill')), bool(obj.get('stroke', kind == 'line'))
        if not fill and not stroke:
            continue
        left, top, right, bottom = obj['x0'], obj['top'], obj['x1'], obj['bottom']
        if right < left or bottom < top:
            continue
        # These marks are already painted inside a complete source-page crop.
        if any(box[0]-.1 <= left and box[1]-.1 <= top and
               box[2]+.1 >= right and box[3]+.1 >= bottom for box in covered_regions):
            continue
        # Omit invisible white page rectangles. Colored or smaller backgrounds
        # remain editable shapes behind text rather than raster page overlays.
        if fill and not stroke and _rgb_hex(obj.get('non_stroking_color')) == 'FFFFFF':
            continue
        rule = kind == 'line' or min(right - left, bottom - top) <= 2
        if rule and any(table.bbox[0] - 1 <= left and right <= table.bbox[2] + 1
                        and table.bbox[1] - 1 <= top and bottom <= table.bbox[3] + 1 for table in tables):
            continue
        width, height = max(.01, right - left), max(.01, bottom - top)
        shape_kind = 'rect' if kind == 'rect' else 'shape'
        path = None if shape_kind == 'rect' else _shape_path(obj, left, top)
        if shape_kind == 'shape' and path is None:
            # pdfplumber line objects may expose only two endpoints.
            points = obj.get('pts', [])
            if kind == 'line' and len(points) >= 2:
                path = 'm {},{} l {},{} e'.format(
                    round((points[0][0] - left) * 1000), round((points[0][1] - top) * 1000),
                    round((points[-1][0] - left) * 1000), round((points[-1][1] - top) * 1000))
            else:
                # Retain an unusual unsupported local vector without flattening
                # the whole editable page. Rendering supplies its exact marks.
                if width >= .1 and height >= .1:
                    _page_picture(document, pictures.crop((left, top, right, bottom)),
                                  (left, top, right, bottom), scale)
                    stats.images += 1
                continue
        paragraph = document.add_paragraph()
        _small_anchor(paragraph)
        pict = OxmlElement('w:pict')
        shape = etree.SubElement(pict, '{' + vns + '}' + shape_kind, nsmap={'v': vns, 'o': ons})
        shape.set('id', 'RuyiVector' + str(stats.vector_shapes + 1))
        shape.set('style', 'position:absolute;margin-left:{:.5f}pt;margin-top:{:.5f}pt;'
                  'width:{:.5f}pt;height:{:.5f}pt;z-index:-251658239;'
                  'mso-position-horizontal-relative:page;mso-position-vertical-relative:page'.format(
                      left * scale, top * scale, width * scale, height * scale))
        shape.set('filled', 't' if fill else 'f')
        shape.set('stroked', 't' if stroke else 'f')
        shape.set('fillcolor', '#' + _rgb_hex(obj.get('non_stroking_color')))
        shape.set('strokecolor', '#' + _rgb_hex(obj.get('stroking_color')))
        shape.set('strokeweight', str(max(.05, obj.get('linewidth', .5)) * scale) + 'pt')
        shape.set('{' + ons + '}allowincell', 'f')
        shape.set('{' + ons + '}allowoverlap', 't')
        if path is not None:
            shape.set('coordorigin', '0,0')
            shape.set('coordsize', f'{max(1, round(width * 1000))},{max(1, round(height * 1000))}')
            shape.set('path', path)
        dash = obj.get('dash')
        if stroke and dash and dash[0]:
            stroke_node = etree.SubElement(shape, '{' + vns + '}stroke')
            stroke_node.set('dashstyle', 'dash' if max(dash[0]) >= 2 else 'dot')
        paragraph.add_run()._r.append(pict)
        stats.vector_shapes += 1


class _PagePictures:
    def __init__(self, pdf_page, width, height, *, render_scale=2):
        self.pdf_page = pdf_page
        self.width = width
        self.height = height
        self.render_scale = render_scale
        self.image = None

    def crop(self, box):
        if self.image is None:
            bitmap = self.pdf_page.render(scale=self.render_scale)
            try:
                image = bitmap.to_pil()
                try:
                    self.image = image.convert('RGB')
                finally:
                    image.close()
            finally:
                bitmap.close()
        sx, sy = self.image.width / self.width, self.image.height / self.height
        coordinates = (max(0, round(box[0] * sx)), max(0, round(box[1] * sy)),
                       min(self.image.width, round(box[2] * sx)), min(self.image.height, round(box[3] * sy)))
        cropped = self.image.crop(coordinates)
        try:
            import pdf_word_images
            original = None
            if pdf_word_images.OPTIMIZE_WORD_IMAGES:
                original = pdf_word_images.matching_source_jpeg(self.pdf_page, box, cropped, self.height)
            return pdf_word_images.encode_word_image(cropped, original)
        finally:
            cropped.close()

    def close(self):
        if self.image is not None:
            self.image.close()


def _set_cell_style(cell, size):
    from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
    tcpr = cell._tc.get_or_add_tcPr()
    borders = _element('w:tcBorders')
    for side in ('top', 'left', 'bottom', 'right'):
        borders.append(_element('w:' + side, val='single', sz=max(2, round(size * 8)), color='000000'))
    tcpr.append(borders)
    margins = _element('w:tcMar')
    for side in ('top', 'bottom', 'left', 'right'):
        margins.append(_element('w:' + side, w=0, type='dxa'))
    tcpr.append(margins)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def _finish_merged_borders(table, xs, ys, source_cells, size):
    """Style each physical vMerge continuation, not just its logical owner.

    python-docx's row.cells maps a continuation back to its restart cell. To
    retain an uninterrupted outside rule we must visit the actual OOXML tcs.
    """
    from docx.oxml.ns import qn
    for row_index, tr in enumerate(table._tbl.tr_lst):
        col_index = 0
        for tc in tr.tc_lst:
            props = tc.get_or_add_tcPr()
            span_node = props.find(qn('w:gridSpan'))
            span = int(span_node.get(qn('w:val'))) if span_node is not None else 1
            center_x = (xs[col_index] + xs[col_index + span]) / 2
            center_y = (ys[row_index] + ys[row_index + 1]) / 2
            owner = next((box for box in source_cells
                          if box[0] - .1 <= center_x < box[2] + .1
                          and box[1] - .1 <= center_y < box[3] + .1), None)
            old = props.find(qn('w:tcBorders'))
            if old is not None:
                props.remove(old)
            borders = _element('w:tcBorders')
            for side in ('top', 'left', 'bottom', 'right'):
                visible = owner is not None and (side in ('left', 'right')
                    or (side == 'top' and abs(ys[row_index] - owner[1]) < .8)
                    or (side == 'bottom' and abs(ys[row_index + 1] - owner[3]) < .8))
                borders.append(_element('w:' + side, val='single' if visible else 'nil',
                                        sz=max(2, round(size * 8)), color='000000'))
            props.append(borders)
            col_index += span


def _add_table(document, page, source_table, scale, pictures, images, stats):
    from docx.enum.table import WD_ROW_HEIGHT_RULE
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt
    from docx.oxml.ns import qn
    xs = _cluster([v for box in source_table.cells for v in (box[0], box[2])])
    ys = _cluster([v for box in source_table.cells for v in (box[1], box[3])])
    if len(xs) * len(ys) > 20000:
        raise ValueError('此页表格网格过于复杂，请拆分 PDF 后重试。')
    table = document.add_table(rows=len(ys) - 1, cols=len(xs) - 1)
    table.autofit = False
    props = table._tbl.tblPr
    existing_width = props.find(qn('w:tblW'))
    if existing_width is not None:
        props.remove(existing_width)
    props.append(_element('w:tblW', w=round((xs[-1] - xs[0]) * scale * 20), type='dxa'))
    props.append(_element('w:tblInd', w=0, type='dxa'))
    props.append(_element('w:tblpPr', leftFromText=0, rightFromText=0, topFromText=0, bottomFromText=0,
                          horzAnchor='page', vertAnchor='page',
                          tblpX=round(xs[0] * scale * 20), tblpY=round(ys[0] * scale * 20)))
    props.append(_element('w:tblOverlap', val='overlap'))
    for index, column in enumerate(table.columns):
        column.width = Pt((xs[index + 1] - xs[index]) * scale)
    for row_index, row in enumerate(table.rows):
        row.height = Pt((ys[row_index + 1] - ys[row_index]) * scale)
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
        row._tr.get_or_add_trPr().append(_element('w:cantSplit'))
        for col_index, cell in enumerate(row.cells):
            cell.width = Pt((xs[col_index + 1] - xs[col_index]) * scale)
    rule_widths = [float(line.get('linewidth', .5)) for line in page.lines
                   if source_table.bbox[0] - 1 <= line['x0'] <= source_table.bbox[2] + 1
                   and source_table.bbox[1] - 1 <= line['top'] <= source_table.bbox[3] + 1]
    border_size = median(rule_widths) * scale if rule_widths else .5
    for box in sorted(source_table.cells, key=lambda b: (b[1], b[0])):
        start_col = min(range(len(xs)), key=lambda i: abs(xs[i] - box[0]))
        end_col = min(range(len(xs)), key=lambda i: abs(xs[i] - box[2]))
        start_row = min(range(len(ys)), key=lambda i: abs(ys[i] - box[1]))
        end_row = min(range(len(ys)), key=lambda i: abs(ys[i] - box[3]))
        if start_col == end_col or start_row == end_row:
            continue
        cell = table.cell(start_row, start_col)
        if end_row > start_row + 1 or end_col > start_col + 1:
            cell = cell.merge(table.cell(end_row - 1, end_col - 1))
        _set_cell_style(cell, border_size)
        chars = [char for char in page.chars if _inside(char, box)]
        lines = _lines(chars)
        p = cell.paragraphs[0]
        if lines:
            for number, line in enumerate(lines):
                if number:
                    p = cell.add_paragraph()
                text_left, text_right = min(c['x0'] for c in line), max(c['x1'] for c in line)
                if abs((text_left + text_right) / 2 - (box[0] + box[2]) / 2) < max(4, (box[2] - box[0]) * .08):
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                else:
                    p.paragraph_format.left_indent = Pt(max(0, text_left - box[0]) * scale)
                _write_line(p, line, scale, max(1, (box[2] - box[0] - .7) * scale))
            if len(lines) > 1:
                spacing = median(min(c['top'] for c in b) - min(c['top'] for c in a)
                                 for a, b in zip(lines, lines[1:]))
                for p in cell.paragraphs:
                    p.paragraph_format.line_spacing = Pt(max(4, spacing * scale))
        else:
            _small_anchor(p)
        stats.cells += 1
    _finish_merged_borders(table, xs, ys, source_table.cells, border_size)
    # A paragraph after the floating table gives its anchor a stable body slot.
    _small_anchor(document.add_paragraph())
    stats.tables += 1


def rebuild_pdf_docx(source, generated, pages, check_cancel, progress=None):
    import pdfplumber
    import pypdfium2 as pdfium
    from docx import Document
    from docx.enum.section import WD_SECTION
    from docx.shared import Pt
    stats = LayoutStats()
    document = Document()
    normal = document.styles['Normal']
    normal.font.name = '宋体'
    normal.font.size = Pt(10)
    normal.paragraph_format.space_after = Pt(0)
    normal.paragraph_format.widow_control = False
    pdf = pdfium.PdfDocument(source)
    try:
        with pdfplumber.open(source) as extracted:
            for page_index, page in enumerate(extracted.pages):
                check_cancel()
                section = document.add_section(WD_SECTION.NEW_PAGE) if page_index else document.sections[0]
                scale = min(1, 22 * 72 / max(page.width, page.height))
                section.page_width = Pt(page.width * scale)
                section.page_height = Pt(page.height * scale)
                section.top_margin = section.bottom_margin = Pt(0)
                section.left_margin = section.right_margin = Pt(0)
                section.header_distance = section.footer_distance = Pt(0)
                _small_anchor(document.add_paragraph())
                pdf_page = pdf[page_index]
                pictures = _PagePictures(pdf_page, page.width, page.height)
                try:
                    if not any(c.get('text', '').strip() for c in page.chars):
                        box = (0, 0, page.width, page.height)
                        _page_picture(document, pictures.crop(box), box, scale)
                        stats.scanned_pages += 1
                        stats.images += 1
                    else:
                        _annotate_bold(page, pdf_page)
                        tables = find_ruled_tables(page)
                        _add_vectors(document, page, tables, scale, pictures, stats)
                        images = [(number, (max(0, image['x0']), max(0, image['top']),
                                            min(page.width, image['x1']), min(page.height, image['bottom'])))
                                  for number, image in enumerate(page.images)
                                  if image['x1'] > image['x0'] and image['bottom'] > image['top']]
                        outside = [char for char in page.chars if not any(_inside(char, table.bbox) for table in tables)]
                        for segment in _outside_segments(outside):
                            _add_text_frame(document, segment, scale)
                        for table in tables:
                            check_cancel()
                            _add_table(document, page, table, scale, pictures, images, stats)
                        for image_index, box in images:
                            _page_picture(document, pictures.crop(box), box, scale)
                            stats.images += 1
                finally:
                    pictures.close()
                    pdf_page.close()
                if progress:
                    progress(page_index + 1, pages)
        document.save(generated)
    finally:
        pdf.close()
    return stats
