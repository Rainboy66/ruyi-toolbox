"""Rebuild PDF pages as editable Excel geometry, retaining pictures and rules.

PDF glyph coordinates are the authority. Values deliberately remain text: a
PDF is a printed representation, so a leading zero, date, symbol or ``=`` must
not be reinterpreted by Excel. Scans are retained as pictures without OCR.
"""
from bisect import bisect_right
from copy import copy
from dataclasses import dataclass, field
from io import BytesIO
from statistics import median
import re


@dataclass
class XlsxLayoutStats:
    tables: int = 0
    cells: int = 0
    images: int = 0
    scanned_pages: int = 0
    text_characters: int = 0
    input_characters: int = 0
    rules: int = 0
    warnings: list[str] = field(default_factory=list)


def _cluster(values, tolerance=.5):
    groups = []
    for value in sorted(float(value) for value in values):
        if groups and abs(value - median(groups[-1])) <= tolerance:
            groups[-1].append(value)
        else:
            groups.append([value])
    return [median(group) for group in groups]


def _inside(char, box):
    x = (char['x0'] + char['x1']) / 2
    y = (char['top'] + char['bottom']) / 2
    return box[0] - .05 <= x < box[2] + .05 and box[1] - .05 <= y < box[3] + .05


def _visible_rule(obj):
    if obj.get('object_type') != 'rect' or obj.get('stroke', False):
        return True
    if not obj.get('fill') or min(obj.get('width', 0), obj.get('height', 0)) > 2:
        return False
    color = obj.get('non_stroking_color')
    if isinstance(color, (int, float)):
        return color < .8
    if isinstance(color, (tuple, list)) and len(color) in (1, 3):
        return sum(color) / len(color) < .8
    if isinstance(color, (tuple, list)) and len(color) == 4:
        return color[3] > .2 or sum(color[:3]) / 3 > .2
    return True


def _tables(page):
    return [table for table in page.filter(_visible_rule).find_tables({
        'vertical_strategy': 'lines', 'horizontal_strategy': 'lines',
        'snap_tolerance': 2, 'join_tolerance': 2, 'intersection_tolerance': 3,
    }) if len(table.cells) >= 2]


def _lines(chars):
    """Group common baselines; Latin and CJK font ascenders need not agree."""
    lines = []
    for char in sorted(chars, key=lambda c: ((c['top'] + c['bottom']) / 2, c['x0'])):
        center = (char['top'] + char['bottom']) / 2
        size = max(1., char.get('size', char['bottom'] - char['top']))
        candidates = [(abs(center - median((c['top'] + c['bottom']) / 2 for c in line)), line)
                      for line in lines[-4:]]
        nearest = min(candidates, default=(float('inf'), None), key=lambda pair: pair[0])
        if nearest[0] <= max(2.0, size * .32):
            nearest[1].append(char)
        else:
            lines.append([char])
    for line in lines:
        line.sort(key=lambda c: c['x0'])
    return sorted(lines, key=lambda line: (min(c['top'] for c in line), min(c['x0'] for c in line)))


def _segments(chars):
    for line in _lines(chars):
        segment = []
        for char in line:
            if segment and char['x0'] - segment[-1]['x1'] > max(10, char.get('size', 10) * 1.3):
                yield segment
                segment = []
            segment.append(char)
        if segment:
            yield segment


def _rgb(color, default='000000'):
    if isinstance(color, (int, float)):
        color = [color] * 3
    if not isinstance(color, (tuple, list)):
        return default
    if len(color) == 1:
        color = list(color) * 3
    if len(color) == 4:
        cyan, magenta, yellow, black = (float(v) for v in color)
        color = [(1 - component) * (1 - black) for component in (cyan, magenta, yellow)]
    if len(color) != 3:
        return default
    return ''.join(f'{max(0, min(255, round(float(value) * 255))):02X}' for value in color)


def _font_name(name):
    name = name.split('+', 1)[-1]
    return {'SimSun': '宋体', 'SimHei': '黑体', 'MicrosoftYaHei': '微软雅黑',
            'KaiTi': '楷体', 'ArialMT': 'Arial',
            'TimesNewRomanPSMT': 'Times New Roman',
            'TimesNewRomanPS-BoldMT': 'Times New Roman'}.get(name, name)


def _clean(text):
    # XML 1.0 cannot store these controls. Preserve meaningful whitespace and
    # Unicode, including non-BMP characters and private-use symbol glyphs.
    return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', text)


def _style_key(char):
    name = char.get('fontname', 'SimSun')
    return (_font_name(name), round(max(1, min(409, float(char.get('size', 10)))), 3),
            bool(char.get('_ruyi_bold')) or any(s in name.lower() for s in ('bold', 'black')),
            'italic' in name.lower() or 'oblique' in name.lower(),
            _rgb(char.get('non_stroking_color')))


def _text_value(lines):
    """Create rich text, retaining source glyphs and intentional inter-run gaps."""
    from openpyxl.cell.rich_text import CellRichText, TextBlock
    from openpyxl.cell.text import InlineFont
    runs = []
    for line_index, line in enumerate(lines):
        if line_index and runs:
            # Excel is stricter than generic XML readers about standalone
            # unformatted LF runs. Keep the break in the preceding styled run;
            # openpyxl then writes xml:space="preserve" for that trailing LF.
            runs[-1].text += '\n'
        last, key, text = None, None, ''
        for char in line:
            char_text = _clean(char.get('text', ''))
            if not char_text:
                continue
            current_key = _style_key(char)
            gap_text = ''
            if last is not None:
                gap = char['x0'] - last['x1']
                if gap > max(1.6, char.get('size', 10) * .25) and not last['text'].endswith(' ') and not char_text.startswith(' '):
                    gap_text = ' ' * max(1, min(24, round(gap / max(2, char.get('size', 10) * .45))))
            if key is not None and current_key != key:
                runs.append(TextBlock(InlineFont(rFont=key[0], sz=key[1], b=key[2], i=key[3], color=key[4]), text))
                text = ''
            key = current_key
            text += gap_text + char_text
            last = char
        if text:
            runs.append(TextBlock(InlineFont(rFont=key[0], sz=key[1], b=key[2], i=key[3], color=key[4]), text))
    value = CellRichText(runs)
    if len({_style_key(char) for line in lines for char in line}) <= 1:
        # A single complete style belongs on the cell. This preserves all font
        # properties while avoiding needless multiline rich-text records.
        return str(value)
    return value


class _Grid:
    def __init__(self, sheet, xs, ys):
        from openpyxl.utils import get_column_letter
        self.sheet, self.xs, self.ys = sheet, xs, ys
        previous_pixel = round(xs[0] * 96 / 72)
        for column, (left, right) in enumerate(zip(xs, xs[1:]), 1):
            # OOXML's width already includes cell padding. The Excel UI's
            # (pixels-5)/7 character count is a different value; using it here
            # repeatedly loses five pixels per coordinate column. Round page
            # boundaries cumulatively so tiny columns do not accumulate drift.
            next_pixel = max(previous_pixel + 1, round(right * 96 / 72))
            pixels = next_pixel - previous_pixel
            previous_pixel = next_pixel
            sheet.column_dimensions[get_column_letter(column)].width = min(255, pixels / 7)
        for row, (top, bottom) in enumerate(zip(ys, ys[1:]), 1):
            sheet.row_dimensions[row].height = max(.1, min(409.5, bottom - top))
        self.occupied = set()

    @staticmethod
    def _closest(values, value):
        index = bisect_right(values, value)
        return min({max(0, index - 1), min(len(values) - 1, index)}, key=lambda i: abs(values[i] - value))

    def indices(self, box):
        c0 = self._closest(self.xs, box[0]) + 1
        c1 = self._closest(self.xs, box[2])
        r0 = self._closest(self.ys, box[1]) + 1
        r1 = self._closest(self.ys, box[3])
        return r0, c0, max(r0, r1), max(c0, c1)

    def cell(self, box):
        r0, c0, r1, c1 = self.indices(box)
        region = {(r, c) for r in range(r0, r1 + 1) for c in range(c0, c1 + 1)}
        if self.occupied.intersection(region):
            return None
        if r0 != r1 or c0 != c1:
            self.sheet.merge_cells(start_row=r0, start_column=c0, end_row=r1, end_column=c1)
        self.occupied.update(region)
        return self.sheet.cell(r0, c0)

    def anchor(self, left, top):
        from openpyxl.drawing.spreadsheet_drawing import AnchorMarker
        from openpyxl.utils.units import pixels_to_EMU
        c = max(0, min(len(self.xs) - 2, bisect_right(self.xs, left) - 1))
        r = max(0, min(len(self.ys) - 2, bisect_right(self.ys, top) - 1))
        return AnchorMarker(col=c, row=r,
                            colOff=pixels_to_EMU(max(0, left - self.xs[c]) * 96 / 72),
                            rowOff=pixels_to_EMU(max(0, top - self.ys[r]) * 96 / 72))


def _decorate_cell(cell, chars, box, lines, bordered=False):
    from openpyxl.styles import Alignment, Border, Font, Side
    if lines:
        cell.value = _text_value(lines)
        cell.data_type = 's'
        common = max(chars, key=lambda char: len(char.get('text', '')) or 1)
        name, size, bold, italic, color = _style_key(common)
        cell.font = Font(name=name, size=size, bold=bold, italic=italic, color=color)
        left = min(c['x0'] for c in chars)
        right = max(c['x1'] for c in chars)
        centered = abs((left + right) / 2 - (box[0] + box[2]) / 2) < max(3, (box[2] - box[0]) * .06)
        right_aligned = abs(right - box[2]) < 3 and left - box[0] > 8
        horizontal = 'center' if bordered and centered else 'right' if right_aligned else 'left'
        vertical = 'center' if bordered else 'top'
        cell.alignment = Alignment(horizontal=horizontal, vertical=vertical, wrap_text=len(lines) > 1,
                                   shrink_to_fit=len(lines) == 1, indent=0)
        cell.number_format = '@'
    if bordered:
        side = Side(style='thin', color='000000')
        cell.border = Border(top=side, bottom=side, left=side, right=side)


def _add_picture(grid, picture_stream, box):
    from openpyxl.drawing.image import Image
    from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor
    from openpyxl.drawing.xdr import XDRPositiveSize2D
    from openpyxl.utils.units import pixels_to_EMU
    picture = Image(picture_stream)
    picture.width = (box[2] - box[0]) * 96 / 72
    picture.height = (box[3] - box[1]) * 96 / 72
    picture.anchor = OneCellAnchor(_from=grid.anchor(box[0], box[1]),
                                  ext=XDRPositiveSize2D(pixels_to_EMU(picture.width), pixels_to_EMU(picture.height)))
    grid.sheet.add_image(picture)


class _PagePictures:
    def __init__(self, pdf_page, width, height):
        self.pdf_page, self.width, self.height = pdf_page, width, height
        self.image = None

    def crop(self, box):
        if self.image is None:
            bitmap = self.pdf_page.render(scale=2.0)
            try:
                image = bitmap.to_pil()
                try:
                    self.image = image.convert('RGB')
                finally:
                    image.close()
            finally:
                bitmap.close()
        sx, sy = self.image.width / self.width, self.image.height / self.height
        coordinates = (max(0, int(box[0] * sx)), max(0, int(box[1] * sy)),
                       min(self.image.width, round(box[2] * sx)), min(self.image.height, round(box[3] * sy)))
        image = self.image.crop(coordinates)
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


def _rules(page):
    """Yield visible axis-aligned strokes and narrow painted rectangles."""
    seen = set()
    for obj in page.lines + page.rects:
        if not _visible_rule(obj):
            continue
        left, top, right, bottom = obj['x0'], obj['top'], obj['x1'], obj['bottom']
        color = _rgb(obj.get('stroking_color') if obj.get('stroke', True) else obj.get('non_stroking_color'))
        if color == 'FFFFFF':
            continue
        width = max(.25, float(obj.get('linewidth', .5)))
        if abs(bottom - top) <= 2:
            edges = [(left, (top + bottom) / 2, right, (top + bottom) / 2)]
        elif abs(right - left) <= 2:
            edges = [((left + right) / 2, top, (left + right) / 2, bottom)]
        elif obj.get('object_type') == 'rect' and obj.get('stroke'):
            edges = [(left, top, right, top), (left, bottom, right, bottom),
                     (left, top, left, bottom), (right, top, right, bottom)]
        else:
            continue
        for edge in edges:
            key = tuple(round(value, 1) for value in edge)
            if key not in seen:
                seen.add(key)
                yield edge, width, color, obj.get('dash')


def _add_rules(grid, rules, stats):
    from openpyxl.styles import Side
    for box, width, color, dash in rules:
        x0, y0, x1, y1 = box
        side = Side(style='dashed' if dash and dash[0] else 'medium' if width >= 1.5 else 'thin', color=color)
        if abs(y1 - y0) <= .1:
            rowline = grid._closest(grid.ys, y0)
            r = max(1, min(len(grid.ys) - 1, rowline))
            side_name = 'top' if rowline == 0 else 'bottom'
            start = grid._closest(grid.xs, x0) + 1
            end = grid._closest(grid.xs, x1)
            for col in range(start, max(start, end) + 1):
                cell = grid.sheet.cell(r, col)
                border = copy(cell.border)
                setattr(border, side_name, side)
                cell.border = border
        else:
            colline = grid._closest(grid.xs, x0)
            c = max(1, min(len(grid.xs) - 1, colline))
            side_name = 'left' if colline == 0 else 'right'
            start = grid._closest(grid.ys, y0) + 1
            end = grid._closest(grid.ys, y1)
            for row in range(start, max(start, end) + 1):
                cell = grid.sheet.cell(row, c)
                border = copy(cell.border)
                setattr(border, side_name, side)
                cell.border = border
        stats.rules += 1


def _add_fills(grid, rectangles):
    from openpyxl.styles import PatternFill
    for box, color in rectangles:
        r0, c0, r1, c1 = grid.indices(box)
        fill = PatternFill(fill_type='solid', fgColor=color)
        for row in range(r0, r1 + 1):
            for col in range(c0, c1 + 1):
                grid.sheet.cell(row, col).fill = fill


def _page_cells(page, stats):
    chars = [dict(char, _source_index=i) for i, char in enumerate(page.chars) if _clean(char.get('text', ''))]
    stats.input_characters += sum(len(_clean(c.get('text', ''))) for c in chars)
    cells, assigned = [], set()
    for table in _tables(page):
        stats.tables += 1
        for box in sorted(table.cells, key=lambda item: (item[1], item[0])):
            contained = [char for char in chars if char['_source_index'] not in assigned and _inside(char, box)]
            assigned.update(char['_source_index'] for char in contained)
            cells.append((box, contained, _lines(contained), True))
    for segment in _segments([char for char in chars if char['_source_index'] not in assigned]):
        left, top = min(c['x0'] for c in segment), min(c['top'] for c in segment)
        right, bottom = max(c['x1'] for c in segment), max(c['bottom'] for c in segment)
        # Reserve the original glyph width plus Excel's cell padding. A
        # neighboring date keeps its own coordinate region rather than tabs.
        right = min(page.width, right + max(3, median(c.get('size', 10) for c in segment) * .30))
        box = (max(0, left), max(0, top - .25), right, min(page.height, bottom + .75))
        cells.append((box, segment, [segment], False))
    return cells


def rebuild_pdf_xlsx(source, generated, pages, check_cancel, progress=None):
    import pdfplumber
    import pypdfium2 as pdfium
    from openpyxl import Workbook
    from openpyxl.styles import Font
    from openpyxl.worksheet.page import PageMargins
    from openpyxl.utils import get_column_letter
    stats = XlsxLayoutStats()
    workbook = Workbook()
    # Column geometry is based on Calibri 11's seven-pixel maximum digit width.
    # Source fonts are applied to each text run rather than the Normal style.
    workbook._named_styles['Normal'].font = Font(name='Calibri', size=11)
    workbook.remove(workbook.active)
    document = pdfium.PdfDocument(source)
    streams = []
    try:
        with pdfplumber.open(source) as extracted:
            if len(extracted.pages) != pages:
                raise ValueError('PDF 页数在处理期间发生变化，请重试。')
            for index, page in enumerate(extracted.pages):
                check_cancel()
                sheet = workbook.create_sheet(f'第{index + 1}页')
                sheet.sheet_view.showGridLines = False
                sheet.sheet_view.zoomScale = 85
                sheet.sheet_properties.pageSetUpPr.fitToPage = True
                sheet.page_setup.orientation = 'landscape' if page.width > page.height else 'portrait'
                sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
                sheet.page_setup.fitToWidth = sheet.page_setup.fitToHeight = 1
                sheet.page_margins = PageMargins(left=0, right=0, top=0, bottom=0, header=0, footer=0)
                cells = _page_cells(page, stats)
                rules = list(_rules(page))
                fills = [((rect['x0'], rect['top'], rect['x1'], rect['bottom']),
                          _rgb(rect.get('non_stroking_color'), 'FFFFFF'))
                         for rect in page.rects if rect.get('fill') and rect['width'] > 2 and rect['height'] > 2
                         and _rgb(rect.get('non_stroking_color'), 'FFFFFF') != 'FFFFFF']
                image_boxes = [(max(0, image['x0']), max(0, image['top']), min(page.width, image['x1']), min(page.height, image['bottom']))
                               for image in page.images if image['x1'] > image['x0'] and image['bottom'] > image['top']]
                scanned = not any(char.get('text', '').strip() for char in page.chars)
                if scanned:
                    image_boxes = [(0, 0, page.width, page.height)]
                    stats.scanned_pages += 1
                xs = [0., page.width]
                ys = [0., page.height]
                for box, *_ in cells:
                    xs.extend((box[0], box[2]))
                    ys.extend((box[1], box[3]))
                for box in image_boxes:
                    xs.extend((box[0], box[2])); ys.extend((box[1], box[3]))
                for box, *_ in rules:
                    xs.extend((box[0], box[2])); ys.extend((box[1], box[3]))
                for box, _ in fills:
                    xs.extend((box[0], box[2])); ys.extend((box[1], box[3]))
                # Excel's native row/column dimensions have finite limits.
                # Split long blank spans without moving any source boundary.
                xs = _cluster(xs, .35)
                ys = _cluster(ys + [v for top, bottom in zip(_cluster(ys), _cluster(ys)[1:])
                                    for v in [top + 350 * step for step in range(1, int((bottom - top) / 350) + 1)]], .35)
                if len(xs) > 2048 or len(ys) > 10000 or len(xs) * len(ys) > 400000:
                    raise ValueError('此页 PDF 坐标过于复杂，无法安全重建 Excel 网格。请先拆分或选择保真模式。')
                grid = _Grid(sheet, xs, ys)
                _add_fills(grid, fills)
                # Table cells first; free text otherwise uses its own position.
                for box, chars, lines, bordered in sorted(cells, key=lambda item: not item[3]):
                    check_cancel()
                    cell = grid.cell(box)
                    if cell is None:
                        # Rare overlapping PDF text can share a visual layer.
                        # Preserve it in the nearest free region on that row.
                        row, col, _, _ = grid.indices(box)
                        while (row, col) in grid.occupied:
                            col += 1
                        cell = sheet.cell(row, col)
                        grid.occupied.add((row, col))
                        warning = '部分重叠文字已保留在相邻单元格，需人工复核位置。'
                        if warning not in stats.warnings:
                            stats.warnings.append(warning)
                    _decorate_cell(cell, chars, box, lines, bordered)
                    stats.cells += 1
                    stats.text_characters += sum(len(_clean(c.get('text', ''))) for c in chars)
                _add_rules(grid, rules, stats)
                pdf_page = document[index]
                pictures = _PagePictures(pdf_page, page.width, page.height)
                try:
                    for box in image_boxes:
                        check_cancel()
                        if box[2] <= box[0] or box[3] <= box[1]:
                            continue
                        stream = pictures.crop(box)
                        streams.append(stream)
                        _add_picture(grid, stream, box)
                        stats.images += 1
                finally:
                    pictures.close()
                    pdf_page.close()
                sheet.print_area = f'A1:{get_column_letter(max(1, len(xs) - 1, sheet.max_column))}{max(1, len(ys) - 1, sheet.max_row)}'
                if any(not char.get('upright', True) for char in page.chars):
                    warning = '旋转文字已按文本保留，方向和字距需复核。'
                    if warning not in stats.warnings:
                        stats.warnings.append(warning)
                if page.curves or any(abs(line['x1'] - line['x0']) > 2 and abs(line['bottom'] - line['top']) > 2 for line in page.lines):
                    warning = '曲线、斜线和复杂矢量图形无法完全映射为 Excel 单元格，请使用原版保真模式核对。'
                    if warning not in stats.warnings:
                        stats.warnings.append(warning)
                if progress:
                    progress(index + 1, pages)
        check_cancel()
        workbook.save(generated)
    finally:
        document.close()
        workbook.close()
        for stream in streams:
            stream.close()
    return stats
