"""Recover editable Word structure instead of drawing one frame per PDF line.

Tagged PDFs retain paragraph identity, including identities crossing a page.
Use that evidence first. Untagged material uses conservative geometry, and
unsupported graphic regions remain local pictures. No Office automation, OCR,
network request, source modification, or legacy Word text frame is used here.
"""
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from math import hypot
from statistics import median
import re

from pdf_docx_layout import (_PagePictures, _anchor_picture, _small_anchor, _baseline, _element, _font_name,
                             _inside, _lines, _rgb_hex, find_ruled_tables)


@dataclass
class FlowStats:
    paragraphs: int = 0
    headings: int = 0
    tables: int = 0
    cells: int = 0
    images: int = 0
    scanned_pages: int = 0
    headers: int = 0
    footers: int = 0
    page_fields: int = 0
    toc_entries: int = 0
    code_blocks: int = 0
    tagged_paragraphs: int = 0
    cross_page_paragraphs: int = 0
    source_chars: int = 0
    consumed_chars: int = 0
    warnings: list = field(default_factory=list)


@dataclass
class Block:
    kind: str
    chars: list = field(default_factory=list)
    element: object = None
    page: int = 1
    box: tuple = None
    bookmark: str = ''
    level: int = 0
    entries: list = field(default_factory=list)
    parts: list = field(default_factory=list)
    tagged: bool = False
    bottom_rule: bool = False
    table_box: tuple = None

    @property
    def key(self):
        return self.page, self.box[1] if self.box else 0


def _clean(text):
    # XML 1.0 cannot store PDF's control glyphs. Keep tab/newline and all
    # Unicode letters, symbols, combining marks and supplementary characters.
    return ''.join(c for c in str(text) if c in '\t\n\r' or
                   (ord(c) >= 32 and not 0xd800 <= ord(c) <= 0xdfff and ord(c) not in (0xfffe, 0xffff)))


def _text(chars):
    return ''.join(_clean(c.get('text', '')) for c in chars)


def _normal(text):
    return re.sub(r'[\s.．·…、，,:：;；\-—（）()]+', '', text).casefold()


def _box(chars):
    if not chars:
        return None
    return min(c['x0'] for c in chars), min(c['top'] for c in chars), \
        max(c['x1'] for c in chars), max(c['bottom'] for c in chars)


def _page_lines(chars):
    grouped = defaultdict(list)
    for char in chars:
        grouped[char['page_number']].append(char)
    return [line for number in sorted(grouped) for line in _lines(grouped[number])]


def _roman(text):
    values = {'i': 1, 'v': 5, 'x': 10, 'l': 50, 'c': 100, 'd': 500, 'm': 1000}
    total, previous = 0, 0
    for c in text.lower()[::-1]:
        value = values[c]
        total += -value if value < previous else value
        previous = max(previous, value)
    return total


def _counter(text):
    """Return a page-number candidate only for narrow, explicit forms."""
    patterns = [r'^\s*[-—–]?\s*(\d{1,4}|[ivxlcdmIVXLCDM]{1,8})\s*[-—–]?\s*$',
                r'^\s*第\s*(\d{1,4})\s*页(?:\s*[/／共]\s*\d{1,4}\s*页?)?\s*$',
                r'^\s*Page\s+(\d{1,4})(?:\s+of\s+\d{1,4})?\s*$']
    for pattern in patterns:
        match = re.match(pattern, text, flags=re.I)
        if match:
            raw = match.group(1)
            value = int(raw) if raw.isdigit() else _roman(raw)
            return value, raw, text[:match.start(1)], text[match.end(1):]
    return None


class Source:
    """Index marked content once; MCID alone is not document-wide identity."""
    def __init__(self, extracted, check_cancel):
        self.pages = extracted.pages
        self.chars = {}
        self.objects = defaultdict(list)
        self.by_mcid = defaultdict(list)
        self.excluded = set()
        self.claimed = set()
        self.profiles = {}
        self.band_candidates = defaultdict(list)
        self.tagged = False
        for page in self.pages:
            check_cancel()
            self.chars[page.page_number] = page.chars
            for index, char in enumerate(page.chars):
                char['_flow_key'] = (page.page_number, index)
                if char.get('mcid') is not None:
                    self.by_mcid[page.page_number, char['mcid']].append(char)
            for objects in page.objects.values():
                for obj in objects:
                    if obj.get('mcid') is not None:
                        self.objects[page.page_number, obj['mcid']].append(obj)
        self._bands()
        try:
            from pdfplumber.structure import PDFStructTree, StructTreeMissing
            self.tree = PDFStructTree(extracted)
            self.tagged = bool(self.tree.children)
        except (ValueError, KeyError, AssertionError, TypeError, AttributeError):
            self.tree = None

    def _bands(self):
        repeated = defaultdict(list)
        numbers = defaultdict(list)
        for page in self.pages:
            for line in _lines(page.chars):
                text = _text(line).strip()
                box = _box(line)
                side = 'header' if box[3] < page.height * .075 else \
                    'footer' if box[1] > page.height * .91 else None
                if not side or not text or len(text) > 120:
                    continue
                candidate = (page.page_number, line, text, box)
                counter = _counter(text)
                if counter and box[2] - box[0] < page.width * .55:
                    value, raw, before, after = counter
                    # At least three agreeing pages are required. A body year,
                    # footnote marker, lone total, or isolated number stays text.
                    fmt = 'decimal' if raw.isdigit() else 'upperRoman' if raw.isupper() else 'lowerRoman'
                    numbers[side, value - page.page_number, before, after, fmt].append(candidate)
                else:
                    repeated[side, re.sub(r'\s+', '', text), round(box[1] / page.height, 2)].append(candidate)
        selected = []
        for key, items in numbers.items():
            if len(items) >= 3 and len({p for p, *_ in items}) >= 3:
                for item in items:
                    selected.append((key[0], item, {'before': key[2], 'after': key[3],
                                                   'number': item[0] + key[1], 'fmt': key[4]}))
        for key, items in repeated.items():
            if len(items) >= max(3, min(8, len(self.pages) // 4)):
                for item in items:
                    selected.append((key[0], item, None))
        by_page = defaultdict(lambda: {'header': [], 'footer': []})
        for side, (number, line, text, box), counter in selected:
            self.excluded.update(c['_flow_key'] for c in line)
            record = {'text': text, 'box': box, 'chars': line, 'counter': counter}
            by_page[number][side].append(record)
        for page in self.pages:
            profile = by_page[page.page_number]
            for side in profile:
                profile[side].sort(key=lambda item: (item['box'][1], item['box'][0]))
            self.profiles[page.page_number] = profile

    def refs(self, element):
        return [(p or element.page_number or 1, mcid) for p, mcid in element.all_mcids()]

    def element_chars(self, element, *, exclude_graphics=False):
        if exclude_graphics:
            skip = set()
            for child in _walk(element):
                if child.type in ('Figure', 'Formula'):
                    skip.update(self.refs(child))
        else:
            skip = set()
        result, seen = [], set()
        for ref in self.refs(element):
            if ref in skip:
                continue
            for char in self.by_mcid.get(ref, []):
                key = char['_flow_key']
                if key not in seen and key not in self.excluded:
                    result.append(char)
                    seen.add(key)
        return sorted(result, key=lambda c: (c['page_number'], _baseline(c), c['x0']))

    def element_box(self, element, number=None):
        number = number or element.page_number or 1
        page = self.pages[number - 1]
        attr = element.attributes.get('BBox')
        if attr and len(attr) == 4 and (element.page_number in (number, None)):
            left, bottom, right, top = map(float, attr)
            return (max(0, min(left, right)), max(0, page.height - max(bottom, top)),
                    min(page.width, max(left, right)), min(page.height, page.height - min(bottom, top)))
        objects = []
        for p, mcid in self.refs(element):
            if p == number:
                objects.extend(self.objects[p, mcid])
        return _box([obj for obj in objects if all(k in obj for k in ('x0', 'x1', 'top', 'bottom'))])

    def claim(self, chars):
        self.claimed.update(c['_flow_key'] for c in chars)

    def graphic(self, element):
        refs = self.refs(element)
        page = min((p for p, _ in refs), default=element.page_number or 1)
        box = self.element_box(element, page)
        chars = self.element_chars(element)
        self.claim(chars)
        return Block('graphic', chars, element, page, box, tagged=True)

    def blocks(self):
        blocks = []
        if self.tree:
            for child in self.tree.children:
                self._visit(child, blocks)
        # Preserve a picture-only page even when its only text is a page number.
        scan_pages = set()
        for page in self.pages:
            remaining_body = [c for c in page.chars if c['_flow_key'] not in self.excluded and _text([c]).strip()]
            if not remaining_body and (page.images or page.curves or page.rects):
                scan_pages.add(page.page_number)
                footer = self.profiles[page.page_number]['footer']
                bottom = min((item['box'][1] - 3 for item in footer), default=page.height)
                blocks.append(Block('scan', page=page.page_number, box=(0, 0, page.width, bottom)))
            elif not remaining_body:
                blocks.append(Block('blank', page=page.page_number))
        blocks = [b for b in blocks if b.kind == 'scan' or b.page not in scan_pages]
        for page in self.pages:
            if page.page_number in scan_pages:
                continue
            # Some producers leave objects outside the structure tree. Retain
            # their visible text, rather than trusting the presence of tags.
            remaining = [c for c in page.chars if c['_flow_key'] not in self.claimed
                         and c['_flow_key'] not in self.excluded]
            ruled = []
            if not self.tagged:
                ruled = find_ruled_tables(page)
                for table in ruled:
                    chars = [c for c in remaining if _inside(c, table.bbox)]
                    blocks.append(Block('ruled_table', chars, table, page.page_number, table.bbox))
                    self.claim(chars)
                remaining = [c for c in remaining if c['_flow_key'] not in self.claimed]
            if remaining:
                blocks.extend(_geometry_blocks(remaining))
                self.claim(remaining)
            covered = []
            for block in blocks:
                if block.page == page.page_number and block.kind in ('graphic', 'scan', 'floating_photo') and block.box:
                    covered.append(block.box)
                covered.extend(g.box for g in block.parts if isinstance(g, Block)
                               and g.page == page.page_number and g.box)
                if block.kind == 'table':
                    for figure in _walk(block.element):
                        if figure.type == 'Figure' and (figure.page_number or block.page) == page.page_number:
                            region = self.element_box(figure, page.page_number)
                            if region: covered.append(region)
            table_boxes = [t.bbox for t in ruled] + [b.box for b in blocks if b.kind == 'table'
                                                     and b.page == page.page_number and b.box]
            for image in page.images:
                box = (max(0, image['x0']), max(0, image['top']),
                       min(page.width, image['x1']), min(page.height, image['bottom']))
                area = (box[2] - box[0]) * (box[3] - box[1])
                if area <= .01:
                    continue
                if any(max(0, min(box[2], region[2]) - max(box[0], region[0])) *
                       max(0, min(box[3], region[3]) - max(box[1], region[1])) >= area * .97 for region in covered):
                    continue
                # A tagged PDF may still contain pictures outside its tree.
                # Retain them unless an emitted graphic already covers them.
                overlap = [region for region in table_boxes if _inside(
                    {'x0': box[0], 'x1': box[2], 'top': box[1], 'bottom': box[3]}, region)]
                graphic = Block('floating_photo' if overlap else 'graphic', page=page.page_number, box=box)
                if overlap:
                    owner = min(overlap, key=lambda region: (region[2] - region[0]) * (region[3] - region[1]))
                    graphic.table_box = owner
                    graphic.parts = [owner[1] - .01]
                blocks.append(graphic)
                covered.append(box)
            if not self.tagged:
                for rule in list(page.lines) + list(page.rects):
                    if (rule.get('width', 0) < page.width * .3 or rule.get('height', 0) > 2 or
                        any(table.bbox[1] - 1 <= rule['top'] <= table.bbox[3] + 1 for table in ruled)):
                        continue
                    color = rule.get('stroking_color') if rule.get('stroke') else rule.get('non_stroking_color')
                    if _rgb_hex(color) == 'FFFFFF':
                        continue
                    candidates = [b for b in blocks if b.page == page.page_number and b.kind == 'P'
                                  and b.box and b.box[1] <= rule['top'] + 3]
                    if candidates:
                        min(candidates, key=lambda b: abs(b.box[3] - rule['top'])).bottom_rule = True
        ordered = sorted(blocks, key=lambda b: (b.page, b.parts[0])
                         if b.kind == 'floating_photo' and b.parts else b.key)
        merged = []
        for block in ordered:
            previous = merged[-1] if merged else None
            adjacent = False
            if previous and previous.chars and block.chars:
                last_page = max(c['page_number'] for c in previous.chars)
                last_box = _box([c for c in previous.chars if c['page_number'] == last_page])
                first_page = min(c['page_number'] for c in block.chars)
                first_box = _box([c for c in block.chars if c['page_number'] == first_page])
                adjacent = ((first_page == last_page and -.5 <= first_box[1] - last_box[3] <= 24)
                            or (first_page == last_page + 1 and first_box[1] < self.pages[first_page - 1].height * .25))
            if (previous and previous.kind in ('P', 'Code') and block.kind in ('P', 'Code')
                    and adjacent and (_is_code(previous) or _code_seed(previous)) and _code_continuation(block)):
                previous.chars.extend(block.chars)
                previous.kind = 'Code'
                previous.box = _box(previous.chars)
            else:
                merged.append(block)
        return merged

    def _visit(self, element, blocks):
        kind = element.type
        if kind == 'TOC':
            entries = []
            for item in _walk(element):
                if item.type != 'TOCI':
                    continue
                chars = self.element_chars(item)
                # Fill in unmarked leader dots and page labels on the same
                # physical lines. MCID-only extraction silently loses these.
                full = []
                for line in _page_lines(chars):
                    p = line[0]['page_number']
                    base = median(_baseline(c) for c in line)
                    size = median(c.get('size', 10) for c in line)
                    full.extend(c for c in self.chars[p]
                                if c['_flow_key'] not in self.excluded
                                and abs(_baseline(c) - base) < max(.8, size * .1))
                full = list({c['_flow_key']: c for c in full or chars}.values())
                full.sort(key=lambda c: (c['page_number'], _baseline(c), c['x0']))
                self.claim(full)
                if full:
                    entries.append(full)
            chars = self.element_chars(element)
            self.claim(chars)
            if entries:
                all_chars = [c for entry in entries for c in entry]
                page = min(c['page_number'] for c in all_chars)
                blocks.append(Block('toc', all_chars, element, page,
                                    _box([c for c in all_chars if c['page_number'] == page]),
                                    entries=entries, tagged=True))
            return
        if kind == 'Table':
            chars = self.element_chars(element)
            self.claim(chars)
            page = min((c['page_number'] for c in chars), default=element.page_number or 1)
            blocks.append(Block('table', chars, element, page,
                                self.element_box(element, page), tagged=True))
            return
        if kind in ('Figure', 'Formula'):
            graphic = self.graphic(element)
            if graphic.box:
                blocks.append(graphic)
            return
        if kind in ('P', 'Caption', 'Code', 'H', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6'):
            chars = self.element_chars(element, exclude_graphics=True)
            nested_tables = [child for child in _walk(element) if child is not element and child.type == 'Table']
            table_nodes = {id(node) for table in nested_tables for node in _walk(table)}
            table_refs = {ref for table in nested_tables for ref in self.refs(table)}
            chars = [char for char in chars if (char['page_number'], char.get('mcid')) not in table_refs]
            graphics = [self.graphic(child) for child in _walk(element)
                        if child is not element and id(child) not in table_nodes
                        and child.type in ('Figure', 'Formula')]
            # A paragraph containing just a figure should not become an empty
            # paragraph plus the same figure emitted a second time.
            if chars:
                self.claim(chars)
                page = min(c['page_number'] for c in chars)
                block = Block(kind, chars, element, page,
                              _box([c for c in chars if c['page_number'] == page]),
                              parts=[g for g in graphics if g.box], tagged=True)
                blocks.append(block)
            else:
                blocks.extend(g for g in graphics if g.box)
            # Text containers may contain a nested floating-frame Div→Table.
            # Retain a real table rather than flattening all its cells into P.
            nested_children = {id(child) for table in nested_tables for child in _walk(table)
                               if child is not table and child.type == 'Table'}
            for table in nested_tables:
                if id(table) not in nested_children:
                    self._visit(table, blocks)
            return
        for child in element.children:
            self._visit(child, blocks)


def _walk(element):
    yield element
    for child in element.children:
        yield from _walk(child)


def _geometry_blocks(chars):
    """Conservative paragraph fallback, independent of Word frame geometry."""
    result, current = [], []
    for line in _page_lines(chars):
        text = _text(line).strip()
        if not text:
            continue
        previous = current[-1] if current else None
        new = previous is None
        if previous:
            size = median(c.get('size', 10) for c in line)
            oldsize = median(c.get('size', 10) for c in previous)
            delta = min(c['top'] for c in line) - min(c['top'] for c in previous)
            indent = min(c['x0'] for c in line) - min(c['x0'] for c in previous)
            oldtext = _text(previous).rstrip()
            new = (line[0]['page_number'] != previous[0]['page_number'] or
                   delta > max(size, oldsize) * 1.85 or delta < size * .65 or
                   abs(size - oldsize) > 1.5 or abs(indent) > size * 2.7 or
                   (oldtext.endswith(('。', '！', '？', ':', '：')) and indent > size * 1.4))
        if new and current:
            flat = [c for group in current for c in group]
            result.append(Block('P', flat, page=flat[0]['page_number'], box=_box(flat)))
            current = []
        current.append(line)
    if current:
        flat = [c for group in current for c in group]
        result.append(Block('P', flat, page=flat[0]['page_number'], box=_box(flat)))
    return result


def _font_key(char):
    name = char.get('fontname', 'SimSun')
    matrix = char.get('matrix')
    ratio = 1
    if matrix and len(matrix) == 6 and hypot(matrix[2], matrix[3]) > 0:
        ratio = hypot(matrix[0], matrix[1]) / hypot(matrix[2], matrix[3])
    return (_font_name(name), round(max(4, min(96, char.get('size', 10))) * 2) / 2,
            bool(char.get('_ruyi_bold')) or any(t in name.lower() for t in ('bold', 'black')),
            any(t in name.lower() for t in ('italic', 'oblique')),
            _rgb_hex(char.get('non_stroking_color')), max(10, min(600, round(ratio * 100))))


def _run(paragraph, text, char, code=False):
    from docx.oxml.ns import qn
    from docx.shared import Pt, RGBColor
    text = _clean(text)
    if not text:
        return
    run = paragraph.add_run(text)
    name, size, bold, italic, color, ratio = _font_key(char)
    run.font.name, run.font.size = name, Pt(size)
    run.font.bold, run.font.italic = bold, italic
    run.font.color.rgb = RGBColor.from_string(color)
    fonts = run._r.get_or_add_rPr().find(qn('w:rFonts'))
    fonts.set(qn('w:eastAsia'), name)
    fonts.set(qn('w:cs'), name)
    if ratio != 100 and not code:
        run._r.get_or_add_rPr().append(_element('w:w', val=ratio))
    return run


def _is_code(block):
    if block.kind == 'Code':
        return True
    # Spaces belong to syntax: deleting them turns "def hello" into
    # "defhello" and makes ordinary-font source code fail recognition.
    chars = [c for c in block.chars if c.get('text', '')]
    if not chars:
        return False
    mono = sum(any(name in c.get('fontname', '').lower()
                   for name in ('courier', 'consol', 'mono', 'menlo')) for c in chars) / len(chars)
    text = _text(chars)
    lines = [_text(line).strip() for line in _page_lines(chars)]
    patterns = [r'^(?:async\s+)?(?:def|class)\s+[A-Za-z_]\w*\s*(?:\([^)]*\))?\s*:',
                r'^(?:from\s+[\w.]+\s+import\s+|import\s+[A-Za-z_])',
                r'^(?:for\s+.+\s+in\s+.+|if\s+.+|elif\s+.+|while\s+.+|with\s+.+)\s*:\s*$',
                r'^(?:return\s+.+|try\s*:|except(?:\s+[^:]+)?\s*:|else\s*:)',
                r'^[A-Za-z_][\w.\[\]\'\"]*\s*=\s*[^=]',
                r'^(?:[A-Za-z_]\w*\.)*[A-Za-z_]\w*\s*\(.*\)\s*;?\s*$']
    evidence = [any(re.search(pattern, line) for pattern in patterns) for line in lines]
    strong = sum(evidence)
    has_definition_or_import = any(re.search(patterns[0], line) or re.search(patterns[1], line)
                                   for line in lines)
    lexical = strong > 0
    # Font is supporting evidence, not a gate: many copied programs use Times
    # New Roman plus SimSun. Multiple anchored syntax lines distinguish a
    # program from prose merely mentioning Python, requests or an assignment.
    return ((mono >= .65 and (lexical or len(lines) >= 2)) or
            (strong >= 3 and strong / max(1, len(lines)) >= .25) or
            (has_definition_or_import and strong >= 2))


def _code_continuation(block):
    if _is_code(block):
        return True
    lines = [_text(line).strip() for line in _page_lines(block.chars)]
    return bool(lines) and all(re.match(r'^(?:#|pass\b|return\b|break\b|continue\b|except\b|else\s*:|[}\])]|'
                                       r'(?:def|class|import|from|for|if|elif|while|with)\s|'
                                       r'[A-Za-z_][\w.\[\]\'\"]*\s*=\s*[^=]|[\w.]+\s*\()',
                                       line, flags=re.I) for line in lines if line)


def _code_seed(block):
    lines = [_text(line).strip() for line in _page_lines(block.chars)]
    return bool(lines) and all(re.match(r'^(?:import\s+[\w.]|from\s+[\w.]+\s+import\b|'
                                       r'(?:async\s+)?def\s+\w+\s*\(.*\)\s*:|class\s+\w+.*:)', line)
                              for line in lines if line)


def _line_fragments(line, *, code=False):
    fragments, previous = [], None
    for char in line:
        text = _clean(char.get('text', ''))
        if not text:
            continue
        if previous and not previous.get('text', '').endswith((' ', '\t')) and not text.startswith((' ', '\t')):
            gap = char['x0'] - previous['x1']
            size = max(4, char.get('size', 10))
            if code and gap > size * .4:
                fragments.append((' ' * min(48, max(1, round(gap / (size * .6)))), char))
            elif gap > max(1.5, size * .24):
                fragments.append((' ', char))
        fragments.append((text, char))
        previous = char
    # Physical line-end spaces do not create artificial word breaks in CJK.
    while fragments and fragments[-1][0].isspace():
        fragments.pop()
    return fragments


def _write_text(paragraph, chars, *, code=False, preserve_lines=False,
                graphics=None, picture_callback=None, orphan_picture_callback=None):
    lines = _page_lines(chars)
    placements = defaultdict(list)
    orphans = []
    for graphic in graphics or []:
        candidates = [(max(0, max(graphic.box[1], min(c['top'] for c in line)) -
                               min(graphic.box[3], max(c['bottom'] for c in line))), index)
                      for index, line in enumerate(lines) if line[0]['page_number'] == graphic.page]
        if candidates:
            placements[min(candidates)[1]].append(graphic)
        else:
            orphans.append(graphic)
    pending, style, sample = '', None, None
    last_text = ''
    left = min((c['x0'] for c in chars), default=0)
    for number, line in enumerate(lines):
        fragments = _line_fragments(line, code=code)
        for graphic in sorted(placements[number], key=lambda g: g.box[0], reverse=True):
            offset = next((j for j, (_, c) in enumerate(fragments)
                           if not isinstance(c, Block) and c['x0'] >= graphic.box[0]), len(fragments))
            fragments.insert(offset, ('', graphic))
        if not fragments:
            continue
        if number:
            if code or preserve_lines:
                if pending:
                    _run(paragraph, pending, sample, code)
                    pending = ''
                paragraph.add_run().add_break()
                if code:
                    size = median(c.get('size', 10) for c in line)
                    indent = max(0, round((min(c['x0'] for c in line) - left) / max(2.4, size * .6)))
                    if indent:
                        _run(paragraph, ' ' * min(100, indent), line[0], True)
            else:
                first = next((text.lstrip() for text, char in fragments if not isinstance(char, Block)), '')
                a, b = last_text[-1:] if last_text else '', first[:1]
                # An English word at a PDF wrap is separated; CJK wrapping is
                # joined directly. Existing spaces are not duplicated.
                if a and b and not a.isspace() and not b.isspace() and \
                        (a.isascii() and b.isascii()) and (a.isalnum() or a in ',;:') and b.isalnum():
                    fragments.insert(0, (' ', next(c for _, c in fragments if not isinstance(c, Block))))
        for text, char in fragments:
            if isinstance(char, Block):
                if pending:
                    _run(paragraph, pending, sample, code)
                    pending = ''
                if picture_callback:
                    picture_callback(char)
                style = None
                continue
            newstyle = _font_key(char)
            if style != newstyle and pending:
                _run(paragraph, pending, sample, code)
                pending = ''
            style, sample = newstyle, char
            pending += text
            last_text = text
    if pending:
        _run(paragraph, pending, sample, code)
    for graphic in orphans:
        callback = orphan_picture_callback or picture_callback
        if callback: callback(graphic)


def _set_paragraph(paragraph, block, margins, *, cell=False, code=False):
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt
    pf = paragraph.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0 if cell else 3)
    pf.keep_together = False
    pf.widow_control = True
    props = paragraph._p.get_or_add_pPr()
    props.append(_element('w:autoSpaceDE', val=0))
    props.append(_element('w:autoSpaceDN', val=0))
    attrs = block.element.attributes if block.element else {}
    align = attrs.get('TextAlign')
    alignment = {'Center': WD_ALIGN_PARAGRAPH.CENTER, 'End': WD_ALIGN_PARAGRAPH.RIGHT,
                 'Right': WD_ALIGN_PARAGRAPH.RIGHT, 'Justify': WD_ALIGN_PARAGRAPH.JUSTIFY,
                 'Start': WD_ALIGN_PARAGRAPH.LEFT, 'Left': WD_ALIGN_PARAGRAPH.LEFT}
    if align in alignment:
        paragraph.alignment = alignment[align]
    lines = _page_lines(block.chars)
    sizes = [c.get('size', 10) for c in block.chars if c.get('text', '').strip()]
    size = median(sizes) if sizes else 10
    spaces = [median(_baseline(c) for c in b) - median(_baseline(c) for c in a)
              for a, b in zip(lines, lines[1:]) if a[0]['page_number'] == b[0]['page_number']]
    spaces = [v for v in spaces if size * .9 <= v <= size * 2.6]
    pf.line_spacing = Pt(median(spaces)) if spaces else 1.25
    if not cell and lines:
        restleft = median(min(c['x0'] for c in line) for line in lines[1:]) if len(lines) > 1 \
            else min(c['x0'] for c in lines[0])
        firstleft = min(c['x0'] for c in lines[0])
        # Indentation is a paragraph property, never a floating text frame.
        pf.left_indent = Pt(max(0, min(100, restleft - margins[0])))
        pf.first_line_indent = Pt(max(-48, min(72, firstleft - restleft)))
        if not align:
            textleft = min(c['x0'] for c in block.chars)
            textright = max(c['x1'] for c in block.chars)
            pagewidth = margins[2]
            if len(lines) == 1 and abs((textleft + textright) / 2 - pagewidth / 2) < 8:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                pf.left_indent = pf.first_line_indent = Pt(0)
    if code:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf.first_line_indent = Pt(0)
        pf.line_spacing = 1.15
        pf.space_after = Pt(5)
    if block.bottom_rule:
        borders = _element('w:pBdr')
        borders.append(_element('w:bottom', val='single', sz=6, space=2, color='000000'))
        props.append(borders)


def _add_bookmark(paragraph, name, number):
    start = _element('w:bookmarkStart', id=number, name=name)
    end = _element('w:bookmarkEnd', id=number)
    ppr = paragraph._p.pPr
    paragraph._p.insert(1 if ppr is not None else 0, start)
    paragraph._p.append(end)


def _field(paragraph, instruction, cached='1'):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    begin = OxmlElement('w:fldChar'); begin.set(qn('w:fldCharType'), 'begin')
    paragraph.add_run()._r.append(begin)
    text = OxmlElement('w:instrText'); text.set(qn('xml:space'), 'preserve'); text.text = ' ' + instruction + ' '
    paragraph.add_run()._r.append(text)
    separate = OxmlElement('w:fldChar'); separate.set(qn('w:fldCharType'), 'separate')
    paragraph.add_run()._r.append(separate)
    paragraph.add_run(cached)
    end = OxmlElement('w:fldChar'); end.set(qn('w:fldCharType'), 'end')
    paragraph.add_run()._r.append(end)


def _profile_key(profile):
    return tuple((side, tuple((item['counter']['before'], item['counter']['after'], item['counter']['fmt'])
                            if item['counter'] else item['text'] for item in profile[side]))
                 for side in ('header', 'footer'))


def _configure_section(section, page, margins):
    from docx.shared import Pt
    scale = min(1, 22 * 72 / max(page.width, page.height))
    section.page_width, section.page_height = Pt(page.width * scale), Pt(page.height * scale)
    section.left_margin, section.right_margin = Pt(margins[0]), Pt(margins[1])
    section.top_margin, section.bottom_margin = Pt(64), Pt(70)
    section.header_distance, section.footer_distance = Pt(30), Pt(42)


def _write_bands(section, profile, stats):
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.shared import Pt
    first_counter = None
    for side in ('header', 'footer'):
        container = getattr(section, side)
        container.is_linked_to_previous = False
        for node in list(container._element):
            container._element.remove(node)
        items = profile[side]
        if not items:
            container.add_paragraph()
        for item in items:
            p = container.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1
            counter = item['counter']
            if counter:
                p.add_run(counter['before'])
                suffix = ' \\* ' + ('ROMAN' if counter['fmt'] == 'upperRoman' else 'roman') \
                    if counter['fmt'] != 'decimal' else ''
                raw = _counter(item['text'])
                _field(p, 'PAGE' + suffix, raw[1] if raw else str(counter['number']))
                total = re.match(r'^(.*?)(\d+)(.*?)$', counter['after'])
                if total:
                    p.add_run(total.group(1))
                    _field(p, 'NUMPAGES', total.group(2))
                    p.add_run(total.group(3))
                else:
                    p.add_run(counter['after'])
                stats.page_fields += 1
                first_counter = first_counter or counter
            else:
                _write_text(p, item['chars'])
            if item['chars']:
                size = median(c.get('size', 9) for c in item['chars'])
                name = _font_name(item['chars'][0].get('fontname', 'SimSun'))
                for run in p.runs:
                    run.font.size, run.font.name = Pt(size), name
            if side == 'header': stats.headers += 1
            else: stats.footers += 1
    if first_counter:
        props = section._sectPr
        old = props.find(qn('w:pgNumType'))
        if old is not None:
            props.remove(old)
        props.append(_element('w:pgNumType', start=first_counter['number'], fmt=first_counter['fmt']))


def _margins(blocks, page):
    lefts, rights = [], []
    for block in blocks:
        if block.kind not in ('P', 'Caption'):
            continue
        for line in _page_lines(block.chars):
            if len(line) >= 25:
                lefts.append(min(c['x0'] for c in line))
                rights.append(page.width - max(c['x1'] for c in line))
    left = max(36, min(90, median(lefts))) if lefts else 54
    right = max(36, min(90, median(rights))) if rights else 54
    return left, right, page.width


def _picture(document, block, source, pdf, margins, stats, *, paragraph=None):
    from docx.shared import Pt
    page = source.pages[block.page - 1]
    box = block.box
    if not box or box[2] - box[0] < .1 or box[3] - box[1] < .1:
        return
    pdf_page = pdf[block.page - 1]
    pictures = _PagePictures(pdf_page, page.width, page.height)
    try:
        stream = pictures.crop(box)
        width, height = box[2] - box[0], box[3] - box[1]
        maxwidth = page.width - margins[0] - margins[1]
        maxheight = page.height - 64 - 70 - 12
        factor = min(1, maxwidth / width, maxheight / height)
        if block.kind == 'floating_photo':
            p = paragraph if paragraph is not None else document.add_paragraph()
            if paragraph is None:
                _small_anchor(p)
            scale = min(1, 22 * 72 / max(page.width, page.height))
            _anchor_picture(p, stream, box, scale)
            if block.table_box:
                _align_table_photo(p._p.xpath('.//wp:anchor')[-1], box, block.table_box, scale)
            stats.images += 1
            return
        p = paragraph or document.add_paragraph()
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        if paragraph is None:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(4)
            p.paragraph_format.keep_together = True
        p.add_run().add_picture(stream, width=Pt(width * factor), height=Pt(height * factor))
        stats.images += 1
        if block.kind == 'scan':
            stats.scanned_pages += 1
    finally:
        pictures.close()
        pdf_page.close()


def _align_table_photo(anchor, box, table_box, scale=1):
    """Keep a body's free photo aligned with its adjacent inline table.

    The table follows a one-point body anchor paragraph. Its page position can
    change during flow layout, so copying PDF page x/y would separate the photo
    from its cell. Column/paragraph coordinates preserve the source offset from
    the table without putting the picture inside a cell or text frame.
    """
    from docx.oxml.ns import qn
    from docx.shared import Pt
    for axis, reference, distance in (
            ('H', 'column', (box[0] - table_box[0]) * scale),
            ('V', 'paragraph', 1 + (box[1] - table_box[1]) * scale)):
        position = anchor.find(qn('wp:position' + axis))
        position.set('relativeFrom', reference)
        position.find(qn('wp:posOffset')).text = str(int(Pt(distance)))


def _semantic_table(document, block, source, pdf, margins, stats):
    from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
    from docx.shared import Pt
    from docx.oxml.ns import qn
    rows = [node for node in _walk(block.element) if node.type == 'TR']
    cells = [[child for child in row.children if child.type in ('TD', 'TH')] for row in rows]
    cells = [row for row in cells if row]
    if not cells:
        _picture(document, block, source, pdf, margins, stats)
        return
    table_figures = [node for node in _walk(block.element) if node.type == 'Figure']
    photo_paragraph = document.add_paragraph() if table_figures else None
    if photo_paragraph is not None:
        _small_anchor(photo_paragraph)
    for figure in table_figures:
        photo = source.graphic(figure)
        photo.kind = 'floating_photo'
        photo.table_box = block.box
        _picture(document, photo, source, pdf, margins, stats, paragraph=photo_paragraph)
    def span(cell, name):
        try: return max(1, min(100, int(cell.attributes.get(name, 1))))
        except (TypeError, ValueError): return 1
    columns = max(sum(span(c, 'ColSpan') for c in row) for row in cells)
    if columns * len(cells) > 15000:
        _picture(document, block, source, pdf, margins, stats)
        stats.warnings.append('有过于复杂的表格按局部图片保留。')
        return
    table = document.add_table(rows=len(cells), cols=columns)
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'
    available = margins[2] - margins[0] - margins[1]
    widths = [available / columns] * columns
    # Tagged cell widths are used as relative proportions, avoiding any
    # producer-specific attribute unit assumptions.
    for row in cells:
        if len(row) == columns and all(span(c, 'ColSpan') == 1 for c in row):
            raw = [float(c.attributes.get('Width', 0) or 0) for c in row]
            if min(raw, default=0) > 0:
                widths = [available * w / sum(raw) for w in raw]
                break
    for column, width in zip(table.columns, widths):
        column.width = Pt(width)
    occupied = set()
    for row_index, row in enumerate(cells):
        col = 0
        for tagged_cell in row:
            while (row_index, col) in occupied and col < columns:
                col += 1
            if col >= columns:
                break
            col_span, row_span = min(span(tagged_cell, 'ColSpan'), columns - col), \
                min(span(tagged_cell, 'RowSpan'), len(cells) - row_index)
            cell = table.cell(row_index, col)
            if row_span > 1 or col_span > 1:
                cell = cell.merge(table.cell(row_index + row_span - 1, col + col_span - 1))
            for r in range(row_index, row_index + row_span):
                for c in range(col, col + col_span): occupied.add((r, c))
            cell.width = Pt(sum(widths[col:col + col_span]))
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            content = [node for node in _walk(tagged_cell)
                       if node.type in ('P', 'Caption', 'Code')]
            if not content:
                content = [tagged_cell]
            used = False
            processed_formulas = set()
            for child in content:
                chars = source.element_chars(child, exclude_graphics=True)
                formulas = [source.graphic(node) for node in _walk(child) if node.type == 'Formula']
                formulas = [graphic for graphic in formulas if graphic.box]
                processed_formulas.update(id(graphic.element) for graphic in formulas)
                if not chars and not formulas:
                    continue
                p = cell.paragraphs[0] if not used else cell.add_paragraph()
                used = True
                b = Block(child.type, chars, child, tagged=True)
                _set_paragraph(p, b, margins, cell=True, code=_is_code(b))
                if chars:
                    _write_text(p, chars, code=_is_code(b), graphics=formulas,
                                picture_callback=lambda graphic: _picture(document, graphic, source, pdf, margins, stats,
                                                                         paragraph=p),
                                orphan_picture_callback=lambda graphic: _picture(document, graphic, source, pdf, margins, stats,
                                                                                paragraph=cell.add_paragraph()))
                else:
                    for graphic in formulas:
                        _picture(document, graphic, source, pdf, margins, stats, paragraph=p)
                stats.paragraphs += 1
            for graphic in _walk(tagged_cell):
                if graphic.type == 'Formula' and id(graphic) not in processed_formulas:
                    # Formulas belong to cell text; photographs were anchored
                    # in a body paragraph before the table and remain free.
                    _picture(document, source.graphic(graphic), source, pdf, margins, stats,
                             paragraph=cell.add_paragraph())
            tcpr = cell._tc.get_or_add_tcPr()
            mar = _element('w:tcMar')
            for side in ('top', 'bottom'): mar.append(_element('w:' + side, w=25, type='dxa'))
            for side in ('left', 'right'): mar.append(_element('w:' + side, w=35, type='dxa'))
            tcpr.append(mar)
            stats.cells += 1
            col += col_span
    # Do not use fixed exact heights: user edits and font fallback must be able
    # to grow a row. Table is inline, with no tblpPr positioning container.
    stats.tables += 1
    document.add_paragraph().paragraph_format.space_after = Pt(3)


def _toc_data(block):
    entries = []
    for chars in block.entries:
        lines = _page_lines(chars)
        text = ''.join(_text(line).strip() for line in lines)
        match = re.match(r'^(.*?)[.．·…]{2,}\s*(\d+|[ivxlcdmIVXLCDM]+)\s*$', text)
        if match:
            title, number = match.groups()
        else:
            match = re.match(r'^(.*?)\s*(\d+)\s*$', text)
            title, number = match.groups() if match and match.group(1).strip() else (text, '')
        title = re.sub(r'[.．·…]+$', '', title).strip()
        if title:
            entries.append((title, number, chars))
    return entries


def _toc(document, block, bookmarks, margins, stats):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.enum.text import WD_TAB_ALIGNMENT, WD_TAB_LEADER
    from docx.shared import Pt
    entries = _toc_data(block)
    figure = sum(bool(re.match(r'^图\s*\d+', title)) for title, _, _ in entries)
    table = sum(bool(re.match(r'^表\s*\d+', title)) for title, _, _ in entries)
    instruction = 'TOC \\h \\z \\t "Ruyi Figure Caption,1"' if figure > len(entries) * .6 else \
        'TOC \\h \\z \\t "Ruyi Table Caption,1"' if table > len(entries) * .6 else \
        'TOC \\o "1-6" \\h \\z \\u'
    start = document.add_paragraph()
    start.paragraph_format.space_after = Pt(0)
    start.paragraph_format.line_spacing = Pt(1)
    begin = OxmlElement('w:fldChar'); begin.set(qn('w:fldCharType'), 'begin')
    start.add_run()._r.append(begin)
    instr = OxmlElement('w:instrText'); instr.set(qn('xml:space'), 'preserve'); instr.text = ' ' + instruction + ' '
    start.add_run()._r.append(instr)
    separate = OxmlElement('w:fldChar'); separate.set(qn('w:fldCharType'), 'separate')
    start.add_run()._r.append(separate)
    for title, number, chars in entries:
        target = bookmarks.get(_normal(title))
        level = target[1] if target else (2 if re.match(r'^[（(]', title) else 3 if re.match(r'^\d+[.．、]', title) else 1)
        level = max(1, min(6, level))
        p = document.add_paragraph(style='TOC ' + str(level))
        pf = p.paragraph_format
        pf.space_after = Pt(0)
        pf.line_spacing = 1.2
        pf.tab_stops.add_tab_stop(Pt(margins[2] - margins[0] - margins[1]),
                                 WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
        if target:
            hyperlink = OxmlElement('w:hyperlink'); hyperlink.set(qn('w:anchor'), target[0])
            run = OxmlElement('w:r'); text = OxmlElement('w:t'); text.text = _clean(title)
            run.append(text); hyperlink.append(run); p._p.append(hyperlink)
        else:
            p.add_run(title)
        p.add_run('\t')
        if target and number:
            _field(p, 'PAGEREF ' + target[0] + ' \\h', number)
        else:
            p.add_run(number)
        stats.toc_entries += 1
    end = OxmlElement('w:fldChar'); end.set(qn('w:fldCharType'), 'end')
    document.paragraphs[-1].add_run()._r.append(end)


def _styles(document):
    from docx.enum.style import WD_STYLE_TYPE
    from docx.shared import Pt, RGBColor
    from docx.oxml.ns import qn
    normal = document.styles['Normal']
    normal.font.name, normal.font.size = '宋体', Pt(12)
    normal._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), '宋体')
    normal.paragraph_format.space_after = Pt(3)
    normal.paragraph_format.line_spacing = 1.3
    for level in range(1, 7):
        style = document.styles['Heading ' + str(level)]
        style.font.color.rgb = RGBColor(0, 0, 0)
        style.paragraph_format.keep_with_next = True
        name = 'TOC ' + str(level)
        if name not in document.styles:
            toc = document.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
            toc.base_style = normal
            toc.paragraph_format.left_indent = Pt((level - 1) * 12)
    for name in ('Ruyi Figure Caption', 'Ruyi Table Caption', 'Ruyi Code'):
        if name not in document.styles:
            style = document.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
            style.base_style = normal


def rebuild_pdf_docx_flow(source, generated, pages, check_cancel, progress=None):
    """Create native continuous paragraphs, semantic tables/TOC/page fields."""
    import pdfplumber
    import pypdfium2 as pdfium
    from docx import Document
    from docx.enum.section import WD_SECTION
    from docx.shared import Pt
    stats = FlowStats()
    document = Document()
    _styles(document)
    pdf = pdfium.PdfDocument(source)
    try:
        with pdfplumber.open(source) as extracted:
            check_cancel()
            indexed = Source(extracted, check_cancel)
            blocks = indexed.blocks()
            stats.source_chars = sum(len(page.chars) for page in extracted.pages)
            stats.consumed_chars = len(indexed.claimed | indexed.excluded)
            margins = _margins(blocks, extracted.pages[0])
            bookmarks = {}
            for i, block in enumerate(blocks):
                if re.fullmatch(r'H[1-6]', block.kind):
                    block.level = int(block.kind[1:])
                elif block.kind == 'Caption':
                    block.level = 1
                if block.level:
                    block.bookmark = 'Ruyi_' + str(i + 1)
                    bookmarks.setdefault(_normal(_text(block.chars)), (block.bookmark, block.level))
            if not indexed.tagged:
                stats.warnings.append('此 PDF 未提供语义标签，段落采用几何推断；未确定的目录和标题保持普通正文。')
            if any(block.kind == 'toc' for block in blocks):
                stats.warnings.append('目录已重建为可更新域，当前保留原稿目录缓存；请在 Word 中更新整个目录以匹配新的分页。')
            section = document.sections[0]
            profile = indexed.profiles[1]
            current_key = _profile_key(profile)
            _configure_section(section, extracted.pages[0], margins)
            _write_bands(section, profile, stats)
            previous_kind, previous_page = None, 0
            reported = set()
            table_photo_anchors = {}
            for i, block in enumerate(blocks):
                check_cancel()
                wanted = indexed.profiles[block.page]
                key = _profile_key(wanted)
                if key != current_key:
                    section = document.add_section(WD_SECTION.NEW_PAGE)
                    _configure_section(section, extracted.pages[block.page - 1], margins)
                    _write_bands(section, wanted, stats)
                    current_key = key
                if block.kind == 'toc':
                    _toc(document, block, bookmarks, margins, stats)
                elif block.kind in ('graphic', 'scan', 'floating_photo'):
                    if block.kind == 'scan' and previous_kind and previous_page != block.page:
                        document.add_page_break()
                    photo_paragraph = None
                    if block.kind == 'floating_photo' and block.table_box:
                        identity = (block.page, block.table_box)
                        if identity not in table_photo_anchors:
                            table_photo_anchors[identity] = document.add_paragraph()
                            _small_anchor(table_photo_anchors[identity])
                        photo_paragraph = table_photo_anchors[identity]
                    _picture(document, block, indexed, pdf, margins, stats, paragraph=photo_paragraph)
                    if block.kind == 'scan' and i + 1 < len(blocks):
                        document.add_page_break()
                elif block.kind == 'blank':
                    if previous_kind and previous_kind not in ('scan', 'blank'):
                        document.add_page_break()
                    document.add_paragraph()
                    if i + 1 < len(blocks): document.add_page_break()
                elif block.kind == 'table':
                    _semantic_table(document, block, indexed, pdf, margins, stats)
                elif block.kind == 'ruled_table':
                    from pdf_docx_layout import _add_table
                    from docx.oxml.ns import qn
                    from docx.enum.table import WD_ROW_HEIGHT_RULE
                    _add_table(document, extracted.pages[block.page - 1], block.element,
                               1, None, [], stats)
                    table = document.tables[-1]
                    for tag in ('w:tblpPr', 'w:tblOverlap'):
                        node = table._tbl.tblPr.find(qn(tag))
                        if node is not None: table._tbl.tblPr.remove(node)
                    for row in table.rows:
                        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
                else:
                    code = _is_code(block)
                    style = 'Heading ' + str(block.level) if block.kind.startswith('H') and block.level else \
                        'Ruyi Figure Caption' if block.kind == 'Caption' and re.match(r'^\s*图\s*\d+', _text(block.chars)) else \
                        'Ruyi Table Caption' if block.kind == 'Caption' and re.match(r'^\s*表\s*\d+', _text(block.chars)) else \
                        'Ruyi Code' if code else 'Normal'
                    p = document.add_paragraph(style=style)
                    _set_paragraph(p, block, margins, code=code)
                    next_block = blocks[i + 1] if i + 1 < len(blocks) else None
                    toc_title = next_block and next_block.kind == 'toc' and next_block.page == block.page \
                        and len(_text(block.chars).strip()) < 40
                    if previous_kind == 'toc' or toc_title:
                        p.paragraph_format.page_break_before = True
                    inline_graphics = [g for g in block.parts if g.element.type == 'Formula']
                    _write_text(p, block.chars, code=code, graphics=inline_graphics,
                                picture_callback=lambda graphic: _picture(document, graphic, indexed, pdf, margins, stats,
                                                                         paragraph=p),
                                orphan_picture_callback=lambda graphic: _picture(document, graphic, indexed, pdf, margins, stats))
                    if block.bookmark:
                        _add_bookmark(p, block.bookmark, i + 1)
                    # Inline formulas are local pictures; other large inline
                    # illustrations get their own ordinary body paragraph.
                    for graphic in block.parts:
                        if graphic.element.type == 'Formula':
                            continue
                        _picture(document, graphic, indexed, pdf, margins, stats,
                                 paragraph=None)
                    stats.paragraphs += 1
                    if block.tagged:
                        stats.tagged_paragraphs += 1
                        if len({c['page_number'] for c in block.chars}) > 1:
                            stats.cross_page_paragraphs += 1
                    if block.kind.startswith('H') and block.level: stats.headings += 1
                    if code: stats.code_blocks += 1
                previous_kind, previous_page = block.kind, block.page
                if block.page not in reported:
                    reported.add(block.page)
                    if progress: progress(min(block.page, pages), pages)
            check_cancel()
            # Deliberately do not set updateFields: cached TOC is readable and
            # opening a conversion must not trigger Word's external update UI.
            document.save(generated)
            if progress: progress(pages, pages)
    finally:
        pdf.close()
    return stats
