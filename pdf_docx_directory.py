"""Recover untagged directories from text and repeated page-label geometry.

This is deliberately separate from automatic TOC fields: inferred page labels
are source text, not evidence of reliable Word heading/caption associations.
No filenames, page indices or document-specific titles are used.
"""
from collections import defaultdict
from dataclasses import replace
from statistics import median
import re

from pdf_docx_flow import Block, _box, _lines


_HEADING = re.compile(r'^(?:目录|总目录|图目录|表目录|插图目录|图表目录|'
                      r'contents|tableofcontents|listoftables|listoffigures)$', re.I)
_LABEL = re.compile(r'(?:\d{1,5}|[ivxlcdmIVXLCDM]{1,10}|[A-Z][-.]\d{1,4})$')
_LEADERS = set('.．·…⋯_—–')


def directory_heading(line):
    return bool(_HEADING.fullmatch(re.sub(r'\s+', '', ''.join(c.get('text', '') for c in line))))


def entry_parts(line):
    """Return title, leaders, page-label glyphs only with visible separation."""
    line = sorted(line, key=lambda c: c['x0'])
    text = ''.join(c.get('text', '') for c in line)
    match = _LABEL.search(text.rstrip())
    if not match:
        return None
    offsets = []; pos = 0
    for char in line:
        offsets.append(pos); pos += len(char.get('text', ''))
    tail = next((i for i, start in enumerate(offsets) if start == match.start()), None)
    if tail is None or tail == 0:
        return None
    label = line[tail:]
    start = tail
    while start and (line[start-1].get('text', '').isspace() or
                     set(line[start-1].get('text', '')) <= _LEADERS):
        start -= 1
    title, leaders = line[:start], line[start:tail]
    if len(''.join(c.get('text', '') for c in title).strip()) < 2:
        return None
    dotted = sum(c.get('text', '') in _LEADERS for c in leaders) >= 2
    size = median(c.get('size', 12) for c in line)
    gap = label[0]['x0'] - max(c['x1'] for c in title)
    if not dotted and gap < size * 1.4:
        return None
    return title, leaders, label


def recover_directories(blocks, pages):
    """Split directory entries without claiming tables, code, or true TOC tags."""
    eligible = defaultdict(list)
    for block in blocks:
        if block.kind in ('P', 'Caption', 'H', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6'):
            for char in block.chars:
                eligible[char['page_number']].append(char)
    used = set(); recovered = []; previous = None
    for page in pages:
        # A tagged parent/child can expose a glyph twice; identity is page-local.
        chars = list({c['_flow_key']: c for c in eligible[page.page_number]}.values())
        lines = _lines(chars)
        if not lines:
            previous = None
            continue
        headings = [i for i, line in enumerate(lines) if directory_heading(line)]
        candidates = [(i, entry_parts(line)) for i, line in enumerate(lines)]
        candidates = [(i, parts) for i, parts in candidates if parts]
        if not candidates:
            previous = None
            continue
        # Cluster the right edges of page labels rather than guessing a margin.
        clusters = []
        for index, parts in candidates:
            right = max(c['x1'] for c in parts[2])
            group = next((g for g in clusters if abs(median(v[2] for v in g)-right) <= max(3, page.width*.012)), None)
            if group is None:
                group = []; clusters.append(group)
            group.append((index, parts, right))
        accepted = []
        for group in clusters:
            first, last = group[0][0], group[-1][0]
            heading = max((h for h in headings if h < first), default=None)
            context = heading is not None or (previous and previous[0] == page.page_number-1 and
                                              abs(previous[1]-median(v[2] for v in group)) < page.width*.025)
            dotted = sum(sum(c.get('text', '') in _LEADERS for c in p[1]) >= 2 for _, p, _ in group)
            # Two entries are sufficient under an explicit heading/continuation;
            # without that evidence require a dense repeated leader pattern.
            if context:
                if len(group) < 2 or len(group)/max(1,last-first+1) < .4:
                    continue
            elif len(group) < 5 or dotted/len(group) < .9 or len(group)/(last-first+1) < .7:
                continue
            accepted.append((group, heading))
        page_rights = []
        for group, heading in accepted:
            entries = []; previous_end = heading if heading is not None else group[0][0]-1
            for index, parts, right in group:
                start = index
                # Wrapped titles end at the line carrying leaders/page label.
                # Never swallow the heading or more than three preceding lines.
                while start > previous_end+1 and index-start < 3:
                    above, below = lines[start-1], lines[start]
                    size = median(c.get('size',12) for c in below)
                    gap = min(c['top'] for c in below)-max(c['bottom'] for c in above)
                    if directory_heading(above) or entry_parts(above) or not -1 <= gap <= size*1.8:
                        break
                    start -= 1
                entry = [c for line in lines[start:index+1] for c in line]
                if any(c['_flow_key'] in used for c in entry):
                    continue
                entries.append(entry); used.update(c['_flow_key'] for c in entry)
                previous_end = index; page_rights.append(right)
            if entries:
                all_chars = [c for entry in entries for c in entry]
                recovered.append(Block('directory', all_chars, page=page.page_number,
                                       box=_box(all_chars), entries=entries))
        previous = (page.page_number, median(page_rights)) if page_rights else None
    result = []
    for block in blocks:
        remaining = [c for c in block.chars if c['_flow_key'] not in used]
        if len(remaining) == len(block.chars):
            result.append(block)
        elif remaining:
            number = min(c['page_number'] for c in remaining)
            result.append(replace(block, chars=remaining, page=number,
                                  box=_box([c for c in remaining if c['page_number']==number])))
        else:
            result.extend(block.parts)
    return sorted(result+recovered, key=lambda block:block.key)
