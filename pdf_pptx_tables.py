"""Recover native PowerPoint cells from visible PDF rules and aligned text."""
from copy import deepcopy
from statistics import median


def find_tables(page):
    from pdf_docx_layout import _inside, _lines
    from pdf_docx_fidelity import Cell, Table, _ruled_table, _rules
    from pdf_docx_regions import captioned_table_regions
    from pdf_docx_tables import open_table_models

    def visible(obj):
        kind = obj.get('object_type')
        if kind not in ('line', 'rect', 'curve'):
            return True
        stroke = obj.get('stroke') or kind == 'line'
        color = obj.get('stroking_color' if stroke else 'non_stroking_color')
        if isinstance(color, (int, float)):
            white = color >= .99
        elif isinstance(color, (tuple, list)) and len(color) in (1, 3):
            white = min(color) >= .99
        else:
            white = False
        if white:
            return False
        return bool(stroke or (obj.get('fill') and min(obj['width'], obj['height']) <= 2))

    models = []
    for table in page.filter(visible).find_tables({'snap_tolerance': 1, 'join_tolerance': 1,
                                                  'intersection_tolerance': 2}):
        model = _ruled_table(page, table)
        if len(model.xs) < 3 or len(model.ys) < 3:
            continue
        populated = sum(any(c.get('text', '').strip() for c in cell.chars) for cell in model.cells)
        if populated >= max(2, len(model.cells)*.3):
            models.append(model)
    rules = _rules(page)
    regions = captioned_table_regions(page, [{'lines': _lines(page.chars)}], rules)
    models += open_table_models(page, regions, rules, Table, Cell, existing=models)
    # A character must belong to one reconstructed cell, never to two tables.
    accepted, claimed = [], set()
    for model in sorted(models, key=lambda m: (m.bbox[1], m.bbox[0])):
        keys = {id(c) for cell in model.cells for c in cell.chars}
        if keys & claimed:
            continue
        claimed.update(keys); accepted.append(model)
    return accepted, rules, claimed


def contains_vector(obj, model):
    a, b, c, d = model.bbox
    return (obj['x0'] >= a-1 and obj['x1'] <= c+1 and
            obj['top'] >= b-1 and obj['bottom'] <= d+1)


def write_table(slide, page, model, rules, transform, write_text, element, emu):
    from pptx.dml.color import RGBColor
    from pptx.util import Pt
    from pdf_docx_layout import _lines
    from pdf_docx_fidelity import _border
    from pdf_pptx_layout import _baseline, _rgb

    ox, oy, scale = transform
    x0, y0, x1, y1 = model.bbox
    shape = slide.shapes.add_table(len(model.ys)-1, len(model.xs)-1,
                                   emu(ox+x0*scale), emu(oy+y0*scale),
                                   emu((x1-x0)*scale), emu((y1-y0)*scale))
    shape.name = 'PDF表格 可编辑单元格'
    table = shape.table
    props = table._tbl.tblPr
    for item in list(props):
        props.remove(item)
    for name in ('firstRow', 'firstCol', 'lastRow', 'lastCol', 'bandRow', 'bandCol'):
        props.set(name, '0')
    for i, width in enumerate(b-a for a, b in zip(model.xs, model.xs[1:])):
        table.columns[i].width = emu(width*scale)
    for i, height in enumerate(b-a for a, b in zip(model.ys, model.ys[1:])):
        table.rows[i].height = emu(height*scale)
    for source in model.cells:
        cell = table.cell(source.row, source.col)
        if source.rowspan > 1 or source.colspan > 1:
            cell.merge(table.cell(source.row+source.rowspan-1, source.col+source.colspan-1))
        a, b, c, d = source.box
        cell.margin_left = cell.margin_right = cell.margin_top = cell.margin_bottom = 0
        cell.fill.background()
        backgrounds = [r for r in page.rects if r.get('fill') and
                       abs(r['x0']-a)<1 and abs(r['top']-b)<1 and
                       abs(r['x1']-c)<1 and abs(r['bottom']-d)<1]
        if backgrounds and backgrounds[-1].get('_ppt_fill_alpha', 255) > 0:
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(*_rgb(backgrounds[-1].get('non_stroking_color')))
            alpha = backgrounds[-1].get('_ppt_fill_alpha', 255)
            if alpha < 255:
                cell.fill.fore_color._xFill.get_or_change_to_srgbClr().append(element('a:alpha', val=round(alpha/255*100000)))
        tcpr = cell._tc.get_or_add_tcPr()
        for name, axis, position, start, end in (
                ('lnL', 'v', a, b, d), ('lnR', 'v', c, b, d),
                ('lnT', 'h', b, a, c), ('lnB', 'h', d, a, c)):
            node = element('a:'+name)
            evidence = _border(rules, axis, position, start, end)
            if evidence:
                node.set('w', str(emu(evidence[4]*scale)))
                fill = element('a:solidFill'); fill.append(element('a:srgbClr', val=evidence[5])); node.append(fill)
            else:
                node.append(element('a:noFill'))
            tcpr.append(node)
        lines = _lines(source.chars)
        frame = cell.text_frame; frame.clear()
        frame.word_wrap = False
        body = frame._txBody.bodyPr
        body.set('anchor', 't'); body.set('wrap', 'none')
        if not lines:
            continue
        left = min(ch['x0'] for line in lines for ch in line)
        size = max(ch.get('size', 10) for ch in lines[0])
        first_base = median(_baseline(ch, page.height) for ch in lines[0])
        cell.margin_left = emu(max(0, left-a)*scale)
        cell.margin_top = emu(max(0, first_base-size*.93-b)*scale)
        for paragraph in list(frame._txBody.p_lst):
            frame._txBody.remove(paragraph)
        previous_base = None
        for line in lines:
            for char in line:
                char['_page_height'] = page.height
            fragment = write_text(slide, line, transform)
            para = deepcopy(fragment.text_frame.paragraphs[0]._p)
            frame._txBody.append(para)
            from pptx.text.text import _Paragraph
            paragraph = _Paragraph(para, frame)
            indent = min(ch['x0'] for ch in line)-left
            paragraph._p.get_or_add_pPr().set('marL', str(emu(indent*scale)))
            base = median(_baseline(ch, page.height) for ch in line)
            paragraph.space_before = paragraph.space_after = Pt(0)
            if previous_base is not None:
                paragraph.line_spacing = Pt(max(.5, base-previous_base)*scale)
            previous_base = base
            slide.shapes._spTree.remove(fragment._element)
    return shape
