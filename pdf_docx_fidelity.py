"""Editable paragraphs on the source PDF pages, with measured typography.

PDF physical line breaks are soft breaks inside a logical paragraph. Page
sections, original advances, rules and pictures keep their measured positions.
No whole-page image is used for a page with an extractable text layer.
"""
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from math import hypot
from statistics import median
import re

from pdf_docx_flow import Source, Block, _box, _clean, _walk, _normal, _counter, _field
from pdf_docx_layout import (_lines, _baseline, _font_name, _rgb_hex, _element,
                             _anchor_picture, _PagePictures, _annotate_bold,
                             _add_vectors, find_ruled_tables, _inside)
from pdf_docx_regions import (equation_regions, captioned_table_regions, merge_boxes,
                              inline_formula_fragments, classify_inline_math, exclude_formula_overlap)
from pdf_docx_tables import open_table_models
from pdf_docx_directory import recover_directories, entry_parts


@dataclass
class FidelityStats:
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
    directory_entries: int = 0
    vector_shapes: int = 0
    positioned_paragraphs: int = 0
    raster_tables: int = 0
    formula_regions: int = 0
    warnings: list = field(default_factory=list)


@dataclass
class Cell:
    row: int
    col: int
    rowspan: int
    colspan: int
    box: tuple
    chars: list


@dataclass
class Table:
    page: int
    bbox: tuple
    xs: list
    ys: list
    cells: list
    fit_scale: float = 1


def _base(line, height):
    return height + median(_baseline(c) for c in line)


def _tiny(p):
    from docx.shared import Pt
    pf=p.paragraph_format
    pf.space_before=pf.space_after=Pt(0)
    pf.line_spacing=Pt(.05)
    pf.widow_control=False
    pf.keep_with_next=False
    p.add_run().font.size=Pt(1)


def _defaults(p):
    from docx.shared import Pt
    pf=p.paragraph_format
    pf.space_after=Pt(0)
    pf.widow_control=False
    pf.keep_together=False
    pf.keep_with_next=False
    pr=p._p.get_or_add_pPr()
    for tag in ('autoSpaceDE','autoSpaceDN','snapToGrid','contextualSpacing'):
        pr.append(_element('w:'+tag,val=0))


def _font(char):
    name=_font_name(char.get('fontname','SimSun'))
    return {'DengXian':'等线','DengXian-Bold':'等线','DengXian-Light':'等线 Light',
            'NSimSun':'新宋体','SimSun,Bold':'宋体','SimHei,Bold':'黑体'}.get(name,name)


def _glyph(p, char, gap=0):
    from docx.shared import Pt, RGBColor
    from docx.oxml.ns import qn
    text=_clean(char.get('text',''))
    if not text:
        return None
    r=p.add_run(text)
    size=max(1,min(200,float(char.get('size',12))))
    r.font.name=_font(char)
    r.font.size=Pt(round(size*2)/2)
    original=char.get('fontname','').lower()
    r.font.bold=bool(char.get('_ruyi_bold')) or 'bold' in original or 'black' in original
    r.font.italic='italic' in original or 'oblique' in original
    r.font.color.rgb=RGBColor.from_string(_rgb_hex(char.get('non_stroking_color')))
    pr=r._r.get_or_add_rPr()
    fonts=pr.find(qn('w:rFonts'))
    for key in ('ascii','hAnsi','eastAsia','cs'):
        fonts.set(qn('w:'+key),_font(char))
    # A separate run prevents Word from applying a new pair-kerning or CJK
    # spacing decision to a PDF which already specifies every glyph advance.
    matrix=char.get('matrix')
    if matrix and len(matrix)==6 and hypot(matrix[2],matrix[3])>0:
        ratio=hypot(matrix[0],matrix[1])/hypot(matrix[2],matrix[3])
        if abs(ratio-1)>.002:
            pr.append(_element('w:w',val=max(1,min(600,round(ratio*100)))))
    if abs(gap)>.015:
        pr.append(_element('w:spacing',val=max(-31680,min(31680,round(gap*20)))))
    if char.get('_ruyi_script_shift'):
        pr.append(_element('w:position',val=char['_ruyi_script_shift']))
    pr.append(_element('w:kern',val=1000))
    return r


_MATH_FONT = re.compile(r'(?i)(cambria\s*math|stix|mathjax|dejavu\s*math|xits|symbolmt|asana\s*math|latin\s*modern\s*math)')
_MATH_SYMBOLS = set('∫∑∏√±×÷≤≥≠≈∞∂∇∈∉⊂⊃∪∩→←↔∀∃αβγδεζηθικλμνξοπρστυφχψωΓΔΘΛΞΠΣΦΨΩ')


def _is_math_char(char):
    """A PDF keeps positioned glyphs only, so math must be recognised by font."""
    if _MATH_FONT.search(str(char.get('fontname') or '')):
        return True
    text=char.get('text') or ''
    return bool(text) and all(ch in _MATH_SYMBOLS for ch in text)


def _needs_math_raster(char):
    # A mapped Greek letter or root sign in an ordinary font is editable text.
    # Special math fonts can have missing/wrong glyph metrics even with Unicode.
    return bool(_MATH_FONT.search(str(char.get('fontname') or '')))


def _formula_picture(p, pictures, chars):
    """Rasterise one math run in place, refusing blank crops.

    A display formula stacks glyphs on several baselines (numerator, fraction
    bar, denominator, superscript), so a run taken from one text line can land
    between the ink strokes. Emitting such a crop would silently delete the
    formula, so a crop without ink is refused and the caller keeps the text.
    """
    from io import BytesIO
    from docx.shared import Pt
    from PIL import Image
    pad=.6
    box=(min(c['x0'] for c in chars)-pad,min(c['top'] for c in chars)-pad*.5,
         max(c['x1'] for c in chars)+pad,max(c['bottom'] for c in chars)+pad*.5)
    stream=pictures.crop(box)
    data=stream.read()
    try:
        with Image.open(BytesIO(data)) as probe:
            grey=probe.convert('L')
            ink=sum(1 for value in grey.getdata() if value<200)
            area=grey.width*grey.height
    except OSError:
        return False
    if area==0 or ink<max(4,int(area*.002)):
        return False
    p.add_run().add_picture(BytesIO(data),width=Pt(box[2]-box[0]),height=Pt(box[3]-box[1]))
    return True


# Formula rasterisation works at page level: a display formula stacks glyphs on
# several baselines (numerator, bar, denominator, superscript), so it has to be
# grouped as one region before text lines are assembled. Cropping per text line
# produced fragments and is never used.
_RASTER_FORMULAS = True
# Native tables use measured fixed columns and source-evidenced borders.
# Keep this legacy fallback disabled; ordinary tables and captions are editable.
_RASTER_TABLES = False


def _formula_regions(page):
    """Whole-line boxes that carry math, for raster fallback.

    A display formula stacks glyphs on several baselines and mixes font subsets,
    so per-glyph boxes miss half the ink. Cropping the whole line keeps every
    glyph and the PDF's own spacing, so the formula stays exactly in place.
    """
    regions=[]
    for line in _lines(page.chars):
        if not line:
            continue
        math=sum(1 for c in line if _is_math_char(c))
        if math<2 or math/len(line)<.3:
            continue
        regions.append([min(c['x0'] for c in line)-1,min(c['top'] for c in line)-1,
                        max(c['x1'] for c in line)+1,max(c['bottom'] for c in line)+1])
    # A fraction puts numerator, bar and denominator on separate baselines, so
    # per-line boxes hold only a slice. Merge vertically adjacent boxes that
    # share horizontal span, which reassembles the whole formula.
    merged=[]
    for box in sorted(regions,key=lambda b:b[1]):
        for other in merged:
            overlap=min(box[2],other[2])-max(box[0],other[0])
            narrow=min(box[2]-box[0],other[2]-other[0])
            if overlap>narrow*.3 and box[1]-other[3]<18:
                other[0]=min(other[0],box[0]);other[1]=min(other[1],box[1])
                other[2]=max(other[2],box[2]);other[3]=max(other[3],box[3])
                break
        else:
            merged.append(box)
    return [tuple(round(v,2) for v in box) for box in merged]


def _write_lines(p, lines, page, leading, *, origin=0, cell=False, field_entry=None, pictures=None):
    from docx.shared import Pt
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    if not lines:
        return
    _defaults(p)
    pf=p.paragraph_format
    # The paragraph mark participates in Word's baseline calculation even
    # when every visible run is smaller than the template's 12-point font.
    # Match it as well, especially for nine-point registration forms.
    mark=p._p.get_or_add_pPr().find(qn('w:rPr'))
    if mark is None:
        mark=OxmlElement('w:rPr');p._p.get_or_add_pPr().append(mark)
    mark.append(_element('w:sz',val=round(max(c.get('size',12) for line in lines for c in line)*2)))
    fonts=OxmlElement('w:rFonts')
    for key in ('ascii','hAnsi','eastAsia','cs'):fonts.set(qn('w:'+key),_font(lines[0][0]))
    mark.append(fonts)
    left=min(c['x0'] for line in lines for c in line)
    first=min(c['x0'] for c in lines[0])
    pf.left_indent=Pt(left-origin)
    pf.first_line_indent=Pt(first-left)
    if not cell:
        right=max(c['x1'] for line in lines for c in line)
        # Source soft breaks already set the right edge of every physical
        # line. A second Word wrapping decision must not move the last glyph
        # solely because the host font has a fractional advance difference.
        pf.right_indent=Pt(0)
    else:
        pf.right_indent=Pt(-4)
    # A stacked fraction or radicand is taller than the text line it sits on.
    # Exact spacing clips the raster crop, so measure grouped math runs (single
    # glyph heights would leave a fraction only its top half visible).
    math_boxes=[c['_ruyi_formula_box'] for line in lines for c in line if '_ruyi_formula_box' in c]
    if math_boxes:
        leading=max(leading,max(box[3]-box[1] for box in math_boxes)*1.15)
    pf.line_spacing=Pt(leading)
    # Only the source indentation levels become tab stops. Soft line breaks
    # then retain code/list indentation without a movable frame for each line.
    stops=sorted({round(min(c['x0'] for c in line)-origin,2) for line in lines[1:]
                  if min(c['x0'] for c in line)>left+.2})
    for stop in stops:
        pf.tab_stops.add_tab_stop(Pt(stop))
    for index,line in enumerate(lines):
        if index:
            p.add_run().add_break()
            desired=min(c['x0'] for c in line)-origin
            for stop in stops:
                if stop<=desired+.1:
                    p.add_run('\t')
        number=0
        while number<len(line):
            char=line[number]
            if pictures is not None and '_ruyi_formula_box' in char:
                box=char['_ruyi_formula_box']
                if number and '_ruyi_formula_box' in line[number-1] and char['x0']-line[number-1]['x1']>char.get('size',12):
                    p.paragraph_format.tab_stops.add_tab_stop(Pt(char['x0']-origin))
                    p.add_run('\t')
                p.add_run().add_picture(pictures.crop(box),width=Pt(box[2]-box[0]),height=Pt(box[3]-box[1]))
                number+=1
                continue
            gap=line[number+1]['x0']-char['x1'] if number+1<len(line) else 0
            _glyph(p,char,gap);number+=1


def _integrate_formula_fragments(items,fragments,page):
    """Keep prose as native runs; attach only formula fragments to its line."""
    removed={id(c) for _,group in fragments for c in group}
    result=[]
    for item in items:
        copied=dict(item)
        copied['lines']=[[c for c in line if id(c) not in removed] for line in item['lines']]
        copied['lines']=[line for line in copied['lines'] if line]
        if copied['lines']:result.append(copied)
    for placeholder,group in fragments:
        desired=_base([placeholder],page.height)
        options=[]
        for item in result:
            for line in item['lines']:
                distance=abs(_base(line,page.height)-desired)
                if distance<=max(2.5,placeholder['size']*.5):options.append((distance,line))
        if options:
            line=min(options,key=lambda x:x[0])[1];line.append(placeholder);line.sort(key=lambda c:c['x0'])
        else:
            block=Block('P',[placeholder],page=page.page_number,box=_box([placeholder]))
            result.append({'kind':'P','lines':[[placeholder]],'block':block,'id':0})
    return result


def _integrate_native_scripts(items,page,tables):
    """Attach mapped small scripts to nearby letters without making pictures."""
    chars=[c for item in items if item['kind'] not in ('toc','directory','Code')
           for line in item['lines'] for c in line
           if not any(_inside(c,t.bbox) for t in tables)]
    if not chars:return items
    sizes=[c.get('size',12) for c in chars if c.get('text','').strip() and c.get('size',0)>0]
    usual=median(sizes) if sizes else 12
    scripts=[c for c in chars if 0<c.get('size',0)<usual*.75 and
             re.fullmatch(r'[A-Za-z0-9]',c.get('text','')) and not _needs_math_raster(c)]
    normal=[c for c in chars if c.get('size',0)>=usual*.75 and
            re.fullmatch(r'[A-Za-z]',c.get('text','')) and not _needs_math_raster(c)]
    attached={}
    for c in scripts:
        base=_base([c],page.height)
        parents=[p for p in normal if -.5<=c['x0']-p['x1']<=max(3,p['size']*.7) and
                 .6<abs(base-_base([p],page.height))<=p['size']*.8]
        if parents:
            p=min(parents,key=lambda p:(abs(c['x0']-p['x1']),abs(base-_base([p],page.height))))
            c['_ruyi_script_shift']=round((_base([p],page.height)-base)*2)
            attached[id(c)]=p
    if not attached:return items
    result=[]
    for item in items:
        entry=dict(item)
        entry['lines']=[[c for c in line if id(c) not in attached] for line in item['lines']]
        entry['lines']=[line for line in entry['lines'] if line]
        if entry['lines']:result.append(entry)
    for c in scripts:
        parent=attached.get(id(c))
        if parent is None:continue
        for item in result:
            line=next((line for line in item['lines'] if any(v is parent for v in line)),None)
            if line is not None:
                line.append(c);line.sort(key=lambda v:v['x0']);break
    return result


def _write_directory(p, lines, page, leading, *, origin=0, positioned=False):
    """Keep editable titles and source page labels with a right-aligned tab."""
    from docx.shared import Pt
    from docx.enum.text import WD_TAB_ALIGNMENT, WD_TAB_LEADER
    parts = entry_parts(lines[-1]) if lines else None
    if not parts:
        _write_lines(p, lines, page, leading, origin=origin, cell=positioned)
        return
    title, leaders, label = parts
    _write_lines(p, lines[:-1]+[title], page, leading, origin=origin, cell=positioned)
    dotted = any(c.get('text','') in '.．·…⋯_' for c in leaders)
    p.paragraph_format.tab_stops.add_tab_stop(Pt(max(c['x1'] for c in label)-origin),
        WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS if dotted else WD_TAB_LEADER.SPACES)
    p.add_run('\t')
    for i, char in enumerate(label):
        _glyph(p, char, label[i+1]['x0']-char['x1'] if i+1<len(label) else 0)


def _groups(chars, height):
    """Tags are evidence, but a large visual paragraph gap can refine them."""
    lines=_lines(chars)
    if not lines:
        return []
    deltas=[_base(b,height)-_base(a,height) for a,b in zip(lines,lines[1:])]
    normal=median(v for v in deltas if v>3) if any(v>3 for v in deltas) else 0
    left=min(c['x0'] for c in chars)
    right=max(c['x1'] for c in chars)
    result=[];current=[]
    for line in lines:
        split=False
        if current:
            previous=current[-1]
            delta=_base(line,height)-_base(previous,height)
            size=median(c.get('size',12) for c in line)
            indent=min(c['x0'] for c in line)-left
            previous_short=max(c['x1'] for c in previous)<right-size*.8
            split=(normal>0 and delta>normal*1.35+.4) or \
                  (indent>size*.8 and previous_short and len(current)>1) or \
                  abs(median(c.get('size',12) for c in previous)-size)>2
        if split:
            result.append(current);current=[]
        current.append(line)
    if current:result.append(current)
    return result


def _leading(lines, default=18):
    if len(lines)>1:
        # The page height cancels; _baseline has a consistent orientation.
        ds=[median(_baseline(c) for c in b)-median(_baseline(c) for c in a)
            for a,b in zip(lines,lines[1:])]
        ds=[v for v in ds if v>2]
        if ds:return median(ds)
    return max(max(c.get('size',12) for c in lines[0])*1.15,default) if lines else default


def _rules(page):
    result=[]
    def add(x0,y0,x1,y1,width,color,dash=None):
        if abs(y1-y0)<.5 and abs(x1-x0)>1:
            result.append(('h',(y0+y1)/2,min(x0,x1),max(x0,x1),max(.1,width),color,dash))
        elif abs(x1-x0)<.5 and abs(y1-y0)>1:
            result.append(('v',(x0+x1)/2,min(y0,y1),max(y0,y1),max(.1,width),color,dash))
    for line in page.lines:
        color=_rgb_hex(line.get('stroking_color'))
        if color!='FFFFFF':
            add(line['x0'],line['top'],line['x1'],line['bottom'],line.get('linewidth',.5),color,line.get('dash'))
    for rect in page.rects:
        x0,y0,x1,y1=rect['x0'],rect['top'],rect['x1'],rect['bottom']
        if rect.get('stroke'):
            color=_rgb_hex(rect.get('stroking_color'))
            if color=='FFFFFF':continue
            for coords in [(x0,y0,x1,y0),(x0,y1,x1,y1),(x0,y0,x0,y1),(x1,y0,x1,y1)]:
                add(*coords,rect.get('linewidth',.5),color,rect.get('dash'))
        elif rect.get('fill') and min(x1-x0,y1-y0)<=2 and _rgb_hex(rect.get('non_stroking_color'))!='FFFFFF':
            if x1-x0>y1-y0:add(x0,(y0+y1)/2,x1,(y0+y1)/2,y1-y0,_rgb_hex(rect.get('non_stroking_color')))
            else:add((x0+x1)/2,y0,(x0+x1)/2,y1,x1-x0,_rgb_hex(rect.get('non_stroking_color')))
    return result


def _tag_scale(blocks):
    scales=[]
    for b in blocks:
        if b.kind!='table' or not b.element:continue
        a=b.element.attributes
        if a.get('BBox') and float(a.get('Width',0) or 0)>0:
            bounds=a['BBox'];scales.append(abs(float(bounds[2])-float(bounds[0]))/float(a['Width']))
    return median(scales) if scales else 1


def _tag_table(source, block, number, unit):
    page=source.pages[number-1]
    allrows=[n for n in _walk(block.element) if n.type=='TR']
    placed=[];occupied=set();maxcol=0
    for rownum,row in enumerate(allrows):
        col=0
        for cell in [n for n in row.children if n.type in ('TD','TH')]:
            while (rownum,col) in occupied:col+=1
            cs=max(1,int(cell.attributes.get('ColSpan',1)))
            rs=max(1,int(cell.attributes.get('RowSpan',1)))
            chars=source.element_chars(cell)
            for r in range(rownum,rownum+rs):
                for c in range(col,col+cs):occupied.add((r,c))
            placed.append((rownum,col,rs,cs,cell,chars));col+=cs;maxcol=max(maxcol,col)
    if not placed or maxcol<1:return None
    active=sorted({r for r,c,rs,cs,node,chars in placed if any(ch['page_number']==number for ch in chars)})
    if not active:return None
    start,end=active[0],active[-1]+1
    # Keep empty rows enclosed between populated rows.
    count=end-start
    widths=[[] for _ in range(maxcol)];heights=[[] for _ in allrows]
    spanwidths=[];spanheights=[]
    for r,c,rs,cs,node,chars in placed:
        w=float(node.attributes.get('Width',0) or 0)*unit
        h=float(node.attributes.get('Height',0) or 0)*unit
        if w>0:
            if cs==1:widths[c].append(w)
            else:spanwidths.append((cs,c,w))
        if h>0:
            if rs==1:heights[r].append(h)
            else:spanheights.append((rs,r,h))
    ws=[median(v) if v else None for v in widths]
    hsall=[median(v) if v else None for v in heights]
    # A full-width merged footnote does not imply equal-width columns. Solve
    # only unknown widths from span constraints after using unmerged cells.
    for spans,values in ((spanwidths,ws),(spanheights,hsall)):
        for count,index,total in sorted(spans):
            stop=min(len(values),index+count)
            missing=[j for j in range(index,stop) if values[j] is None]
            if missing:
                remaining=total-sum(values[j] or 0 for j in range(index,stop))
                for j in missing:values[j]=max(.1,remaining/len(missing))
    ws=[v if v is not None else 30 for v in ws]
    hs=[hsall[i] if hsall[i] is not None else 18 for i in range(start,end)]
    chars=[ch for *_,cc in placed for ch in cc if ch['page_number']==number]
    box=_box(chars)
    if not box:return None
    attrs=block.element.attributes
    bbox=attrs.get('BBox')
    single=len({ch['page_number'] for *_,cc in placed for ch in cc})==1
    width=float(attrs.get('Width',0) or 0)*unit or sum(ws)
    left=(page.width-width)/2
    top=box[1]-max(0,(hs[0]-(max(c['bottom'] for c in chars if c['top']<box[1]+2)-box[1]))/2)
    height=sum(hs)
    if bbox and single:
        left=min(bbox[0],bbox[2]);width=abs(bbox[2]-bbox[0])
        top=page.height-max(bbox[1],bbox[3]);height=abs(bbox[3]-bbox[1])
    # The producer's BBox can describe the containing paragraph. Table indents
    # then narrow or expand the actual grid inside that paragraph.
    startindent=float(attrs.get('StartIndent',0) or 0)*unit
    endindent=float(attrs.get('EndIndent',0) or 0)*unit
    left+=startindent;width=max(1,width-startindent-endindent)
    # Visible full-width horizontal rules supply stronger bounds than glyph
    # boxes or producer-specific attribute units, including three-line tables.
    horizontal=[rule for rule in _rules(page) if rule[0]=='h' and
                abs(rule[3]-rule[2]-width)<max(12,width*.2) and
                rule[2]<=box[0]+3 and rule[3]>=box[2]-3 and
                box[1]-max(hs[0],24)*1.4<=rule[1]<=box[3]+max(hs[-1],24)*1.4]
    above=[v for v in horizontal if v[1]<=box[1]+1]
    below=[v for v in horizontal if v[1]>=box[3]-1]
    if above and below:
        a=max(above,key=lambda v:v[1]);b=min(below,key=lambda v:v[1])
        if b[1]>a[1]+4:
            left=min(a[2],b[2]);width=max(a[3],b[3])-left
            top=a[1];height=b[1]-top
    width=max(width,1);height=max(height,1)
    # Empty TDs may be omitted in any row. A complete unmerged row supplies
    # column widths without mixing widths from shifted ordinal TD positions.
    complete=[]
    for rownum in range(len(allrows)):
        values=[(c,cs,float(node.attributes.get('Width',0) or 0)*unit)
                for r,c,rs,cs,node,cc in placed if r<=rownum<r+rs]
        values.sort()
        if len(values)==maxcol and all(cs==1 and w>0 for c,cs,w in values):
            total=sum(w for c,cs,w in values)
            if abs(total-width)<max(1.5,width*.01):
                complete.append([w*width/total for c,cs,w in values])
    if complete:
        ws=[median(row[j] for row in complete) for j in range(maxcol)]
    xs=[left]
    for w in ws:xs.append(xs[-1]+w/sum(ws)*width)
    # Some producers omit completely empty TD elements. Ordinal TD order is
    # then not the visual column index (e.g. sparse factor-loading matrices).
    # Centered numeric/code cells plus their declared widths can recover the
    # full grid, including columns which are empty in an individual row.
    from pdf_docx_layout import _cluster
    evidence=[left,left+width]
    for r,c,rs,cs,node,cc in placed:
        current=[ch for ch in cc if ch['page_number']==number]
        text=re.sub(r'\s+','',''.join(ch.get('text','') for ch in current))
        cellwidth=float(node.attributes.get('Width',0) or 0)*unit
        if cs==1 and current and cellwidth>0 and re.fullmatch(r'(?:[A-Za-z]{1,4}\d+|[-+]?\d+(?:[.,]\d+)?(?:[eE][-+]?\d+)?%?)',text):
            lines=_lines(current)
            center=median((min(ch['x0'] for ch in line)+max(ch['x1'] for ch in line))/2 for line in lines)
            lo,hi=center-cellwidth/2,center+cellwidth/2
            if left-1.5<=lo and hi<=left+width+1.5:evidence.extend((lo,hi))
    inferred=_cluster(evidence,tolerance=1.3)
    inferred_grid=len(inferred)==maxcol+1 and abs(inferred[0]-left)<1.5 and abs(inferred[-1]-left-width)<1.5
    if inferred_grid:
        xs=[left]+inferred[1:-1]+[left+width]
    ys=[top]
    for h in hs:ys.append(ys[-1]+h/sum(hs)*height)
    rowchars=defaultdict(list)
    vertical_mismatch=False
    for r,c,rs,cs,node,cc in placed:
        current=[ch for ch in cc if ch['page_number']==number]
        if not current or r>=end or r+rs<=start:continue
        rr=max(r,start)-start;rend=min(r+rs,end)-start
        bounds=_box(current)
        if bounds[1]<ys[rr]-1 or bounds[3]>ys[rend]+1:vertical_mismatch=True
        if rs==1:rowchars[rr].extend(current)
    if vertical_mismatch:
        # A row continuing on the next page can retain only its first-page
        # Height attribute. Do not scale every following row to compensate.
        # Visible rules between the adjacent row text give page-local edges;
        # without a rule, use the blank gap between those rows.
        rowboxes={r:_box(cc) for r,cc in rowchars.items()}
        refined=list(ys);rules=_rules(page)
        for boundary in range(1,len(ys)-1):
            before,after=rowboxes.get(boundary-1),rowboxes.get(boundary)
            if not before or not after or before[3]>after[1]:continue
            candidates=[rule[1] for rule in rules if rule[0]=='h' and
                        before[3]-.5<=rule[1]<=after[1]+.5 and
                        min(left+width,rule[3])-max(left,rule[2])>=width*.2]
            refined[boundary]=min(candidates,key=lambda y:abs(y-(before[3]+after[1])/2)) \
                                if candidates else (before[3]+after[1])/2
        if all(b>a for a,b in zip(refined,refined[1:])):ys=refined
    cells=[]
    for r,c,rs,cs,node,cc in placed:
        if r>=end or r+rs<=start:continue
        rr=max(r,start)-start;rend=min(r+rs,end)-start
        filtered=[ch for ch in cc if ch['page_number']==number]
        if filtered:
            glyphbox=_box(filtered);textcenter=(glyphbox[0]+glyphbox[2])/2
            declared=float(node.attributes.get('Width',0) or 0)*unit
            def fits(j):
                actual=xs[j+cs]-xs[j]
                return xs[j]-1<=glyphbox[0] and glyphbox[2]<=xs[j+cs]+1 and \
                       (declared<=0 or abs(actual-declared)<=max(2,declared*.04))
            def score(j):
                return abs((xs[j]+xs[j+cs])/2-textcenter)+abs((xs[j+cs]-xs[j])-declared)*.6
            candidates=[j for j in range(maxcol-cs+1) if fits(j)]
            # Use geometry for labels, significance symbols and merged
            # headers as well as numbers. Preserve a containing, width-matched
            # placement rather than assuming every label is centered.
            if candidates and c not in candidates:c=min(candidates,key=score)
        cells.append(Cell(rr,c,rend-rr,cs,(xs[c],ys[rr],xs[min(c+cs,maxcol)],ys[rend]),filtered))
    slots={}
    for item in cells:
        for r in range(item.row,item.row+item.rowspan):
            for c in range(item.col,item.col+item.colspan):
                if (r,c) in slots:
                    raise ValueError('表格定位出现重叠单元格：第'+str(number)+'页，行'+str(r)+'列'+str(c))
                slots[r,c]=item
    return Table(number,(left,top,left+width,top+height),xs,ys,cells)


def _ruled_table(page, model):
    from pdf_docx_layout import _cluster
    xs=_cluster([v for b in model.cells for v in (b[0],b[2])])
    ys=_cluster([v for b in model.cells for v in (b[1],b[3])])
    cells=[]
    for b in model.cells:
        x0=min(range(len(xs)),key=lambda i:abs(xs[i]-b[0]));x1=min(range(len(xs)),key=lambda i:abs(xs[i]-b[2]))
        y0=min(range(len(ys)),key=lambda i:abs(ys[i]-b[1]));y1=min(range(len(ys)),key=lambda i:abs(ys[i]-b[3]))
        if x1>x0 and y1>y0:
            cells.append(Cell(y0,x0,y1-y0,x1-x0,b,[c for c in page.chars if _inside(c,b)]))
    return Table(page.page_number,model.bbox,xs,ys,cells)


def _border(rules, axis, position, start, end):
    candidates=[r for r in rules if r[0]==axis and abs(r[1]-position)<=1.1 and
                min(end,r[3])-max(start,r[2])>=(end-start)*.6]
    return min(candidates,key=lambda r:abs(r[1]-position)) if candidates else None


def _add_table(document, page, model, stats, pictures=None):
    from docx.shared import Pt
    from docx.enum.table import WD_ROW_HEIGHT_RULE, WD_CELL_VERTICAL_ALIGNMENT
    from docx.oxml.ns import qn
    rows,cols=len(model.ys)-1,len(model.xs)-1
    if rows*cols>20000:raise ValueError('表格网格过于复杂，请拆分PDF后重试。')
    table=document.add_table(rows=rows,cols=cols);table.autofit=False
    props=table._tbl.tblPr
    for name in ('tblW','tblBorders','tblCellMar','tblInd','jc'):
        old=props.find(qn('w:'+name))
        if old is not None:props.remove(old)
    props.append(_element('w:tblW',w=round((model.bbox[2]-model.bbox[0])*20),type='dxa'))
    props.append(_element('w:tblInd',w=round(model.bbox[0]*20),type='dxa'))
    borders=_element('w:tblBorders')
    for side in ('top','left','bottom','right','insideH','insideV'):borders.append(_element('w:'+side,val='nil'))
    props.append(borders)
    margins=_element('w:tblCellMar')
    for side in ('top','left','bottom','right'):margins.append(_element('w:'+side,w=0,type='dxa'))
    props.append(margins)
    for i,col in enumerate(table.columns):col.width=Pt(model.xs[i+1]-model.xs[i])
    for i,row in enumerate(table.rows):
        row.height=Pt((model.ys[i+1]-model.ys[i])*model.fit_scale);row.height_rule=WD_ROW_HEIGHT_RULE.EXACTLY
        row._tr.get_or_add_trPr().append(_element('w:cantSplit'))
        for j,cell in enumerate(row.cells):cell.width=Pt(model.xs[j+1]-model.xs[j])
    rules=_rules(page)
    for item in sorted(model.cells,key=lambda c:(c.row,c.col)):
        cell=table.cell(item.row,item.col)
        if item.rowspan>1 or item.colspan>1:
            cell=cell.merge(table.cell(item.row+item.rowspan-1,item.col+item.colspan-1))
        cell.width=Pt(item.box[2]-item.box[0]);cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.TOP
        pr=cell._tc.get_or_add_tcPr();old=pr.find(qn('w:tcBorders'))
        if old is not None:pr.remove(old)
        borders=_element('w:tcBorders')
        x0,y0,x1,y1=item.box
        for side,axis,pos,start,end in [('top','h',y0,x0,x1),('bottom','h',y1,x0,x1),('left','v',x0,y0,y1),('right','v',x1,y0,y1)]:
            rule=_border(rules,axis,pos,start,end)
            borders.append(_element('w:'+side,val=('dashed' if rule[6] and rule[6][0] else 'single') if rule else 'nil',
                                    sz=max(2,round(rule[4]*8)) if rule else 0,color=rule[5] if rule else 'auto'))
        pr.append(borders)
        p=cell.paragraphs[0]
        lines=_lines(item.chars)
        if lines:
            info={'kind':'P','lines':lines,'block':Block('P',item.chars,page=page.page_number),'id':0}
            infos=_integrate_native_scripts([info],page,[])
            lines=sorted([line for entry in infos for line in entry['lines']],key=lambda line:_base(line,page.height))
        if item.chars and pictures is not None:
            fragments=inline_formula_fragments(page,item.chars,_needs_math_raster,_baseline,pictures,
                body_top=y0,body_bottom=y1)
            if fragments:
                info={'kind':'P','lines':lines,'block':Block('P',item.chars,page=page.page_number),'id':0}
                infos=_integrate_formula_fragments([info],fragments,page)
                lines=sorted([line for entry in infos for line in entry['lines']],key=lambda line:_base(line,page.height))
                stats.images+=len(fragments);stats.formula_regions+=len(fragments)
        if lines:
            leading=_leading(lines,max(c.get('size',12) for c in lines[0])*1.15)
            # Source-centered labels can fall between the data-row baselines.
            # Keep them inside their cell instead of preserving an offset that
            # clips half the glyphs against an exact-height Word row.
            available=(y1-y0)*model.fit_scale
            required=max(c.get('size',12) for c in item.chars)
            leading=max(required,min(leading,available/max(1,len(lines))))
            _write_lines(p,lines,page,leading,origin=x0,cell=True,pictures=pictures)
            p.paragraph_format.space_before=Pt(max(0,min(_base(lines[0],page.height)-y0-leading*.8,
                                                        available-leading*len(lines))))
        else:_tiny(p)
        stats.cells+=1
    # Word requires the side borders on vertical-merge continuation cells as
    # well as their master. Otherwise long left/right rules disappear at row
    # boundaries, even though the logical merged cell has a border property.
    for row_index,tr in enumerate(table._tbl.tr_lst):
        col=0
        for tc in tr.tc_lst:
            pr=tc.get_or_add_tcPr();span=pr.find(qn('w:gridSpan'))
            width=int(span.get(qn('w:val'))) if span is not None else 1
            old=pr.find(qn('w:tcBorders'))
            if old is not None:pr.remove(old)
            borders=_element('w:tcBorders')
            x0,x1=model.xs[col],model.xs[min(cols,col+width)]
            y0,y1=model.ys[row_index],model.ys[row_index+1]
            for side,axis,pos,start,end in [('top','h',y0,x0,x1),('bottom','h',y1,x0,x1),('left','v',x0,y0,y1),('right','v',x1,y0,y1)]:
                rule=_border(rules,axis,pos,start,end)
                borders.append(_element('w:'+side,val=('dashed' if rule[6] and rule[6][0] else 'single') if rule else 'nil',
                                        sz=max(2,round(rule[4]*8)) if rule else 0,color=rule[5] if rule else 'auto'))
            pr.append(borders);col+=width
    stats.tables+=1
    return table


def _bookmark(p,name,number):
    begin=_element('w:bookmarkStart',id=number,name=name);end=_element('w:bookmarkEnd',id=number)
    p._p.insert(1 if p._p.pPr is not None else 0,begin);p._p.append(end)


def _page_profile(section, page, profile, stats):
    from docx.shared import Pt
    from docx.oxml.ns import qn
    section.page_width=Pt(page.width);section.page_height=Pt(page.height)
    section.top_margin=section.bottom_margin=section.left_margin=section.right_margin=Pt(0)
    section.header_distance=section.footer_distance=Pt(0)
    # A paragraph grid inherited from the Word template changes line baselines.
    grid=section._sectPr.find(qn('w:docGrid'))
    if grid is not None:section._sectPr.remove(grid)
    counter=None
    for side in ('header','footer'):
        container=getattr(section,side);container.is_linked_to_previous=False
        for node in list(container._element):container._element.remove(node)
        for item in profile[side]:
            p=container.add_paragraph();line=item['chars'];leading=max(c.get('size',9) for c in line)*1.3
            _defaults(p);p.paragraph_format.line_spacing=Pt(leading);p.paragraph_format.space_before=Pt(0)
            box=item['box'];y=_base(line,page.height)-leading*.8
            p._p.get_or_add_pPr().append(_element('w:framePr',w=round((box[2]-box[0]+4)*20),
                 h=round(leading*20),x=round(box[0]*20),y=round(y*20),hAnchor='page',vAnchor='page',
                 hRule='atLeast',wrap='none',hSpace=0,vSpace=0))
            if item['counter']:
                counter=item['counter'];raw=_counter(item['text']);suffix=''
                if counter['fmt']!='decimal':suffix=' \\* '+('ROMAN' if counter['fmt']=='upperRoman' else 'roman')
                p.add_run(counter['before']);_field(p,'PAGE'+suffix,raw[1]);p.add_run(counter['after'])
                from docx.oxml.ns import qn
                for r in p.runs:
                    r.font.name=_font(line[0]);r.font.size=Pt(line[0].get('size',9))
                    r._r.get_or_add_rPr().find(qn('w:rFonts')).set(qn('w:eastAsia'),_font(line[0]))
                stats.page_fields+=1
            else:
                for i,c in enumerate(line):_glyph(p,c,line[i+1]['x0']-c['x1'] if i+1<len(line) else 0)
            if side=='header':stats.headers+=1
            else:stats.footers+=1
        if not profile[side]:_tiny(container.add_paragraph())
    if counter:
        old=section._sectPr.find(qn('w:pgNumType'))
        if old is not None:section._sectPr.remove(old)
        section._sectPr.append(_element('w:pgNumType',start=counter['number'],fmt=counter['fmt']))


def _toc_marker(p, begin=False, instruction=None, end=False):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    if begin:
        p.add_run()._r.append(_element('w:fldChar',fldCharType='begin'))
        field=OxmlElement('w:instrText');field.set(qn('xml:space'),'preserve');field.text=' '+instruction+' '
        p.add_run()._r.append(field);p.add_run()._r.append(_element('w:fldChar',fldCharType='separate'))
    if end:p.add_run()._r.append(_element('w:fldChar',fldCharType='end'))


def _picture_regions(graphics, images, width, height):
    """Unify semantic figures and raw image coverage before raster cropping.

    A crop already contains every source object in its rectangle. Emitting a
    second, near-contained crop duplicates the picture even if the PDF tags
    report a slightly different bounding box. Union preserves both bounds.
    Separate placements of the same image remain separate objects.
    """
    boxes = [g.box for g in graphics if g.box]
    boxes.extend((i['x0'], i['top'], i['x1'], i['bottom']) for i in images)
    regions = []
    for box in boxes:
        box = (max(0, box[0]), max(0, box[1]), min(width, box[2]), min(height, box[3]))
        if box[2]-box[0] < .1 or box[3]-box[1] < .1:
            continue
        changed = True
        while changed:
            changed = False
            for index, other in enumerate(regions):
                overlap = max(0, min(box[2],other[2])-max(box[0],other[0])) * max(0, min(box[3],other[3])-max(box[1],other[1]))
                area = min((box[2]-box[0])*(box[3]-box[1]), (other[2]-other[0])*(other[3]-other[1]))
                if overlap >= area * .97:
                    box = (min(box[0],other[0]),min(box[1],other[1]),max(box[2],other[2]),max(box[3],other[3]))
                    regions.pop(index)
                    changed = True
                    break
        regions.append(box)
    return regions


def _footer_number(char, page_height):
    """A bare number in the footer strip is a page number, not body text.

    The rebuild positions paragraphs absolutely, so a footer number that leaks
    into the body stream lands inside the last text line ("即判 66 断…").
    """
    text=(char.get('text') or '').strip()
    return bool(text) and text.isdigit() and char.get('top',0) > page_height-45


def rebuild_pdf_docx_fidelity(source, generated, pages, check_cancel, progress=None, *, selection=None):
    import pdfplumber
    import pypdfium2 as pdfium
    from docx import Document
    from docx.shared import Pt
    from docx.enum.section import WD_SECTION
    from docx.enum.style import WD_STYLE_TYPE
    from docx.oxml.ns import qn
    stats=FidelityStats();document=Document()
    normal=document.styles['Normal'];normal.font.name='宋体';normal.font.size=Pt(12)
    normal.paragraph_format.space_after=Pt(0);normal.paragraph_format.widow_control=False
    settings=document.settings.element
    compat=settings.find(qn('w:compat'))
    if compat is None:
        compat=_element('w:compat');settings.append(compat)
    compat.append(_element('w:doNotExpandShiftReturn'))
    for name in ('Ruyi Figure Caption','Ruyi Table Caption'):
        if name not in document.styles:document.styles.add_style(name,WD_STYLE_TYPE.PARAGRAPH)
    pdf=pdfium.PdfDocument(source)
    try:
        with pdfplumber.open(source) as extracted:
            indexed=Source(extracted,check_cancel)
            blocks=recover_directories(indexed.blocks(),extracted.pages);unit=_tag_scale(blocks)
            for block in blocks:
                if block.chars:
                    height=extracted.pages[block.page-1].height
                    block.chars=[c for c in block.chars if not _footer_number(c,height)]
            expanded=[]
            for block in blocks:
                if block.kind!='table':expanded.append(block);continue
                figures=[n for n in _walk(block.element) if n.type=='Figure']
                table_graphics=[indexed.graphic(n) for n in _walk(block.element) if n.type in ('Figure','Formula')]
                bounds=block.box
                page=extracted.pages[block.page-1]
                has_outer_rules=False
                if bounds:
                    x0,y0,x1,y1=bounds
                    has_outer_rules=any((r[0]=='v' and min(abs(r[1]-x0),abs(r[1]-x1))<1.2 and r[3]-r[2]>(y1-y0)*.5) or
                                        (r[0]=='h' and min(abs(r[1]-y0),abs(r[1]-y1))<1.2 and r[3]-r[2]>(x1-x0)*.8)
                                        for r in _rules(page))
                # PDF authors often use an invisible table only to arrange
                # several charts. Its chart labels are not spreadsheet cells.
                # Keep complete figures and normal caption paragraphs instead
                # of forcing those labels into fixed-height native rows.
                layoutlike=len(figures)>1 or any(g.box and g.box[2]-g.box[0]>80 and g.box[3]-g.box[1]>80
                                                 for g in table_graphics if g.element.type=='Figure')
                if figures and layoutlike and not has_outer_rules:
                    expanded.extend(table_graphics)
                    used=set()
                    for node in _walk(block.element):
                        if node.type not in ('P','Caption','Code','H1','H2','H3','H4','H5','H6'):continue
                        chars=[c for c in indexed.element_chars(node,exclude_graphics=True) if c['_flow_key'] not in used]
                        used.update(c['_flow_key'] for c in chars)
                        if chars:
                            kind=node.type
                            if re.match(r'^\s*[图表]\s*\d+',''.join(c.get('text','') for c in chars)):kind='Caption'
                            expanded.append(Block(kind,chars,node,chars[0]['page_number'],_box(chars),tagged=True))
                else:
                    expanded.append(block);expanded.extend(table_graphics)
            blocks=expanded
            selected=list(selection) if selection else list(range(1,pages+1))
            bypage=defaultdict(list);graphics=defaultdict(list);taggedtables=[];tocdata={}
            bookmarkid=0
            for block in blocks:
                if re.fullmatch(r'H[1-6]',block.kind):block.level=int(block.kind[-1])
                if block.kind=='table':taggedtables.append(block);continue
                if block.kind in ('graphic','scan','floating_photo'):
                    graphics[block.page].append(block);continue
                if block.kind=='ruled_table':continue
                if block.kind=='blank':continue
                groups=block.entries if block.kind in ('toc','directory') else [block.chars]
                if block.kind=='toc':
                    labels=[''.join(c.get('text','') for c in cc).strip() for cc in groups]
                    fig=sum(bool(re.match(r'^图\s*\d+',s)) for s in labels)
                    tab=sum(bool(re.match(r'^表\s*\d+',s)) for s in labels)
                    instruction='TOC \\h \\z \\t "Ruyi Figure Caption,1"' if fig>len(labels)*.6 else \
                                'TOC \\h \\z \\t "Ruyi Table Caption,1"' if tab>len(labels)*.6 else 'TOC \\o "1-6" \\h \\z \\u'
                    tocid=len(tocdata);tocdata[tocid]=instruction
                for g,chars in enumerate(groups):
                    pagegroups=defaultdict(list)
                    for c in chars:pagegroups[c['page_number']].append(c)
                    for number,cc in pagegroups.items():
                        for part,lines in enumerate(_groups(cc,extracted.pages[number-1].height)):
                            bookmarkid+=1
                            info={'kind':block.kind,'lines':lines,'block':block,'id':bookmarkid}
                            if block.kind=='toc':
                                info.update(toc=tocid,first=g==0 and part==0,last=g==len(groups)-1)
                                stats.toc_entries+=1
                            elif block.kind=='directory' and number in selected:
                                stats.directory_entries+=1
                            bypage[number].append(info)
                for graphic in block.parts:graphics[graphic.page].append(graphic)
            previous_tables=[]
            for output_index,number in enumerate(selected):
                check_cancel();page=extracted.pages[number-1]
                section=document.add_section(WD_SECTION.NEW_PAGE) if output_index else document.sections[0]
                if output_index:
                    _tiny(document.paragraphs[-1])
                _page_profile(section,page,indexed.profiles[number],stats)
                anchor=document.add_paragraph();_tiny(anchor)
                pdfpage=pdf[number-1];pictures=_PagePictures(pdfpage,page.width,page.height,
                    render_scale=min(300/72,4000/max(page.width,page.height)))
                try:
                    _annotate_bold(page,pdfpage)
                    tables=[]
                    for block in taggedtables:
                        if any(c['page_number']==number for c in block.chars):
                            model=_tag_table(indexed,block,number,unit)
                            if model:
                                graphic_keys={c['_flow_key'] for node in _walk(block.element) if node.type in ('Figure','Formula')
                                              for c in indexed.element_chars(node)}
                                for cell in model.cells:cell.chars=[c for c in cell.chars if c['_flow_key'] not in graphic_keys]
                                tables.append(model)
                    for model in find_ruled_tables(page):
                        # Chart axes/gridlines also form rectangles. A ruled
                        # candidate inside an already identified figure is
                        # part of that figure, not a new native data table.
                        if any(g.box and g.box[0]-1<=model.bbox[0] and g.box[1]-1<=model.bbox[1] and
                               g.box[2]+1>=model.bbox[2] and g.box[3]+1>=model.bbox[3]
                               for g in graphics[number]):continue
                        if not any(_inside({'x0':model.bbox[0],'x1':model.bbox[2],'top':model.bbox[1],'bottom':model.bbox[3]},t.bbox) for t in tables):
                            tables.append(_ruled_table(page,model))
                    # Vector helper produces ordinary body anchor paragraphs;
                    # consolidate them into one tiny slot to avoid cumulative
                    # blank lines displacing source-aligned editable paragraphs.
                    scan=next((g for g in graphics[number] if g.kind=='scan'),None)
                    regions = [(0,0,page.width,page.height)] if scan else _picture_regions(
                        graphics[number],page.images,page.width,page.height)
                    profile=indexed.profiles[number]
                    body_top=max((item['box'][3]+1 for item in profile['header']),default=0)
                    body_bottom=min((item['box'][1]-1 for item in profile['footer']),default=page.height)
                    open_tables=captioned_table_regions(page,bypage[number],_rules(page),
                        graphics=graphics[number],body_top=body_top,body_bottom=body_bottom) if not scan else []
                    if not _RASTER_TABLES and not scan:
                        tables.extend(open_table_models(page,open_tables,_rules(page),Table,Cell,
                            existing=tables,previous=previous_tables,body_bottom=body_bottom,graphics=graphics[number]))
                        if tables:previous_tables=tables
                        # WPS reserves a small strip around page-footer frames.
                        # Compress only table row whitespace, at most six per
                        # cent, rather than letting the last row jump below the
                        # footer or forcing another page. Font sizes stay intact.
                        if profile['footer']:
                            safe_bottom=min(v['box'][1] for v in profile['footer'])-12
                            for table in tables:
                                height=table.bbox[3]-table.bbox[1]
                                if height>0 and table.bbox[3]>safe_bottom:
                                    table.fit_scale=max(.94,min(1,(safe_bottom-table.bbox[1])/height))
                    bypage[number]=_integrate_native_scripts(bypage[number],page,tables)
                    protected=([t.bbox for t in tables] if not _RASTER_TABLES else [])+[
                        _box([c for line in info['lines'] for c in line]) for info in bypage[number]
                        if info['kind'] in ('toc','directory')]+regions
                    inline_ids=classify_inline_math(page,page.chars,_needs_math_raster,_baseline) if not scan else set()
                    display=equation_regions(page,[],lambda c:_needs_math_raster(c) and id(c) not in inline_ids,
                        _baseline,pictures,body_top=body_top,body_bottom=body_bottom,exclude=protected,
                        include_scripts=False,clip_prose=True) if _RASTER_FORMULAS and not scan else []
                    if display and protected:
                        seeds=[c for c in page.chars if _needs_math_raster(c) and id(c) not in inline_ids]
                        display=exclude_formula_overlap(display,protected,seeds,_baseline,page,pictures)
                    fragments=inline_formula_fragments(page,page.chars,lambda c:id(c) in inline_ids,_baseline,pictures,
                        body_top=body_top,body_bottom=body_bottom,exclude=protected+display) if _RASTER_FORMULAS and not scan else []
                    bypage[number]=_integrate_formula_fragments(bypage[number],fragments,page)
                    equations=[p['_ruyi_formula_box'] for p,_ in fragments]
                    special=display+(open_tables if _RASTER_TABLES else [])
                    if _RASTER_TABLES and not scan:
                        table_boxes=[(max(0,t.bbox[0]-1),max(0,t.bbox[1]-1),min(page.width,t.bbox[2]+1),min(page.height,t.bbox[3]+1)) for t in tables]
                        special.extend(table_boxes)
                        stats.raster_tables+=len(merge_boxes(table_boxes+open_tables))
                    stats.formula_regions+=len(equations)+len(display)
                    stats.images+=len(equations)
                    if special:regions=merge_boxes(regions+special)
                    start=len(document.paragraphs)
                    _add_vectors(document,page,tables,1,pictures,stats,covered_regions=regions+equations)
                    for p in document.paragraphs[start:]:
                        for run in list(p._p):
                            if run.tag.endswith('}r'):anchor._p.append(run)
                        p._p.getparent().remove(p._p)
                    if scan:
                        # An image-only cover remains at its original page
                        # coordinates and size, without body-margin shrinkage.
                        box=(0,0,page.width,page.height)
                        _anchor_picture(anchor,pictures.crop(box),box,1)
                        stats.images+=1;stats.scanned_pages+=1
                    else:
                        for box in regions:
                            _anchor_picture(anchor,pictures.crop(box),box,1);stats.images+=1
                    cursor=.05
                    page_lines=_lines([c for c in page.chars if c['_flow_key'] not in indexed.excluded])
                    page_deltas=[_base(b,page.height)-_base(a,page.height) for a,b in zip(page_lines,page_lines[1:])]
                    typical=median(d for d in page_deltas if 8<d<40) if any(8<d<40 for d in page_deltas) else 18
                    # Page-anchored crops do not consume Word body height. Keep
                    # the flow cursor at the last actual paragraph; the next
                    # paragraph's measured space_before reserves the entire gap.
                    events=[] if _RASTER_TABLES else [(t.bbox[1],'table',t) for t in tables]
                    for item in bypage[number]:
                        chunks=[];previous=-2
                        for line_index,line in enumerate(item['lines']):
                            kept=[c for c in line if not any(_inside(c,t.bbox) for t in tables)
                                  and not any(box[0]-.1<=(c['x0']+c['x1'])/2<=box[2]+.1 and
                                              box[1]-.1<=(_base([c],page.height)-c.get('size',12)*.4)<=box[3]+.1 for box in special)
                                  ]
                            if not kept:continue
                            if line_index!=previous+1:chunks.append([])
                            chunks[-1].append(kept);previous=line_index
                        for chunk_index,chunk in enumerate(chunks):
                            part=dict(item,lines=chunk)
                            if item.get('first'):part['first']=chunk_index==0
                            if item.get('last'):part['last']=chunk_index==len(chunks)-1
                            events.append((_base(chunk[0],page.height),'text',part))
                    for _,kind,info in sorted(events,key=lambda e:e[0]):
                        if kind=='table':
                            gap=info.bbox[1]-cursor
                            if gap>.025:
                                spacer=document.add_paragraph();_tiny(spacer)
                                spacer.paragraph_format.line_spacing=Pt(gap)
                            if not _RASTER_TABLES:
                                _add_table(document,page,info,stats,pictures)
                            cursor=max(cursor,info.bbox[1])+(info.bbox[3]-info.bbox[1])*info.fit_scale
                            continue
                        lines=info['lines'];block=info['block']
                        # The event builder removes covered text and splits a
                        # paragraph at every raster gap before measuring leading.
                        if not lines:continue
                        style='Heading '+str(block.level) if block.kind.startswith('H') and block.level else \
                              'Ruyi Figure Caption' if block.kind=='Caption' and re.match(r'^\s*图\s*\d+',''.join(c['text'] for c in lines[0])) else \
                              'Ruyi Table Caption' if block.kind=='Caption' and re.match(r'^\s*表\s*\d+',''.join(c['text'] for c in lines[0])) else None
                        single=max(c.get('size',12) for c in lines[0])*1.15
                        p=document.add_paragraph(style=style);leading=_leading(lines,min(typical,single))
                        target=_base(lines[0],page.height)-leading*.8
                        positioned=target<cursor-.5
                        if positioned:
                            left=min(c['x0'] for line in lines for c in line)
                            right=max(c['x1'] for line in lines for c in line)
                            p.paragraph_format.space_before=Pt(0)
                            height=leading*len(lines)
                            # The rebuild sets page margins to zero and positions
                            # paragraphs absolutely, so a frame near the bottom edge
                            # lands on the footer page number as soon as the hosting
                            # application reflows by a point or two. Keep frames out
                            # of the bottom strip instead of trusting the source y.
                            top=max(0,min(target,page.height-36-height))
                            p._p.get_or_add_pPr().append(_element('w:framePr',w=round((right-left+10)*20),
                                h=round(height*20),x=round(left*20),y=round(top*20),
                                hAnchor='page',vAnchor='page',hRule='atLeast',wrap='none',hSpace=0,vSpace=0))
                            stats.positioned_paragraphs+=1
                        else:
                            p.paragraph_format.space_before=Pt(max(0,target-cursor))
                        if info.get('first') and selection is None:_toc_marker(p,begin=True,instruction=tocdata[info['toc']])
                        if info['kind']=='directory':
                            _write_directory(p,lines,page,leading,origin=left if positioned else 0,positioned=positioned)
                        else:
                            _write_lines(p,lines,page,leading,origin=left if positioned else 0,cell=positioned,pictures=pictures)
                        if info.get('last') and selection is None:_toc_marker(p,end=True)
                        if style and style.startswith('Heading'):
                            _bookmark(p,'Ruyi_'+str(info['id']),info['id']);stats.headings+=1
                        if not positioned:cursor=max(cursor,target)+leading*len(lines)
                        stats.paragraphs+=1
                    if progress:progress(output_index+1,len(selected))
                finally:
                    pictures.close();pdfpage.close()
            check_cancel();document.save(generated)
    finally:pdf.close()
    return stats
