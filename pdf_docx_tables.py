"""Recover open-rule table cells from visible rules and text gutters.

The table model is native Word data, not a screenshot. Titles and prose stay
outside the model. Every visible source character is assigned to one cell.
"""
from collections import Counter
from statistics import median
import re
from pdf_docx_layout import _lines,_baseline,_cluster


def _y(c,height):
    return height+_baseline(c)-c.get('size',12)*.4


def _segments(chars):
    result=[]
    for line in _lines(chars):
        segment=[]
        for c in line:
            if segment and c['x0']-segment[-1]['x1']>max(3,median(v.get('size',10) for v in segment)*.45):
                result.append(segment);segment=[]
            segment.append(c)
        if segment:result.append(segment)
    return result


def _columns(chars,left,right,header_end,height):
    body=[c for c in chars if _y(c,height)>header_end+.5]
    segments=_segments(body or chars)
    # Continuous whitespace across body rows is strong column evidence. Work
    # with whole segments, so letter spacing within one word is not a column.
    intervals=sorted((min(c['x0'] for c in s)-1,max(c['x1'] for c in s)+1) for s in segments)
    joined=[]
    for a,b in intervals:
        if joined and a<=joined[-1][1]:joined[-1]=(joined[-1][0],max(joined[-1][1],b))
        else:joined.append((a,b))
    cuts=[(a[1]+b[0])/2 for a,b in zip(joined,joined[1:]) if b[0]-a[1]>=4]
    # Long merged notes can bridge otherwise persistent gutters. Use repeated
    # segment gaps in that case, accepting only gaps crossed by very few rows.
    if not cuts and len(segments)>=4:
        lines=_lines(body or chars);gaps=[]
        for line in lines:
            for a,b in zip(line,line[1:]):
                if b['x0']-a['x1']>=8:gaps.append((a['x1']+b['x0'])/2)
        for x in _cluster(gaps,tolerance=4):
            support=sum(any(a['x1']+1<x<b['x0']-1 for a,b in zip(line,line[1:])) for line in lines)
            crossing=sum(min(c['x0'] for c in s)<x<max(c['x1'] for c in s) for s in segments)
            if support>=2 and crossing<=max(1,len(lines)*.12):cuts.append(x)
    return [left]+[x for x in sorted(cuts) if left+8<x<right-8]+[right]


def _row_edges(chars,xs,top,bottom,header_end,height):
    data=[c for c in chars if header_end+.25<_y(c,height)<bottom]
    if not data:return [top,bottom]
    bycol=[]
    for lo,hi in zip(xs,xs[1:]):
        cc=[c for c in data if lo<=(c['x0']+c['x1'])/2<hi]
        bycol.append(_lines(cc))
    # A short label/numeric column gives one anchor per logical row. Unlike
    # midpoint-only row division, choose the actual blank stripe between
    # labels, which handles adjacent rows with very different text heights.
    populated=[(i,lines) for i,lines in enumerate(bycol) if lines]
    def data_key(line):
        text=''.join(c.get('text','') for c in line).strip()
        return bool(re.fullmatch(r'[-+]?\d*(?:[.,]\d+)?(?:[eE][-+]?\d+)?%?|[A-Za-z]{1,5}[_-]?\d+|_cons',text))
    numeric=[item for item in populated if sum(data_key(line) for line in item[1])>=len(item[1])*.7]
    _,anchors=max(numeric,key=lambda item:(len(item[1]),-item[0])) if numeric else min(populated,key=lambda item:(len(item[1]),item[0]))
    centers=[median(_y(c,height) for c in line) for line in anchors]
    ink=sorted((c['top'],c['bottom']) for c in data)
    bands=[]
    for a,b in ink:
        if bands and a<=bands[-1][1]:bands[-1]=(bands[-1][0],max(bands[-1][1],b))
        else:bands.append((a,b))
    whitespace=[(a[1],b[0]) for a,b in zip(bands,bands[1:]) if b[0]>a[1]+.3]
    edges=[header_end]
    for a,b in zip(centers,centers[1:]):
        options=[(lo,hi) for lo,hi in whitespace if a<lo<hi<b]
        edge=sum(max(options,key=lambda g:g[1]-g[0]))/2 if options else (a+b)/2
        if edge>edges[-1]+2:edges.append(edge)
    edges.append(bottom)
    if header_end>top+.5:edges.insert(0,top)
    return edges


def open_table_models(page,regions,rules,Table,Cell,*,existing=(),previous=(),body_bottom=None,graphics=()):
    result=[]
    regions=list(regions);single_rule_regions=set()
    # A matrix may continue with no top rule on the next page. The previous
    # page's confirmed grid plus several aligned rows and one closing rule
    # provide evidence; a lone underline below prose still does not.
    for model in previous:
        if model.page!=page.page_number-1 or model.bbox[3]<page.height*.8:continue
        candidates=[r for r in rules if r[0]=='h' and abs(r[2]-model.bbox[0])<3 and
                    abs(r[3]-model.bbox[2])<3 and 80<r[1]<(body_bottom or page.height)]
        if not candidates:continue
        closing=min(candidates,key=lambda r:r[1])
        chars=[c for c in page.chars if c.get('text','').strip() and model.bbox[0]<=(c['x0']+c['x1'])/2<=model.bbox[2]
               and 0<_y(c,page.height)<closing[1]]
        lines=_lines(chars)
        if len(lines)<3 or min(c['top'] for c in chars)>page.height*.13:continue
        aligned=sum(len(_segments(line))>=2 for line in lines)
        if aligned<3 or aligned<len(lines)*.4:continue
        if any(len(''.join(c.get('text','') for c in line))>25 and
               max(c['x1'] for c in line)-min(c['x0'] for c in line)>(model.bbox[2]-model.bbox[0])*.7 for line in lines):continue
        box=(model.bbox[0]-1,max(0,min(c['top'] for c in chars)-1),model.bbox[2]+1,closing[1]+1)
        if any(g.box and g.box[0]-2<=box[0] and g.box[1]-2<=box[1] and g.box[2]+2>=box[2] and g.box[3]+2>=box[3] for g in graphics):continue
        regions.append(box);single_rule_regions.add(box)
    for region in regions:
        horizontal=sorted([r for r in rules if r[0]=='h' and region[1]-.1<=r[1]<=region[3]+.1
                           and abs(r[2]-region[0])<3 and abs(r[3]-region[2])<3],key=lambda r:r[1])
        if region in single_rule_regions and len(horizontal)==1:
            closing=horizontal[0]
            horizontal.insert(0,('h',region[1]+1,closing[2],closing[3],0,'000000',None))
        if len(horizontal)<2:continue
        first,last=horizontal[0],horizontal[-1]
        bbox=(first[2],first[1],first[3],last[1])
        # A caption followed by only two rules can be an opening header whose
        # body continues to the page bottom, not a complete two-rule table.
        # Require separately aligned cells after the second rule; normal prose
        # under a finished table is not enough evidence to extend its bounds.
        opening=len(horizontal)==2 and region[1]<first[1]-2 and last[1]-first[1]<90
        if opening:
            below=[c for c in page.chars if c.get('text','').strip() and
                   bbox[0]<=(c['x0']+c['x1'])/2<=bbox[2] and
                   last[1]+.5<_y(c,page.height)<min(body_bottom or page.height,page.height-55)]
            accepted=[];aligned=0;waiting=0;previous_y=last[1]
            for line in _lines(below):
                size=median(c.get('size',12) for c in line)
                text=''.join(c.get('text','') for c in line).strip()
                width=max(c['x1'] for c in line)-min(c['x0'] for c in line)
                y=median(_y(c,page.height) for c in line)
                if re.match(r'^(?:表\s*\d|Table\s+\d|注[:：])',text,re.I):break
                if width>(bbox[2]-bbox[0])*.7 and len(text)>25:break
                if y-previous_y>max(35,size*3):break
                if len(_segments(line))>=2:aligned+=1
                elif not aligned:
                    waiting+=1
                    if waiting>3 or len(text)>16:break
                accepted.extend(line);previous_y=y
            if aligned and accepted:bbox=(bbox[0],bbox[1],bbox[2],max(c['bottom'] for c in accepted)+2)
        if any(min(bbox[2],m.bbox[2])-max(bbox[0],m.bbox[0])>min(bbox[2]-bbox[0],m.bbox[2]-m.bbox[0])*.8
               and min(bbox[3],m.bbox[3])-max(bbox[1],m.bbox[1])>min(bbox[3]-bbox[1],m.bbox[3]-m.bbox[1])*.8 for m in tuple(existing)+tuple(result)):continue
        chars=[c for c in page.chars if c.get('text','').strip() and bbox[0]-.1<=(c['x0']+c['x1'])/2<=bbox[2]+.1
               and bbox[1]< _y(c,page.height)<bbox[3]]
        if len(chars)<4:continue
        header_end=horizontal[1][1] if len(horizontal)>=3 or opening else bbox[1]
        xs=_columns(chars,bbox[0],bbox[2],header_end,page.height)
        continuation=bbox[1]<page.height*.13 and region[1]>=first[1]-2
        if continuation:
            match=next((m for m in reversed(previous) if abs(m.bbox[0]-bbox[0])<3 and abs(m.bbox[2]-bbox[2])<3),None)
            if match:xs=list(match.xs)
        if len(xs)>40:continue
        ys=_row_edges(chars,xs,bbox[1],bbox[3],header_end,page.height)
        cells=[]
        for row,(lo,hi) in enumerate(zip(ys,ys[1:])):
            rowchars=[c for c in chars if lo<=_y(c,page.height)<hi]
            spans=[]
            for segment in _segments(rowchars):
                columns=[j for j,(left,right) in enumerate(zip(xs,xs[1:])) if
                         any(left<=(c['x0']+c['x1'])/2<right for c in segment)]
                if len(columns)>1:spans.append((min(columns),max(columns)+1))
            groups=[]
            for a,b in sorted(spans):
                if groups and a<groups[-1][1]:groups[-1]=(groups[-1][0],max(groups[-1][1],b))
                else:groups.append((a,b))
            col=0
            while col<len(xs)-1:
                span=next((b-a for a,b in groups if a==col),1)
                left,right=xs[col],xs[col+span]
                cc=[c for c in rowchars if left<=(c['x0']+c['x1'])/2<right]
                cells.append(Cell(row,col,1,span,(left,lo,right,hi),cc));col+=span
        assigned=Counter(id(c) for cell in cells for c in cell.chars)
        if any(assigned[id(c)]!=1 for c in chars):raise ValueError('表格内容分配不完整或重复。')
        result.append(Table(page.page_number,bbox,xs,ys,cells))
    return result
