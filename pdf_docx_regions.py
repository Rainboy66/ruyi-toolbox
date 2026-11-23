"""Conservative local appearance fallbacks for equations and open ruled tables.

Math font boxes can be wrong even when the PDF renders correctly. Use text
origins as seeds, include surrounding physical paragraphs, and snap to ink.
"""
import re
from statistics import median


def merge_boxes(boxes, tolerance=.1):
    result=[]
    for box in boxes:
        if box[2]<=box[0] or box[3]<=box[1]:continue
        box=tuple(box)
        changed=True
        while changed:
            changed=False
            for i,other in enumerate(result):
                if min(box[2],other[2])+tolerance>max(box[0],other[0]) and min(box[3],other[3])+tolerance>max(box[1],other[1]):
                    box=(min(box[0],other[0]),min(box[1],other[1]),max(box[2],other[2]),max(box[3],other[3]))
                    result.pop(i);changed=True;break
        result.append(box)
    return sorted(result,key=lambda b:(b[1],b[0]))


def exclude_formula_overlap(boxes,protected,seeds,baseline,page,pictures):
    """Subtract table, directory and picture bounds from display formula crops."""
    pieces=list(boxes)
    for blocked in protected:
        # Leave a small gap so neighboring pictures cannot be re-merged.
        bx0,by0,bx1,by1=blocked[0]-.2,blocked[1]-.2,blocked[2]+.2,blocked[3]+.2
        result=[]
        for x0,y0,x1,y1 in pieces:
            ix0,iy0,ix1,iy1=max(x0,bx0),max(y0,by0),min(x1,bx1),min(y1,by1)
            if ix0>=ix1 or iy0>=iy1:
                result.append((x0,y0,x1,y1));continue
            for part in ((x0,y0,x1,iy0),(x0,iy1,x1,y1),(x0,iy0,ix0,iy1),(ix1,iy0,x1,iy1)):
                if part[2]-part[0]>.2 and part[3]-part[1]>.2:result.append(part)
        pieces=result
    result=[]
    for part in pieces:
        if not any(part[0]<=((c['x0']+c['x1'])/2)<=part[2] and
                   part[1]<=page.height+baseline(c)-c.get('size',12)*.4<=part[3] for c in seeds):
            continue
        ink=_ink_edges(part,pictures)
        if not ink:continue
        ink=(max(part[0],ink[0]),max(part[1],ink[1]),min(part[2],ink[2]),min(part[3],ink[3]))
        if ink[2]>ink[0] and ink[3]>ink[1]:result.append(ink)
    return result


def box_of(chars):
    return (min(c['x0'] for c in chars),min(c['top'] for c in chars),max(c['x1'] for c in chars),max(c['bottom'] for c in chars))


def overlaps(a,b):
    return min(a[2],b[2])>max(a[0],b[0])+.05 and min(a[3],b[3])>max(a[1],b[1])+.05


def _ink_edges(box,pictures):
    """Find actual ink in a complete band, expanding an edge that cuts a glyph."""
    pictures.crop(box).close()  # Initializes the cached page render once.
    image=pictures.image
    sx,sy=image.width/pictures.width,image.height/pictures.height
    left=max(0,int(box[0]*sx));right=min(image.width,int(box[2]*sx+.999))
    top=max(0,int(box[1]*sy));bottom=min(image.height,int(box[3]*sy+.999))
    grey=image.crop((left,0,right,image.height)).convert('L')
    mask=grey.point(lambda value:255 if value<235 else 0)
    try:
        for _ in range(max(1,round(8*sy))):
            if top<=0 or mask.crop((0,top,mask.width,min(top+1,mask.height))).getbbox() is None:break
            top-=1
        for _ in range(max(1,round(8*sy))):
            if bottom>=mask.height or mask.crop((0,bottom-1,mask.width,bottom)).getbbox() is None:break
            bottom+=1
        bounds=mask.crop((0,top,mask.width,bottom)).getbbox()
        if not bounds:return None
        # Keep one rendered pixel of white around the original ink.
        return (left/sx,max(0,top+bounds[1]-1)/sy,right/sx,min(image.height,top+bounds[3]+1)/sy)
    finally:mask.close();grey.close()


def inline_formula_fragments(page,chars,is_math,baseline,pictures,*,exclude=(),body_top=0,body_bottom=None):
    """Identify compact math runs using glyph origins, including unmapped glyphs.

    Ordinary text between two formula runs is a hard boundary. Returned
    fragments can be inserted as inline images into an editable paragraph.
    """
    bottom=page.height if body_bottom is None else body_bottom
    def y(c):return page.height+baseline(c)
    def outside(c):
        x=(c['x0']+c['x1'])/2;at=y(c)-c.get('size',12)*.4
        return not any(b[0]<=x<=b[2] and b[1]<=at<=b[3] for b in exclude)
    usable=[c for c in chars if body_top<y(c)<bottom and outside(c)]
    seeds=[c for c in usable if is_math(c)]
    if not seeds:return []
    lines=[]
    for char in sorted(seeds,key=lambda c:(y(c),c['x0'])):
        size=max(4,min(80,char.get('size',12)))
        match=min(lines,key=lambda line:abs(median(y(c) for c in line)-y(char)),default=None)
        group_size=max(size,median(c.get('size',12) for c in match)) if match else size
        if match is None or abs(median(y(c) for c in match)-y(char))>max(3,group_size*.6):
            match=[];lines.append(match)
        match.append(char)
    fragments=[];clusters=[]
    for line in lines:
        line.sort(key=lambda c:c['x0']);groups=[]
        for char in line:
            previous=groups[-1][-1] if groups else None
            gap=char['x0']-previous['x1'] if previous else 0
            size=max(char.get('size',12),previous.get('size',12) if previous else 0)
            barrier=previous and any(not is_math(c) and c.get('text','').strip() and
                abs(y(c)-y(char))<max(3,size*.45) and
                previous['x1']-.15<=(c['x0']+c['x1'])/2<=char['x0']+.15 for c in usable)
            if previous is None or gap>max(9,size*1.5) or barrier:groups.append([])
            groups[-1].append(char)
        clusters.extend(groups)
    # Fractions stack numerator and denominator around a short rule. Rejoin
    # only with actual rule evidence; proximity alone would join adjacent lines.
    rules=[]
    for obj in list(getattr(page,'lines',[]))+list(getattr(page,'rects',[])):
        if abs(obj.get('bottom',0)-obj.get('top',0))<=2 and obj.get('x1',0)-obj.get('x0',0)>1:
            rules.append((obj['x0'],(obj['top']+obj['bottom'])/2,obj['x1']))
    changed=True
    while changed:
        changed=False
        for i,a in enumerate(clusters):
            for j in range(i+1,len(clusters)):
                b=clusters[j]; ya=median(y(c) for c in a);yb=median(y(c) for c in b)
                x0=max(min(c['x0'] for c in a),min(c['x0'] for c in b))
                x1=min(max(c['x1'] for c in a),max(c['x1'] for c in b))
                size=max(c.get('size',12) for c in a+b)
                if x1<=x0 or abs(ya-yb)>size*2.5:continue
                if any(min(ya,yb)-3<=ry<=max(ya,yb)+3 and rx0<=x1 and rx1>=x0 and
                       rx1-rx0<=max(30,(x1-x0)*5) for rx0,ry,rx1 in rules):
                    clusters[i]=a+b;clusters.pop(j);changed=True;break
            if changed:break
    for group in clusters:
            size=max(4,min(80,max(c.get('size',12) for c in group)))
            main=[c for c in group if c.get('size',12)>=size*.85] or group
            base=median(y(c) for c in main)
            x0=max(0,min(c['x0'] for c in group)-.35);x1=min(page.width,max(c['x1'] for c in group)+.35)
            ys=[y(c) for c in group]
            for rx0,ry,rx1 in rules:
                if min(ys)-3<=ry<=max(ys)+3 and rx0<=x1 and rx1>=x0 and rx1-rx0<=max(30,(x1-x0)*5):
                    x0=max(0,min(x0,rx0-.35));x1=min(page.width,max(x1,rx1+.35))
            # Neighboring prose can have ink in the same narrow x-strip. Do
            # not allow ink-edge refinement to absorb the next physical line.
            other=[y(c) for c in usable if not is_math(c) and c.get('text','').strip() and
                   abs(y(c)-base)>max(8,size*.8)]
            prev=max((v for v in other if v<base),default=None)
            nextline=min((v for v in other if v>base),default=None)
            upper=max(body_top,(prev+base)/2+1 if prev is not None else body_top)
            lower=min(bottom,(nextline+base)/2-1 if nextline is not None else bottom)
            initial=(x0,max(upper,min(ys)-size*1.15),x1,min(lower,max(ys)+size*.65))
            if initial[3]<=initial[1]:continue
            ink=_ink_edges(initial,pictures)
            if not ink:continue
            ink=(ink[0],max(upper,ink[1]),ink[2],min(lower,ink[3]))
            if ink[3]<=ink[1]:continue
            # Keep the PDF source position; text origin is more reliable than
            # the math font's frequently displaced glyph bounding box.
            placeholder={'text':'\ufffc','x0':ink[0],'x1':ink[2],'top':base-size*.85,
                'bottom':base+size*.2,'size':size,'fontname':'SimSun',
                'matrix':(1,0,0,1,ink[0],page.height-base),'page_number':page.page_number,
                '_ruyi_formula_box':ink}
            fragments.append((placeholder,group))
    return fragments


def equation_regions(page,paragraphs,is_math,baseline,pictures,*,body_top=0,body_bottom=None,exclude=(),include_scripts=True,clip_prose=False):
    bottom=page.height if body_bottom is None else body_bottom
    def outside(c):
        x=(c['x0']+c['x1'])/2;y=page.height+baseline(c)-c.get('size',12)*.4
        return not any(b[0]<=x<=b[2] and b[1]<=y<=b[3] for b in exclude)
    body=[c for c in page.chars if body_top<c.get('top',0)<bottom and outside(c)]
    seeds=[c for c in page.chars if is_math(c) and body_top<page.height+baseline(c)<bottom and outside(c)]
    sized=[c for c in body if c.get('text','').strip() and c.get('size',0)>0]
    usual=median(c['size'] for c in sized) if sized else 12
    # Formula subscripts may use the ordinary CJK font (e.g. S_i squared).
    # Detect small scripts by their size and a nearby full-size baseline, not
    # by assuming every mathematical character belongs to a math font.
    for char in sized if include_scripts else []:
        if char['size']>=usual*.72 or not re.fullmatch(r'[A-Za-z0-9]+',char.get('text','')):continue
        y=page.height+baseline(char)
        if any(other['size']>char['size']*1.3 and abs(page.height+baseline(other)-y)<usual and
               -2<char['x0']-other['x1']<usual for other in sized):seeds.append(char)
    if not seeds:return []
    left=max(0,min((c['x0'] for c in body),default=0)-2)
    right=min(page.width,max((c['x1'] for c in body),default=page.width)+2)
    boxes=[]
    for char in seeds:
        size=max(4,min(80,char.get('size',12)))
        y=page.height+baseline(char)
        boxes.append((left,max(body_top,y-size*1.5),right,min(bottom,y+size*.45)))
    regions=merge_boxes(boxes)
    # Complete a mixed physical line, not its entire inferred paragraph.
    # Untagged producers can combine a whole directory/page into one P block.
    # Expanding to that block would rasterize unrelated editable content.
    text_boxes=[]
    for info in paragraphs:
        for line in info['lines']:
            cc=[c for c in line if c.get('text','').strip() and not is_math(c) and outside(c)]
            if cc:text_boxes.append(box_of(cc))
    changed=True
    while changed:
        previous_regions=list(regions)
        expanded=[]
        for region in regions:
            matches=[b for b in text_boxes if overlaps(region,b)]
            if matches:
                region=(left,max(body_top,min(region[1],*(b[1] for b in matches))-1),right,
                        min(bottom,max(region[3],*(b[3] for b in matches))+1))
            expanded.append(region)
        regions=merge_boxes(expanded)
        # A source glyph box may cross the permitted body band. Reaching that
        # boundary is a fixed point, even if the source box extends beyond it.
        # Stop on actual geometry stability instead of endlessly retrying it.
        changed=regions!=previous_regions
    output=[]
    prose=[page.height+baseline(c) for c in body if not is_math(c) and
           re.search(r'[\u3400-\u9fff]',c.get('text',''))]
    for region in regions:
        box=_ink_edges(region,pictures)
        if not box:continue
        if clip_prose:
            matching=[c for c in seeds if region[0]<=c['x0']<=region[2] and
                      region[1]-1<=page.height+baseline(c)<=region[3]+1]
            if matching:
                first=min(page.height+baseline(c) for c in matching)
                last=max(page.height+baseline(c) for c in matching)
                size=max(c.get('size',12) for c in matching)
                before=max((v for v in prose if v<first-size*.4),default=None)
                after=min((v for v in prose if v>last+size*.4),default=None)
                top=max(box[1],(before+first)/2-2 if before is not None else body_top)
                clipped_bottom=min(box[3],(after+last)/2+2 if after is not None else bottom)
                box=(box[0],top,box[2],clipped_bottom)
        if box[3]>box[1]:output.append(box)
    return output


def classify_inline_math(page,chars,is_math,baseline):
    """Distinguish formula glyphs embedded in prose from display equations."""
    def y(c):return page.height+baseline(c)
    regular=[c for c in chars if not is_math(c) and c.get('text','')]
    result=set()
    for char in chars:
        if not is_math(char):continue
        near=[c for c in regular if abs(y(c)-y(char))<=max(10.5,char.get('size',12)*.8)]
        cjk=sum(bool(re.search(r'[\u3400-\u9fff]',c.get('text',''))) for c in near)
        ordered=''.join(c.get('text','') for c in sorted(near,key=lambda c:c['x0']))
        # Whitespace can be omitted in extraction. A clearly readable English
        # sentence needs several long word runs, not a function named Prob.
        english=len(re.findall(r'[A-Za-z]{3,}',ordered))>=2 and len(ordered)>=14 and ' ' in ordered
        if cjk>=2 or english:result.add(id(char))
    return result


def captioned_table_regions(page,paragraphs,rules,*,graphics=(),body_top=0,body_bottom=None):
    """Recognize open two/three-line tables using captions and aligned rules.

    Unlike rectangle grid detection, this also accepts tables without vertical
    borders. Chart gridlines alone are insufficient evidence.
    """
    bottom=page.height if body_bottom is None else body_bottom
    captions=[]
    for info in paragraphs:
        for line in info['lines']:
            text=''.join(c.get('text','') for c in line).strip()
            if re.match(r'^(?:表\s*\d|Table\s+\d)',text,re.I):
                box=box_of(line)
                if abs((box[0]+box[2])/2-page.width/2)<page.width*.22 and len(text)<110:
                    captions.append(box)
    captions.sort(key=lambda b:b[1])
    horizontal=[r for r in rules if r[0]=='h' and r[3]-r[2]>page.width*.3 and body_top<r[1]<bottom]
    regions=[]
    for i,caption in enumerate(captions):
        limit=min(bottom,captions[i+1][1]-2 if i+1<len(captions) else bottom)
        starts=[r for r in horizontal if caption[1]-5<=r[1]<=caption[3]+24]
        if not starts:continue
        first=min(starts,key=lambda r:abs(r[1]-caption[3]))
        lines=[r for r in horizontal if first[1]-.2<=r[1]<limit and abs(r[2]-first[2])<2 and abs(r[3]-first[3])<2]
        lines.sort(key=lambda r:r[1])
        if len(lines)<2 or lines[-1][1]-lines[0][1]<12:continue
        top=lines[0][1];last=lines[-1][1]
        table=(max(0,first[2]-1),max(0,min(caption[1]-1,top-1)),min(page.width,first[3]+1),last+1)
        if any(g.box and g.box[0]-2<=table[0] and g.box[1]-2<=table[1] and g.box[2]+2>=table[2] and g.box[3]+2>=table[3] for g in graphics):continue
        # At least two horizontally separate cells below the top rule.
        content=[c for c in page.chars if c.get('text','').strip() and top<c['top']<last and first[2]<=c['x0']<first[3]]
        if len(content)<4:continue
        regions.append(table)
    # Continued tables may have no caption on this page. Three aligned rules
    # with a short header row, or a short open continuation at the page top,
    # supply independent evidence; a lone underline is still insufficient.
    groups=[]
    for rule in sorted(horizontal,key=lambda r:r[1]):
        group=next((g for g in groups if abs(g[0][2]-rule[2])<2 and abs(g[0][3]-rule[3])<2),None)
        if group is None:groups.append([rule])
        else:group.append(rule)
    # A new caption starts a new table even when both tables share exactly
    # the same left/right rules. Never union the page-top continuation with
    # the later table across an intervening prose paragraph.
    separated=[]
    for lines in groups:
        current=[]
        for line in lines:
            if current and any(current[-1][1]<c[1]<line[1] for c in captions):
                separated.append(current);current=[]
            current.append(line)
        if current:separated.append(current)
    for lines in separated:
        if len(lines)<2:continue
        first,last=lines[0],lines[-1]
        height=last[1]-first[1]
        evidence=(len(lines)>=3 and lines[1][1]-first[1]<=90 and 12<height<page.height*.85) or (
            len(lines)==2 and first[1]<page.height*.13 and 12<height<page.height*.35)
        if not evidence and len(lines)==2 and first[1]<page.height*.13 and 12<height<page.height*.85:
            from pdf_docx_layout import _lines
            cc=[c for c in page.chars if c.get('text','').strip() and first[1]<c['top']<last[1]
                and first[2]<=c['x0']<first[3]]
            physical=_lines(cc)
            aligned=sum(any(b['x0']-a['x1']>max(12,a.get('size',12)*.9) for a,b in zip(line,line[1:])) for line in physical)
            evidence=aligned>=3 and aligned>=len(physical)*.4
        if not evidence:continue
        box=(max(0,first[2]-1),first[1]-1,min(page.width,first[3]+1),last[1]+1)
        if any(g.box and g.box[0]-2<=box[0] and g.box[1]-2<=box[1] and g.box[2]+2>=box[2] and g.box[3]+2>=box[3] for g in graphics):continue
        content=[c for c in page.chars if c.get('text','').strip() and first[1]<c['top']<last[1] and first[2]<=c['x0']<first[3]]
        if len(content)>=4:regions.append(box)
    return merge_boxes(regions)
