"""Execute OutputPlan.rules with vector merges. Intermediate PDFs stay in memory."""
from dataclasses import replace
from io import BytesIO
import math
from .model import PageRef,Sheet,SourceInfo,page_geometry

A4=(595.275591,841.889764)


def stopped(cancel):
    if cancel and cancel.is_set():raise InterruptedError('已停止导出。')


def reader_of(writer):
    from pypdf import PdfReader
    stream=BytesIO();writer.write(stream);stream.seek(0)
    return PdfReader(stream)


def physical_size(page):
    unit=float(page.get('/UserUnit',1));w,h=float(page.cropbox.width)*unit,float(page.cropbox.height)*unit
    return (h,w) if page.rotation%180 else (w,h)


def blank(writer,size):
    from pypdf.generic import NameObject,NumberObject
    unit=max(1,math.ceil(max(size)/14400))
    if unit>75000:raise ValueError('长页超出PDF可表达尺寸。')
    writer._header=b'%PDF-1.7'
    page=writer.add_blank_page(width=size[0]/unit,height=size[1]/unit)
    if unit>1:page[NameObject('/UserUnit')]=NumberObject(unit)
    return page


def place(dest,source,target,size,mode='fit',over=True,interactive=True):
    from pypdf import PageObject,Transformation
    from pypdf.generic import NameObject,RectangleObject,ArrayObject
    from .exporter import layout,_annotations
    raw=tuple(map(float,source.cropbox))+(source.rotation%360,float(source.get('/UserUnit',1)))
    info=SourceInfo('internal','','',0,0,(raw,))
    unit=float(dest.get('/UserUnit',1));target=tuple(v/unit for v in target);size=tuple(v/unit for v in size)
    _,_,clip,op=layout(PageRef('internal',0,fit_mode=mode,target_rect=target),info,Sheet(size))
    # Intersect the source crop with inverse(target), preventing fill spill into neighbours.
    a,b,c,d,e,f=op.ctm;det=a*d-b*c
    inverse=Transformation((d/det,-b/det,-c/det,a/det,(c*f-d*e)/det,(b*e-a*f)/det))
    x,y,r,t=target;points=[inverse.apply_on(p) for p in ((x,y),(x,t),(r,y),(r,t))]
    clipped=(max(clip[0],min(p[0] for p in points)),max(clip[1],min(p[1] for p in points)),
             min(clip[2],max(p[0] for p in points)),min(clip[3],max(p[1] for p in points)))
    page=PageObject(source.pdf);page.update(source);page.cropbox=RectangleObject(clipped)
    annots=[]
    for ref in source.get('/Annots',[]) if interactive else ():
        obj=ref.get_object()
        if obj.get('/Subtype')=='/Widget':raise ValueError('交互表单暂不支持多页拼版；请保留原文件并使用普通整理导出。')
        if '/Dest' in obj or obj.get('/A',{}).get('/S')=='/GoTo':raise ValueError('含内部跳转的页面暂不支持多页拼版。')
        box=tuple(map(float,obj.get('/Rect',(0,0,0,0))))
        if box[2]<=clipped[0] or box[0]>=clipped[2] or box[3]<=clipped[1] or box[1]>=clipped[3]:continue
        if box[0]<clipped[0]-.01 or box[1]<clipped[1]-.01 or box[2]>clipped[2]+.01 or box[3]>clipped[3]+.01:
            raise ValueError('拼版将裁断注释，请改用适应模式或增加留白。')
        annots.append(ref)
    page[NameObject('/Annots')]=ArrayObject(annots)
    matrix=op
    start=len(dest.get('/Annots',[]));dest.merge_transformed_page(page,matrix,over=over,expand=False)
    # Reset geometry/AP to raw values, then transform exactly once, independent of
    # which annotation fields the pinned pypdf merge implementation transforms.
    new_refs=list(dest.get('/Annots',[]))[start:]
    for new,old in zip(new_refs,annots):
        obj,original=new.get_object(),old.get_object()
        for key in ('/Rect','/QuadPoints','/Vertices','/L','/CL','/InkList','/AP'):
            if key in original:obj[NameObject(key)]=original[key].clone(dest.indirect_reference.pdf,force_duplicate=True)
        obj.pop(NameObject('/Popup'),None) if '/Popup' not in original else None
    holder=PageObject();holder[NameObject('/Annots')]=ArrayObject(new_refs)
    _annotations(holder,matrix,dest.indirect_reference.pdf)


def positions(rule,count,size):
    n=rule.get('n',4);cols,rows={2:(2,1),4:(2,2),6:(2,3),9:(3,3)}[n]
    margin=rule.get('margin',24);gap=rule.get('gap',0);gutter=rule.get('gutter',0)
    left=margin+(gutter if rule.get('binding','left')=='left' else 0)
    cw=(size[0]-2*margin-gutter-(cols-1)*gap)/cols;ch=(size[1]-2*margin-(rows-1)*gap)/rows
    if min(cw,ch)<=0:raise ValueError('纸张容不下当前边距、装订边和间距。')
    slots=[]
    for i in range(count):
        row,col=divmod(i,cols) if rule.get('order','z')=='z' else (i%rows,i//rows)
        x=left+col*(cw+gap);top=size[1]-margin-row*(ch+gap)
        slots.append((x,top-ch,x+cw,top))
    return slots


def impose(reader,rule,templates,cancel):
    from pypdf import PdfWriter
    from pypdf.generic import NameObject,DictionaryObject
    writer=PdfWriter();op=rule['op'];n=len(reader.pages)
    if not n:raise ValueError('没有可拼版的页面。')
    if op=='concat':
        margin=rule.get('margin',24);gap=rule.get('gap',0);gutter=rule.get('gutter',0)
        width=rule.get('width',max(physical_size(p)[0] for p in reader.pages))
        heights=[physical_size(p)[1]*width/physical_size(p)[0] for p in reader.pages]
        size=(width+2*margin+gutter,sum(heights)+2*margin+gap*(n-1));page=blank(writer,size)
        y=size[1]-margin;rects=[];x=margin+(gutter if rule.get('binding','left')=='left' else 0)
        for source,height in zip(reader.pages,heights):
            stopped(cancel);rect=(x,y-height,x+width,y);rects.append(rect);place(page,source,rect,size);y-=height+gap
        if rule.get('border'):templates.borders(page,size,rects)
        return writer
    size=tuple(rule.get('size',A4[::-1] if op=='booklet' or rule.get('n',4)==2 else A4))
    if op=='nup':
        k=rule.get('n',4);slots=positions(rule,k,size)
        for start in range(0,n,k):
            stopped(cancel);page=blank(writer,size)
            for source,rect in zip(reader.pages[start:start+k],slots):place(page,source,rect,size,rule.get('fit_mode','fit'))
            if rule.get('border'):templates.borders(page,size,slots)
    else:
        margin=rule.get('margin',24);gap=rule.get('gap',0);gutter=rule.get('gutter',0)
        width=(size[0]-2*margin-gap)/2
        slots=[(margin,margin,margin+width-gutter,size[1]-margin),
               (margin+width+gap+gutter,margin,size[0]-margin,size[1]-margin)]
        if width<=gutter or size[1]<=2*margin:raise ValueError('小册子装订留白超出纸张。')
        padded=4*math.ceil(n/4)
        for sheet in range(padded//4):
            for back,pair in enumerate(((padded-2*sheet,1+2*sheet),(2+2*sheet,padded-1-2*sheet))):
                stopped(cancel);page=blank(writer,size)
                if rule.get('binding','left')=='right':pair=pair[::-1]
                for number,rect in zip(pair,slots):
                    if number<=n:place(page,reader.pages[number-1],rect,size)
                if rule.get('border'):templates.borders(page,size,slots)
                if back and rule.get('duplex','short')=='long':page.rotation=180
        writer._root_object[NameObject('/ViewerPreferences')]=DictionaryObject({NameObject('/Duplex'):NameObject('/DuplexFlipLongEdge' if rule.get('duplex','short')=='long' else '/DuplexFlipShortEdge')})
    return writer


def compose(plan,sources,progress=None,cancel=None):
    from .exporter import build_writer
    from .overlays import Templates,apply_overlay,verify_assets
    verify_assets(plan)
    # Keep stage-one geometry, but force explicit normalization of source Rotate.
    refs=tuple(replace(ref,target_rect=ref.target_rect or page_geometry(ref,sources[ref.source_id],plan.sheet)[1]) for ref in plan.pages)
    base=replace(plan,pages=refs,rules=(),overlays=())
    writer=build_writer(base,sources,progress,cancel);writer._header=b'%PDF-1.7';templates=Templates()
    try:
        for step in plan.rules:
            stopped(cancel)
            if step['op']=='overlay':apply_overlay(writer,step['overlay'],templates,cancel)
            else:
                settings=dict(step);settings.setdefault('gap',plan.sheet.gap)
                if plan.sheet.size is not None and step['op']!='concat':settings.setdefault('size',plan.sheet.size)
                reader=reader_of(writer);new=impose(reader,settings,templates,cancel);writer.close();writer=new
        for layer in plan.overlays:apply_overlay(writer,layer,templates,cancel)
        verify_assets(plan);return writer
    except BaseException:
        writer.close();raise
