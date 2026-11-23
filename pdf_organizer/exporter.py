"""Export an immutable OutputPlan. No PDF/Office engine is imported here."""
from collections import Counter
from io import BytesIO
from pathlib import Path
import os,tempfile
from .model import page_geometry,validate_plan,verify_sources


def layout(ref,source,sheet):
    from pypdf import Transformation
    l,b,r,t,native,unit=source.boxes[ref.page_index]
    if ref.crop is not None:l,b,r,t=ref.crop
    rot=(native+ref.rotate)%360
    transforms={0:(1,0,0,1,-l,-b),90:(0,-1,1,0,-b,r),180:(-1,0,0,-1,r,t),270:(0,1,-1,0,t,-l)}
    size,target,(w,h)=page_geometry(ref,source,sheet);x,y,x2,y2=target
    sx,sy=(x2-x)/w,(y2-y)/h
    if ref.fit_mode!='stretch':sx=sy=(min if ref.fit_mode=='fit' else max)(sx,sy)
    op=Transformation(transforms[rot]).scale(unit*sx,unit*sy).translate(x+(x2-x-w*sx)/2,y+(y2-y-h*sy)/2)
    return size,target,(l,b,r,t),op


def _annotations(page,op,writer):
    from pypdf.generic import NameObject,ArrayObject,FloatObject,RectangleObject
    from pypdf.generic._appearance_stream import transform_annotation_appearance
    seen=set()
    def pairs(values):return ArrayObject(FloatObject(v) for i in range(0,len(values),2) for v in op.apply_on((values[i],values[i+1])))
    def transform(ref):
        obj=ref.get_object()
        key=getattr(obj,'indirect_reference',None)
        if key and key.idnum in seen:return
        if key:seen.add(key.idnum)
        if '/Rect' in obj:
            l,b,r,t=map(float,obj['/Rect']);pts=[op.apply_on(p) for p in ((l,b),(l,t),(r,b),(r,t))]
            obj[NameObject('/Rect')]=RectangleObject((min(p[0] for p in pts),min(p[1] for p in pts),max(p[0] for p in pts),max(p[1] for p in pts)))
        for field in ('/QuadPoints','/Vertices','/L','/CL'):
            if field in obj:obj[NameObject(field)]=pairs(obj[field])
        if '/InkList' in obj:obj[NameObject('/InkList')]=ArrayObject(pairs(line) for line in obj['/InkList'])
        if '/AP' in obj:
            obj[NameObject('/AP')]=obj['/AP'].clone(writer,force_duplicate=True)
            transform_annotation_appearance(obj,op)
        if '/Popup' in obj:transform(obj['/Popup'])
    for ref in page.get('/Annots',[]):transform(ref)


def build_writer(plan,sources,progress=None,cancel=None):
    import pypdf
    if pypdf.__version__!='6.21.4':raise RuntimeError('PDF整理组件需要pypdf 6.21.4，请使用配套开发包。')
    if plan.rules or plan.overlays:
        from .composition import compose
        return compose(plan,sources,progress,cancel)
    from pypdf import PdfReader,PdfWriter
    from pypdf.generic import NameObject,NumberObject,RectangleObject,DecodedStreamObject
    validate_plan(plan,sources)
    if not plan.pages:raise ValueError('请先添加页面。')
    verify_sources(plan,sources)
    writer=PdfWriter();readers={};data={};counts=Counter()
    for index,ref in enumerate(plan.pages):
        if cancel and cancel.is_set():raise InterruptedError('已停止导出。')
        source=sources[ref.source_id];key=(ref.source_id,ref.page_index);counts[key]+=1
        if ref.source_id not in readers:
            # Reader works on a stable in-memory snapshot, never a writable source.
            data[ref.source_id]=Path(source.path).read_bytes()
            reader=PdfReader(BytesIO(data[ref.source_id]));reader.add_form_topname('s'+ref.source_id)
            readers[ref.source_id]=reader
        reader=readers[ref.source_id]
        if counts[key]>1:
            # Copies have independent widgets/field names and annotation objects.
            reader=PdfReader(BytesIO(data[ref.source_id]));reader.add_form_topname(f's{ref.source_id}_copy{index}')
        writer.append(reader,pages=[ref.page_index],import_outline=False)
        page=writer.pages[-1]
        native=source.boxes[ref.page_index][4]
        simple=plan.sheet.size is None and plan.sheet.margin==0 and ref.crop is None and ref.target_rect is None
        if simple:
            page.rotation=int((native+ref.rotate)%360)
        else:
            # Keep the appended AcroForm tree; transform content and annotations in place.
            size,target,clip,op=layout(ref,source,plan.sheet)
            stream=DecodedStreamObject();contents=page.get_contents()
            l,b,r,t=clip
            body=(f'q {l:.8f} {b:.8f} {r-l:.8f} {t-b:.8f} re W n\n'.encode()+
                  (contents.get_data() if contents is not None else b'')+b'\nQ')
            stream.set_data(body);page.replace_contents(stream)
            page.add_transformation(op)
            contents=page.get_contents();x,y,x2,y2=target
            stream=DecodedStreamObject();stream.set_data(f'q {x:.8f} {y:.8f} {x2-x:.8f} {y2-y:.8f} re W n\n'.encode()+contents.get_data()+b'\nQ')
            page.replace_contents(stream);_annotations(page,op,writer)
            page.rotation=0;page[NameObject('/UserUnit')]=NumberObject(1)
            for field in ('/MediaBox','/CropBox','/TrimBox','/BleedBox','/ArtBox'):
                page[NameObject(field)]=RectangleObject((0,0,*size))
        if progress:progress(index+1,len(plan.pages))
    verify_sources(plan,sources)
    return writer


def export_pdf(plan,sources,path,progress=None,cancel=None):
    from .rule_spec import assets
    from .overlays import verify_assets
    path=Path(path).absolute()
    if path.suffix.lower()!='.pdf':raise ValueError('输出文件须为PDF。')
    if any(path.resolve()==Path(s.path).resolve() for s in sources.values()):raise ValueError('不能覆盖原始PDF，请选择新的文件名。')
    if any(path.resolve()==Path(a['path']).resolve() for a in assets(plan)):raise ValueError('不能覆盖背景或印章原件。')
    verify_assets(plan)
    writer=build_writer(plan,sources,progress,cancel)
    path.parent.mkdir(parents=True,exist_ok=True)
    fd,name=tempfile.mkstemp(prefix='.ruyi-pdf-',suffix='.tmp',dir=path.parent);os.close(fd)
    try:
        writer.write(name)
        if cancel and cancel.is_set():raise InterruptedError('已停止导出。')
        verify_sources(plan,sources);verify_assets(plan);os.replace(name,path)
    finally:
        writer.close()
        if os.path.exists(name):os.unlink(name)
    return path


def split_pdf(plan,sources,directory,count,progress=None,cancel=None):
    from dataclasses import replace
    import shutil
    if type(count) is not int or count<1:raise ValueError('每份页数须大于零。')
    if plan.rules or plan.overlays:
        from .presets import split_composed
        return split_composed(plan,sources,directory,count,progress,cancel)
    directory=Path(directory)
    if directory.exists():raise ValueError('拆分目录已存在，请选择新目录。')
    directory.parent.mkdir(parents=True,exist_ok=True)
    workspace=Path(tempfile.mkdtemp(prefix='.ruyi-split-',dir=directory.parent));outputs=[]
    try:
        for start in range(0,len(plan.pages),count):
            part=replace(plan,pages=plan.pages[start:start+count]);name=f'页面_{start+1:04}-{start+len(part.pages):04}.pdf'
            export_pdf(part,sources,workspace/name,cancel=cancel);outputs.append(name)
            if progress:progress(start+len(part.pages),len(plan.pages))
        if not outputs:raise ValueError('没有可拆分的页面。')
        workspace.rename(directory)
    finally:
        if workspace.exists():shutil.rmtree(workspace)
    return [directory/name for name in outputs]
