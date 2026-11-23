"""Canonical plan, source registry and immutable undo history. Coordinates: pt."""
from __future__ import annotations
from dataclasses import asdict, dataclass, replace
from pathlib import Path
import hashlib,json,math,os,tempfile,uuid


def number(value):
    if isinstance(value,bool) or not isinstance(value,(int,float)) or not math.isfinite(value):
        raise ValueError('坐标和尺寸必须为有限数字。')
    return float(value)


def rect(value):
    if value is None:return None
    if not isinstance(value,(list,tuple)) or len(value)!=4:raise ValueError('区域必须包含四个坐标。')
    result=tuple(number(v) for v in value)
    if result[2]<=result[0] or result[3]<=result[1]:raise ValueError('区域宽高必须大于零。')
    return result


@dataclass(frozen=True)
class PageRef:
    source_id: str
    page_index: int
    rotate: int = 0
    crop: tuple | None = None
    fit_mode: str = 'fit'
    target_rect: tuple | None = None

    def __post_init__(self):
        if not isinstance(self.source_id,str) or not self.source_id:raise ValueError('缺少源文件标识。')
        if type(self.page_index) is not int or self.page_index<0:raise ValueError('页面索引无效。')
        if type(self.rotate) is not int or self.rotate%90:raise ValueError('旋转必须为90度的整数倍。')
        object.__setattr__(self,'rotate',self.rotate%360)
        object.__setattr__(self,'crop',rect(self.crop));object.__setattr__(self,'target_rect',rect(self.target_rect))
        if self.fit_mode not in ('fit','fill','stretch'):raise ValueError('页面适配模式无效。')


@dataclass(frozen=True)
class Sheet:
    size: tuple | None = None
    gap: float = 0
    margin: float = 0

    def __post_init__(self):
        if self.size is not None:
            if not isinstance(self.size,(tuple,list)) or len(self.size)!=2:raise ValueError('纸张尺寸无效。')
            size=tuple(number(v) for v in self.size)
            if min(size)<=0 or max(size)>14400:raise ValueError('纸张尺寸超出范围。')
            object.__setattr__(self,'size',size)
        object.__setattr__(self,'margin',number(self.margin));object.__setattr__(self,'gap',number(self.gap))
        if self.margin<0:raise ValueError('页边距不能小于零。')
        if self.gap<0:raise ValueError('页面间距不能小于零。')


@dataclass(frozen=True)
class OutputPlan:
    pages: tuple[PageRef,...] = ()
    sheet: Sheet = Sheet()
    overlays: tuple = ()
    rules: tuple = ()

    def __post_init__(self):
        from .rule_spec import rule,overlay
        object.__setattr__(self,'pages',tuple(self.pages))
        object.__setattr__(self,'overlays',tuple(overlay(s) for s in self.overlays))
        object.__setattr__(self,'rules',tuple(rule(s) for s in self.rules))
        if len(self.rules)>16:raise ValueError('规则链最多16步。')
        if len(self.pages)>10000:raise ValueError('方案最多支持10000页。')


@dataclass(frozen=True)
class SourceInfo:
    source_id: str
    path: str
    sha256: str
    byte_size: int
    mtime_ns: int
    # (CropBox left,bottom,right,top, native Rotate, UserUnit) per source page.
    boxes: tuple

    @property
    def name(self):return Path(self.path).name


def sha256(path):
    h=hashlib.sha256()
    with Path(path).open('rb') as f:
        for block in iter(lambda:f.read(1024*1024),b''):h.update(block)
    return h.hexdigest()


def inspect_sources(paths):
    from pypdf import PdfReader
    result=[]
    for name in paths:
        path=Path(name).resolve(strict=True)
        if path.suffix.lower()!='.pdf':raise ValueError(f'仅支持PDF：{path.name}')
        before=path.stat();digest=sha256(path)
        reader=PdfReader(path)
        if reader.is_encrypted:raise ValueError(f'请先解密PDF：{path.name}')
        boxes=tuple(tuple(map(float,p.cropbox))+(p.rotation%360,float(p.get('/UserUnit',1))) for p in reader.pages)
        if not boxes:raise ValueError(f'PDF没有页面：{path.name}')
        after=path.stat()
        if (before.st_size,before.st_mtime_ns)!=(after.st_size,after.st_mtime_ns):raise ValueError('导入期间源文件发生变化。')
        result.append(SourceInfo(uuid.uuid4().hex,str(path),digest,after.st_size,after.st_mtime_ns,boxes))
    return result


def page_geometry(ref,source,sheet):
    l,b,r,t,native,unit=source.boxes[ref.page_index]
    if ref.crop is not None:l,b,r,t=ref.crop
    w,h=(r-l)*unit,(t-b)*unit
    if (native+ref.rotate)%180:w,h=h,w
    size=sheet.size or (w,h)
    target=ref.target_rect or (sheet.margin,sheet.margin,size[0]-sheet.margin,size[1]-sheet.margin)
    rect(target)
    if target[0]<0 or target[1]<0 or target[2]>size[0]+.001 or target[3]>size[1]+.001:raise ValueError('目标区域超出纸张。')
    return size,target,(w,h)


def validate_plan(plan,sources):
    for ref in plan.pages:
        source=sources.get(ref.source_id)
        if source is None or ref.page_index>=len(source.boxes):raise ValueError('方案引用了不存在的源页。')
        box=source.boxes[ref.page_index]
        if ref.crop is not None and (ref.crop[0]<box[0]-.001 or ref.crop[1]<box[1]-.001 or ref.crop[2]>box[2]+.001 or ref.crop[3]>box[3]+.001):
            raise ValueError('裁切区域超出源页可见区域。')
        page_geometry(ref,source,plan.sheet)


class PlanDocument:
    def __init__(self,plan=None,sources=None):
        self.plan=plan or OutputPlan();self.sources=dict(sources or {});self.selection=()
        self.undo_stack=[];self.redo_stack=[];self.revision=0;self.history_limit=50

    def _state(self):return self.plan,self.sources.copy(),self.selection

    def commit(self,plan,selection=(),sources=None):
        registry=self.sources if sources is None else dict(sources)
        validate_plan(plan,registry)
        selected=tuple(sorted(set(selection)))
        if any(i<0 or i>=len(plan.pages) for i in selected):raise ValueError('选择的页面无效。')
        if plan==self.plan and registry==self.sources:return False
        self.undo_stack.append(self._state());self.undo_stack=self.undo_stack[-self.history_limit:];self.redo_stack.clear()
        self.plan,self.sources,self.selection=plan,registry.copy(),selected;self.revision+=1
        return True

    def import_sources(self,items):
        registry=self.sources.copy();pages=list(self.plan.pages);start=len(pages)
        for source in items:
            registry[source.source_id]=source
            pages.extend(PageRef(source.source_id,i,fit_mode=self.plan.pages[0].fit_mode if self.plan.pages else 'fit') for i in range(len(source.boxes)))
        self.commit(replace(self.plan,pages=tuple(pages)),(start,) if len(pages)>start else (),registry)

    def select(self,indices):self.selection=tuple(sorted(set(i for i in indices if 0<=i<len(self.plan.pages))))

    def delete(self):
        chosen=set(self.selection)
        self.commit(replace(self.plan,pages=tuple(p for i,p in enumerate(self.plan.pages) if i not in chosen)))

    def duplicate(self):
        chosen=set(self.selection);pages=[];selected=[]
        for i,p in enumerate(self.plan.pages):
            pages.append(p)
            if i in chosen:selected.append(len(pages));pages.append(replace(p))
        self.commit(replace(self.plan,pages=tuple(pages)),selected)

    def rotate(self):
        chosen=set(self.selection)
        self.commit(replace(self.plan,pages=tuple(replace(p,rotate=(p.rotate+90)%360) if i in chosen else p for i,p in enumerate(self.plan.pages))),self.selection)

    def move(self,indices,boundary):
        indices=tuple(sorted(set(indices)));chosen=set(indices);n=len(self.plan.pages)
        if any(i<0 or i>=n for i in indices) or not 0<=boundary<=n:raise ValueError('拖放位置无效。')
        rest=[p for i,p in enumerate(self.plan.pages) if i not in chosen]
        insert=boundary-sum(i<boundary for i in indices)
        group=[self.plan.pages[i] for i in indices];rest[insert:insert]=group
        self.commit(replace(self.plan,pages=tuple(rest)),range(insert,insert+len(group)))

    def configure(self,sheet,mode):
        self.commit(replace(self.plan,pages=tuple(replace(p,fit_mode=mode,target_rect=None) for p in self.plan.pages),sheet=sheet),self.selection)

    def undo(self):return self._restore(self.undo_stack,self.redo_stack)
    def redo(self):return self._restore(self.redo_stack,self.undo_stack)
    def _restore(self,source,target):
        if not source:return False
        target.append(self._state());self.plan,self.sources,self.selection=source.pop();self.revision+=1;return True

    def to_dict(self):
        p=asdict(self.plan)
        if not self.plan.rules:p.pop('rules')
        return {'format':'ruyi.pdf.plan','schema_version':2 if self.plan.rules or self.plan.overlays else 1,'sources':[asdict(s) for s in self.sources.values()],'plan':p}

    def save(self,path):
        path=Path(path)
        if not path.name.endswith('.ruyiplan.json'):raise ValueError('方案扩展名应为.ruyiplan.json。')
        data=self.to_dict()
        for item in data['sources']:
            try:item['path']=os.path.relpath(item['path'],path.parent)
            except ValueError:pass
        data['plan']=asset_paths(data['plan'],path.parent,False)
        atomic_write(path,json.dumps(data,ensure_ascii=False,indent=2).encode('utf-8'))

    @classmethod
    def load(cls,path):
        path=Path(path).resolve(strict=True)
        if path.stat().st_size>20*1024*1024:raise ValueError('方案文件过大。')
        data=json.loads(path.read_text(encoding='utf-8-sig'))
        if data.get('format')!='ruyi.pdf.plan' or data.get('schema_version') not in (1,2):raise ValueError('不支持的方案格式。')
        registry={}
        for item in data['sources']:
            item=dict(item);item['path']=str((path.parent/item['path']).resolve());item['boxes']=tuple(tuple(x) for x in item['boxes'])
            source=SourceInfo(**item)
            if source.source_id in registry:raise ValueError('源文件标识重复。')
            registry[source.source_id]=source
        p=asset_paths(data['plan'],path.parent,True);plan=OutputPlan(tuple(PageRef(**x) for x in p['pages']),Sheet(**p['sheet']),tuple(p['overlays']),tuple(p.get('rules',())))
        validate_plan(plan,registry);verify_sources(plan,registry)
        return cls(plan,registry)


def verify_sources(plan,sources):
    for sid in dict.fromkeys(p.source_id for p in plan.pages):
        source=sources[sid];path=Path(source.path)
        if not path.is_file() or sha256(path)!=source.sha256:raise ValueError(f'源文件丢失或已变化：{source.name}')


def asset_paths(value,parent,absolute):
    if isinstance(value,dict):
        result={k:asset_paths(v,parent,absolute) for k,v in value.items()}
        for key in ('path','font_path'):
            if key in result:
                if absolute:result[key]=str((parent/result[key]).resolve())
                else:
                    try:result[key]=os.path.relpath(result[key],parent)
                    except ValueError:pass
        return result
    if isinstance(value,(list,tuple)):return [asset_paths(v,parent,absolute) for v in value]
    return value


def atomic_write(path,data):
    path=Path(path);path.parent.mkdir(parents=True,exist_ok=True)
    fd,name=tempfile.mkstemp(prefix='.ruyi-',suffix='.tmp',dir=path.parent)
    try:
        with os.fdopen(fd,'wb') as f:f.write(data)
        os.replace(name,path)
    finally:
        if os.path.exists(name):os.unlink(name)
