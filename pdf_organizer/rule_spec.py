"""JSON specifications inside OutputPlan; no second document/placement model."""
import math,re


class FrozenDict(dict):
    def _readonly(self,*args,**kwargs):raise TypeError('方案规则不可原位修改，请提交新方案。')
    __setitem__=__delitem__=clear=pop=popitem=setdefault=update=__ior__=_readonly
    def __deepcopy__(self,memo):return self


def freeze(value):
    if isinstance(value,dict):return FrozenDict((k,freeze(v)) for k,v in value.items())
    if isinstance(value,(tuple,list)):return tuple(freeze(v) for v in value)
    return value


def numeric(value,minimum=0,maximum=1e9):
    if isinstance(value,bool) or not isinstance(value,(int,float)) or not math.isfinite(value) or not minimum<=value<=maximum:
        raise ValueError('规则中的数值超出有效范围。')
    return float(value)


def overlay(spec):
    if not isinstance(spec,dict):raise ValueError('叠加层须为对象。')
    s=dict(spec);kind=s.get('kind')
    allowed={'kind','pages','opacity'}|{'page_number':{'color','font_size','font_path','align','inset','format','start'},
        'text':{'color','font_size','font_path','angle','rect','text'},'stamp':{'path','sha256','rect','fit_mode'},
        'background':{'path','sha256','page_index','fit_mode'}}.get(kind,set())
    if set(s)-allowed or kind not in ('page_number','text','stamp','background'):raise ValueError('不支持的叠加层。')
    if s.get('pages') is not None and (not isinstance(s['pages'],(tuple,list)) or any(type(i) is not int or i<0 for i in s['pages'])):raise ValueError('叠加层页号须为零基整数。')
    numeric(s.get('opacity',1),0,1);numeric(s.get('font_size',12),4,300);numeric(s.get('inset',12),0,14400)
    numeric(s.get('angle',0),-360,360)
    if not re.fullmatch(r'#[0-9a-fA-F]{6}',s.get('color','#243047')):raise ValueError('颜色须为#RRGGBB。')
    if s.get('align','right') not in ('left','center','right'):raise ValueError('文字对齐无效。')
    if 'rect' in s:
        r=s['rect']
        if len(r)!=4 or any(numeric(v)<0 for v in r) or r[2]<=r[0] or r[3]<=r[1]:raise ValueError('叠加区域无效。')
    if kind=='page_number':
        from string import Formatter
        if type(s.get('start',1)) is not int or s.get('start',1)<0:raise ValueError('起始页码无效。')
        for literal,field,fmt,conversion in Formatter().parse(s.get('format','{page} / {total}')):
            if field is not None and (field not in ('page','total') or fmt or conversion):raise ValueError('页码格式仅支持{page}和{total}。')
    if kind=='text' and (not isinstance(s.get('text'),str) or not 0<len(s['text'])<=500):raise ValueError('水印文字须为1至500字。')
    if kind in ('stamp','background'):
        if not isinstance(s.get('path'),str) or not s['path'] or not re.fullmatch(r'[0-9a-fA-F]{64}',s.get('sha256','')):raise ValueError('图片/PDF叠加层缺少路径或SHA256。')
        if kind=='stamp' and 'rect' not in s:raise ValueError('印章须指定区域。')
        if kind=='background' and (type(s.get('page_index',0)) is not int or s.get('page_index',0)<0):raise ValueError('背景源页无效。')
    if s.get('fit_mode','fit') not in ('fit','fill','stretch'):raise ValueError('适配模式无效。')
    return freeze(s)


def rule(spec):
    if not isinstance(spec,dict):raise ValueError('规则须为对象。')
    s=dict(spec);op=s.get('op')
    allowed={'op','margin','gap','border','gutter','binding'}|{'nup':{'n','order','size','fit_mode'},'booklet':{'size','duplex'},'concat':{'width'},'overlay':{'overlay'}}.get(op,set())
    if set(s)-allowed or op not in ('nup','booklet','concat','overlay'):raise ValueError('不支持的拼版规则。')
    if op=='overlay':
        if set(s)-{'op','overlay'}:raise ValueError('叠加规则含多余字段。')
        s['overlay']=overlay(s.get('overlay'));return freeze(s)
    for key in ('margin','gap','gutter'):numeric(s.get(key,0),0,14400)
    if s.get('binding','left') not in ('left','right'):raise ValueError('装订边须为left/right。')
    if s.get('fit_mode','fit') not in ('fit','fill','stretch'):raise ValueError('适配模式无效。')
    if 'size' in s and (len(s['size'])!=2 or any(numeric(v,1,14400)<=0 for v in s['size'])):raise ValueError('纸张尺寸无效。')
    if type(s.get('border',False)) is not bool:raise ValueError('边框须为布尔值。')
    if op=='nup' and (type(s.get('n',4)) is not int or s.get('n',4) not in (2,4,6,9)):raise ValueError('仅支持2/4/6/9宫格。')
    if s.get('order','z') not in ('z','column'):raise ValueError('宫格顺序无效。')
    if op=='booklet' and s.get('duplex','short') not in ('short','long'):raise ValueError('双面翻转边须为short/long。')
    if op=='concat' and 'width' in s:numeric(s['width'],1,14400)
    return freeze(s)


def assets(plan):
    layers=list(plan.overlays)+[r['overlay'] for r in plan.rules if r['op']=='overlay']
    return [layer for layer in layers if layer['kind'] in ('stamp','background')]


def page_count(plan):
    n=len(plan.pages)
    for r in plan.rules:
        if r['op']=='nup':n=math.ceil(n/r.get('n',4))
        elif r['op']=='booklet':n=2*math.ceil(n/4)
        elif r['op']=='concat':n=int(n>0)
    return n
