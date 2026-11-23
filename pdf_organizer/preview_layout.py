"""Derived sheet-to-source mapping; OutputPlan remains the only saved model."""
from dataclasses import replace
from .model import OutputPlan, PageRef, Sheet, SourceInfo, page_geometry
from .composition import positions, A4
from .rule_spec import page_count


def sheet_view(plan,sources,index):
    """Return one equivalent sheet plan plus hit areas, or the full complex plan."""
    count=page_count(plan)
    index=max(0,min(index,max(0,count-1)))
    layouts=[r for r in plan.rules if r['op']!='overlay']
    simple=not plan.rules or (len(layouts)==1 and plan.rules[0]['op']=='nup')
    layers=list(plan.overlays)+[r['overlay'] for r in plan.rules if r['op']=='overlay']
    # Preserve exact template/font choices and selected-page semantics in complex cases.
    if any(layer.get('pages') is not None or (layer['kind']=='page_number' and
           (not layer.get('format','').isascii() or '{{' in layer.get('format',''))) for layer in layers):simple=False
    if not simple:return plan,index,None,(),count
    if layouts:
        r=dict(layouts[0]);r.setdefault('gap',plan.sheet.gap);n=r.get('n',4);size=tuple(r.get('size',plan.sheet.size or (A4[::-1] if n==2 else A4)))
        start=index*n;refs=plan.pages[start:start+n];rects=positions(r,len(refs),size)
        rules=[r]+[dict(step) for step in plan.rules[1:]]
    else:
        start=index;refs=plan.pages[index:index+1];rules=[]
        if not refs:return plan,index,None,(),count
        size=page_geometry(refs[0],sources[refs[0].source_id],plan.sheet)[0]
        rects=[(0,0,*size)]
    def number(layer):
        layer=dict(layer)
        if layer['kind']=='page_number':
            layer['format']=layer.get('format','{page} / {total}').replace('{page}',str(index+layer.get('start',1))).replace('{total}',str(count))
        return layer
    rules=tuple(dict(step,overlay=number(step['overlay'])) if step['op']=='overlay' else step for step in rules)
    overlays=tuple(number(layer) for layer in plan.overlays)
    local=replace(plan,pages=tuple(refs),rules=rules,overlays=overlays)
    return local,0,size,tuple((start+i,tuple(rect)) for i,rect in enumerate(rects)),count


def unpack(data):
    p=data['plan']
    return OutputPlan(tuple(PageRef(**r) for r in p['pages']),Sheet(**p['sheet']),tuple(p.get('overlays',())),tuple(p.get('rules',()))),{k:SourceInfo(**v) for k,v in data['sources'].items()}


def swap_pages(document,a,b):
    if a==b:return False
    if not (0<=a<len(document.plan.pages) and 0<=b<len(document.plan.pages)):return False
    refs=list(document.plan.pages);refs[a],refs[b]=refs[b],refs[a]
    return document.commit(replace(document.plan,pages=tuple(refs)),(b,))
