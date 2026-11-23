"""Cached single-page templates for borders, text, digit glyphs and alpha images."""
from io import BytesIO
from pathlib import Path
from string import Formatter
import hashlib,math,os,threading
from .model import sha256
from .rule_spec import assets

FONT_LOCK=threading.RLock()


def verify_assets(plan):
    for item in assets(plan):
        path=Path(item['path'])
        if not path.is_file() or sha256(path).lower()!=item['sha256'].lower():raise ValueError(f'印章/背景丢失或已变化：{path.name}')


class Templates:
    def __init__(self):self.cache={};self.readers=[];self.created=0

    def make(self,key,size,draw):
        if key not in self.cache:
            from reportlab.pdfgen import canvas
            from pypdf import PdfReader
            with FONT_LOCK:
                stream=BytesIO();c=canvas.Canvas(stream,pagesize=size,invariant=1,pageCompression=1);draw(c);c.showPage();c.save()
            stream.seek(0);reader=PdfReader(stream);self.readers.append(reader);self.cache[key]=reader.pages[0];self.created+=1
        return self.cache[key]

    def font(self,text,spec):
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        if all(ord(c)<128 for c in text) and not spec.get('font_path'):return 'Helvetica'
        path=spec.get('font_path')
        if not path:
            folder=Path(os.environ.get('SystemRoot','C:/Windows'))/'Fonts'
            path=next((str(folder/name) for name in ('msyh.ttc','simhei.ttf','simsun.ttc') if (folder/name).is_file()),None)
        if not path:raise ValueError('中文水印需要本地中文TrueType字体。')
        name='Ruyi_'+hashlib.sha256(str(path).encode()).hexdigest()[:12]
        with FONT_LOCK:
            if name not in pdfmetrics.getRegisteredFontNames():pdfmetrics.registerFont(TTFont(name,path,subfontIndex=0))
        return name

    def text(self,text,spec):
        from reportlab.pdfbase import pdfmetrics
        from reportlab.lib.colors import HexColor
        font=self.font(text,spec);size=spec.get('font_size',12);width=pdfmetrics.stringWidth(text,font,size)
        def draw(c):
            c.setFillColor(HexColor(spec.get('color','#243047')));c.setFillAlpha(spec.get('opacity',1));c.setFont(font,size);c.drawString(1,size*.35,text)
        key=('text',text,font,size,spec.get('color'),spec.get('opacity',1))
        return self.make(key,(max(width+2,2),size*1.6),draw),width

    def borders(self,page,size,rects):
        from pypdf import Transformation
        rects=tuple(tuple(r) for r in rects)
        def draw(c):
            c.setStrokeColorRGB(.35,.4,.48);c.setLineWidth(.5)
            for x,y,r,t in rects:c.rect(x+.25,y+.25,r-x-.5,t-y-.5,stroke=1,fill=0)
        template=self.make(('border',size,rects),size,draw);unit=float(page.get('/UserUnit',1))
        page.merge_transformed_page(template,Transformation().scale(1/unit))

    def stamp(self,spec):
        from PIL import Image
        from reportlab.lib.utils import ImageReader
        with Image.open(spec['path']) as im:size=im.size
        def draw(c):
            c.setFillAlpha(spec.get('opacity',1));c.drawImage(ImageReader(spec['path']),0,0,*size,mask='auto')
        return self.make(('stamp',spec['sha256'],spec.get('opacity',1)),size,draw)


def apply_overlay(writer,spec,templates,cancel=None):
    from pypdf import Transformation,PdfReader
    from .composition import physical_size,place,stopped
    from .rule_spec import overlay
    spec=overlay(dict(spec));kind=spec['kind'];total=len(writer.pages)
    selected=spec.get('pages')
    if selected is not None and any(i>=total for i in selected):raise ValueError('叠加层页号超出当前结果页数。')
    for index,page in enumerate(writer.pages):
        stopped(cancel)
        if selected is not None and index not in selected:continue
        # Booklet long-edge back /Rotate=180 is deliberately retained: overlays
        # are positioned in the same raw sheet coordinate system before printing.
        unit=float(page.get('/UserUnit',1));w,h=float(page.mediabox.width)*unit,float(page.mediabox.height)*unit
        if kind=='page_number':
            pieces=[]
            for literal,field,fmt,conversion in Formatter().parse(spec.get('format','{page} / {total}')):
                if literal:pieces.append(literal)
                if field:pieces.extend(str(index+spec.get('start',1) if field=='page' else total))
            fragments=[templates.text(text,spec) for text in pieces];width=sum(v[1] for v in fragments)
            inset=spec.get('inset',12);align=spec.get('align','right')
            x={'left':inset,'center':(w-width)/2,'right':w-inset-width}[align]
            height=spec.get('font_size',12)*1.6
            if x<0 or inset+height>h or x+width>w:raise ValueError('页码无法放入纸张，请减小字号或留白。')
            for template,advance in fragments:
                page.merge_transformed_page(template,Transformation((1/unit,0,0,1/unit,(x-1)/unit,inset/unit)));x+=advance
        elif kind=='text':
            template,width=templates.text(spec['text'],spec);tw,th=map(float,(template.mediabox.width,template.mediabox.height))
            target=spec.get('rect',(0,0,w,h));x,y,r,t=target
            if r>w or t>h:raise ValueError('水印区域超出纸张。')
            angle=spec.get('angle',30);rad=math.radians(angle)
            bw,bh=abs(tw*math.cos(rad))+abs(th*math.sin(rad)),abs(tw*math.sin(rad))+abs(th*math.cos(rad))
            scale=min(1,(r-x)/bw,(t-y)/bh)
            op=Transformation().translate(-tw/2,-th/2).rotate(angle).scale(scale).translate((x+r)/2,(y+t)/2)
            page.merge_transformed_page(template,Transformation(tuple(v/unit for v in op.ctm)))
        elif kind=='stamp':
            target=spec['rect']
            if target[2]>w or target[3]>h:raise ValueError('印章区域超出纸张。')
            place(page,templates.stamp(spec),target,(w,h),spec.get('fit_mode','fit'),interactive=False)
        else:
            key=('background',spec['sha256'],spec.get('page_index',0),spec.get('opacity',1))
            if key not in templates.cache:
                reader=PdfReader(BytesIO(Path(spec['path']).read_bytes()))
                if reader.is_encrypted:raise ValueError('背景PDF需先解密。')
                if spec.get('page_index',0)>=len(reader.pages):raise ValueError('背景PDF没有指定页面。')
                from pypdf import PdfWriter
                from pypdf.generic import NameObject,DictionaryObject,FloatObject,DecodedStreamObject
                owner=PdfWriter();owner.append(reader,pages=[spec.get('page_index',0)],import_outline=False)
                template=owner.pages[0]
                resources=template['/Resources'];states=DictionaryObject(resources.get('/ExtGState',{}));name='/RuyiBackgroundOpacity'
                while name in states:name+='X'
                states[NameObject(name)]=DictionaryObject({NameObject('/ca'):FloatObject(spec.get('opacity',1)),NameObject('/CA'):FloatObject(spec.get('opacity',1))})
                resources[NameObject('/ExtGState')]=states
                contents=template.get_contents();stream=DecodedStreamObject()
                stream.set_data(('q '+name+' gs\n').encode()+(contents.get_data() if contents is not None else b'')+b'\nQ')
                template.replace_contents(stream)
                templates.readers.extend((reader,owner));templates.cache[key]=template
            place(page,templates.cache[key],(0,0,w,h),(w,h),spec.get('fit_mode','fit'),over=False,interactive=False)
