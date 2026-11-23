"""Retained native background and small cached card tiles, never full-screen hover renders."""
from collections import OrderedDict
from itertools import groupby
from tkinter.font import Font
from PIL import Image, ImageTk, ImageColor


def rounded(x,y,r,b,k):
    return (x+k,y,r-k,y,r,y,r,y+k,r,b-k,r,b,r-k,b,x+k,b,x,b,x,b-k,x,y+k,x,y)


def edge_curve(index):
    controls=((-28,278-index*23),(-9,114-index*23*.4),(102-index,6-index*3),(318-index*26,-28))
    points=[]
    for step in range(33):
        u=step/32
        weights=((1-u)**3,3*(1-u)**2*u,3*(1-u)*u*u,u**3)
        points.extend(sum(weight*value for weight,value in zip(weights,axis)) for axis in zip(*controls))
    return tuple(points)


EDGE_CURVES=tuple(edge_curve(index) for index in range(9))
ICON_SIZES={'files':(25,29),'arrow-right':(25,25),'arrow-left':(25,25),
            'image-down':(32,36),'contact-round':(32,36),'settings-2':(20,20)}


class HomeSurface:
    def __init__(self, home):
        self.home=home
        self.items={};self.signature=None;self.style=None
        self.card_layouts={};self.fills={}
        self.brand_fonts={}
        self.photos=OrderedDict()
        self.bands=();self.band_count=0

    def photo(self,name,size,color=None):
        key=(name,size,color)
        photo=self.photos.get(key)
        if photo is None:
            if color is None:
                with self.home.logo.resize(size,Image.Resampling.LANCZOS) as image:
                    photo=ImageTk.PhotoImage(image,master=self.home)
            else:
                with self.home.icons[name].getchannel('A') as mask:
                    with mask.resize(size,Image.Resampling.LANCZOS) as alpha, Image.new('RGBA',size,color) as image:
                        image.putalpha(alpha)
                        photo=ImageTk.PhotoImage(image,master=self.home)
            self.photos[key]=photo
            if len(self.photos)>96:
                self.photos.popitem(last=False)
        else:
            self.photos.move_to_end(key)
        return photo

    def prepare(self,width,height,accent):
        scale=min(width/1024,height/760)
        logo_size=max(1,round(74*scale))
        images={'logo':self.photo('logo',(logo_size,logo_size))}
        for name,dimensions in ICON_SIZES.items():
            size=tuple(max(1,round(value*scale)) for value in dimensions)
            images[name]=self.photo(name,size,accent)
        brand_size=max(1,round(34*scale))
        if brand_size not in self.brand_fonts:
            font=Font(root=self.home,family='Microsoft YaHei',size=-brand_size,weight='bold')
            self.brand_fonts[brand_size]=(font,font.measure('如意'))
        return images,self.brand_fonts[brand_size]

    def item(self,key,kind,coords,tags='surface',**options):
        h=self.home
        old=self.items.get(key)
        if old is None:
            item=getattr(h,'create_'+kind)(*coords,tags=tags,**options)
        else:
            item,old_coords,old_options=old
            if coords!=old_coords:h.coords(item,*coords)
            if options!=old_options:h.itemconfigure(item,**options)
        self.items[key]=(item,coords,options)
        return item

    def set_style(self,t,gradient):
        self.home.configure(background=t['bg'])
        self.bands=()
        if gradient:
            start,end=ImageColor.getrgb(t['bg']),ImageColor.getrgb(t['hover'])
            colors=('#%02x%02x%02x'%tuple(round(a+(b-a)*i/127*.65) for a,b in zip(start,end)) for i in range(128))
            bands=[];top=0
            for color,group in groupby(colors):
                bottom=top+sum(1 for _ in group)
                bands.append((top/128,bottom/128,color));top=bottom
            self.bands=tuple(bands)
        for index in range(len(self.bands),self.band_count):
            self.item(('band',index),'rectangle',(0,0,0,0),state='hidden')
        self.fills={}
        active=t['hover'];active_value=sum(ImageColor.getrgb(active))
        for key in (*dict(self.home.MODES),'settings'):
            normal=t['bg'] if key=='settings' else t['soft']
            rgb=ImageColor.getrgb(normal)
            darker=active if active_value<sum(rgb) else '#%02x%02x%02x'%tuple(round(channel*.72) for channel in rgb)
            self.fills[key]=(normal,darker)
        self.home._tile_cache=self.fills

    def paint(self,signature,artwork):
        from desktop_ui import VERSION
        h=self.home;t=h.app.tokens;w,hh,style=signature
        images,(brand_font,word_width)=artwork
        _,gradient=style
        style_changed=style!=self.style
        if style_changed:self.set_style(t,gradient)
        for index,(top,bottom,color) in enumerate(self.bands):
            item=self.item(('band',index),'rectangle',(0,round(top*hh),w,round(bottom*hh)),fill=color,width=0,state='normal')
            if index>=self.band_count:h.tag_lower(item)
        self.band_count=max(self.band_count,len(self.bands))
        s=min(w/1024,hh/760)
        for index,curve in enumerate(EDGE_CURVES):
            points=tuple(value*s for value in curve)
            self.item(f'curve{index}','line',points,fill=t['pattern'],width=1)
            self.item(f'end{index}','line',tuple(w-v if j%2==0 else hh-v for j,v in enumerate(points)),fill=t['pattern'],width=1)
        logo_size=max(1,round(74*s))
        x=(w-logo_size-20*s-word_width)/2;y=hh*.148
        logo_item=self.item('logo','image',(x,y),image=images['logo'],anchor='nw')
        brand_item=self.item('brand','text',(x+logo_size+20*s,y+logo_size/2),text='如意',font=brand_font,fill=t['ink'],anchor='w')
        pw=min(w-76*s,916*s);px=(w-pw)/2;py=hh*.275;padding=22*s;gap=16*s;ch=220*s;cw=(pw-2*padding-3*gap)/4
        self.item('panel','polygon',rounded(px,py,px+pw,py+2*ch+2*padding+gap,28*s),smooth=True,splinesteps=12,fill=t['surface'],outline=t['line'],width=1)
        # Keep branding above the panel and the edge curves behind all cards.
        if self.signature is None:h.tag_raise(logo_item);h.tag_raise(brand_item)
        self.item('version','text',(w/2,hh-29*s),text=VERSION,font=('Microsoft YaHei',-max(1,round(12*s))),fill=t['muted'])
        boxes={key:(px+padding+(i%4)*(cw+gap),py+padding+(i//4)*(ch+gap),px+padding+(i%4)*(cw+gap)+cw,py+padding+(i//4)*(ch+gap)+ch) for i,(key,_) in enumerate(h.MODES)}
        boxes['settings']=(w-66*s,18*s,w-24*s,60*s)
        h.hitboxes=boxes
        for key,rect in boxes.items():
            box=tuple(round(v) for v in rect)
            if key not in h._card_tags:
                tag='card:'+key
                h.tag_bind(tag,'<Enter>',lambda _e,k=key:h.enter_card(k))
                h.tag_bind(tag,'<Leave>',lambda _e,k=key:h.leave_card(k))
                h.tag_bind(tag,'<Button-1>',lambda _e,k=key:h.activate(k))
                h._card_tags[key]=tag
            layout=(box,s)
            if style_changed or layout!=self.card_layouts.get(key):
                self.card(key,box,s,t,images)
                self.card_layouts[key]=layout
        h._cache_signature=signature;self.signature=signature;self.style=style
        if style_changed:self.select(h.hover)

    def select(self,key):
        for name in self.home._card_tags:
            fill=self.fills[name][int(name==key)]
            item,_,options=self.items[(name,'tile')]
            if options['fill']!=fill:
                self.home.itemconfigure(item,fill=fill)
                options['fill']=fill

    def card(self,key,box,s,t,images):
        left,top,right,bottom=box
        w,h=right-left,bottom-top
        tags=('surface',self.home._card_tags[key])
        self.item((key,'area'),'rectangle',box,tags=tags,outline='',
                  fill=t['bg'] if key=='settings' else t['surface'])
        self.item((key,'tile'),'polygon',rounded(*box,(9 if key=='settings' else 15)*s),
                  tags=tags,smooth=True,splinesteps=12,outline='',
                  fill=self.fills[key][int(key==self.home.hover)])
        def text(tag,x,y,value,size,anchor='center',color=None):
            self.item((key,tag),'text',(left+x,top+y),tags=tags,text=value,
                      anchor=anchor,fill=color or t['ink'],font=('Microsoft YaHei',-max(1,round(size))))
        def icon(tag,name,x,y):
            self.item((key,tag),'image',(left+x,top+y),tags=tags,image=images[name],anchor='nw')
        if key=='settings':icon('settings','settings-2',12*s,12*s);return
        cx=w/2;iy=h/2-37*s
        if key in ('office','pdf'):
            start=cx-46*s;icon('files','files',start,iy)
            icon('arrow','arrow-right' if key=='office' else 'arrow-left',start+33*s,iy+2*s)
            text('pdf',start+66*s,iy+14*s,'PDF',15*s,'w',t['accent'])
        elif key=='organizer':
            for row in range(2):
                for col in range(2):
                    tag=f'grid{row}{col}';x=left+cx-16*s+col*19*s;y=top+iy-2*s+row*19*s
                    self.item((key,tag),'polygon',rounded(x,y,x+13*s,y+13*s,2*s),tags=tags,
                              smooth=True,splinesteps=8,fill='',outline=t['accent'],width=max(1,1.8*s))
        else:icon('function','image-down' if key=='image' else 'contact-round',cx-16*s,iy-2*s)
        text('title',cx,h/2+32*s,dict(self.home.MODES)[key],15*s)
