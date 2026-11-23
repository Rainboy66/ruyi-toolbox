"""Viewport-only Canvas items and images; drag displays one placeholder."""
import math,time,tkinter as tk
from io import BytesIO
from tkinter import ttk
from PIL import Image,ImageTk
from window_paint import window_is_resizing
from .thumbnails import ThumbnailService
from .grid_input import CanvasInput


class PageGrid(ttk.Frame):
    cell_w=216;row_h=286
    def __init__(self,parent,host):
        super().__init__(parent);self.host=host;self.cells={};self.photos={};self.columns=1
        self.service=ThumbnailService();self.paint_timer=None;self.poll_timer=None;self.auto_timer=None
        self.resize_timer=None;self.viewport=None
        self.drag_start=None;self.dragging=False;self.drag_boundary=None;self.anchor=0;self.closed=False
        self.visible_indices=();self.paint_times=[];self.scroll_times=[];self.max_items=0;self.max_visible=0
        self.canvas=tk.Canvas(self,bg=host.colors['bg'],highlightthickness=0,takefocus=True)
        self.input=CanvasInput(self.canvas)
        self.bar=ttk.Scrollbar(self,orient='vertical',command=self.scroll)
        self.canvas.configure(yscrollcommand=self.scrolled);self.bar.pack(side='right',fill='y');self.canvas.pack(fill='both',expand=True)
        self.canvas.bind('<Configure>',self.schedule_resize)
        self.canvas.bind('<MouseWheel>',self.wheel)
        self.canvas.bind('<Button-1>',self.press);self.canvas.bind('<B1-Motion>',self.motion);self.canvas.bind('<ButtonRelease-1>',self.release)
        self.poll_timer=self.after(16,self.poll)

    @property
    def doc(self):return self.host.document

    def schedule_resize(self,event=None):
        if event is not None:
            size=(event.width,event.height)
            if self.viewport==size:return
            self.viewport=size
        if self.resize_timer is not None:
            self.after_cancel(self.resize_timer);self.resize_timer=None
        if not window_is_resizing(self):
            self.resize_timer=self.after_idle(self.configure_view)

    def configure_view(self):
        if self.resize_timer is not None:self.after_cancel(self.resize_timer);self.resize_timer=None
        if window_is_resizing(self):return
        columns=max(1,(self.canvas.winfo_width()-24)//self.cell_w)
        if columns!=self.columns:
            self.canvas.delete('all');self.cells.clear();self.photos.clear()
            self.columns=columns
        height=24+math.ceil(len(self.doc.plan.pages)/self.columns)*self.row_h
        self.canvas.configure(scrollregion=(0,0,max(self.canvas.winfo_width(),240),max(height,self.canvas.winfo_height())))
        self.schedule()

    def scrolled(self,first,last):
        self.bar.set(first,last)
        if self.resize_timer is None:self.schedule()
    def scroll(self,*args):self.canvas.yview(*args);self.schedule()
    def wheel(self,event):
        start=time.perf_counter();self.canvas.yview_scroll(-int(event.delta/120)*3,'units');self.schedule()
        self.scroll_times.append((time.perf_counter()-start)*1000);return 'break'

    def schedule(self):
        if not self.closed and self.paint_timer is None:self.paint_timer=self.after_idle(self.paint)

    def invalidate(self):
        self.cancel_drag();self.canvas.delete('all');self.cells.clear();self.photos.clear();self.configure_view()

    def paint(self):
        self.paint_timer=None
        if self.closed or window_is_resizing(self):return
        start=time.perf_counter();top=self.canvas.canvasy(0);bottom=top+self.canvas.winfo_height()
        first=max(0,int((top-12)//self.row_h));last=max(first,int((bottom-12)//self.row_h))
        wanted=[i for row in range(first,last+1) for i in range(row*self.columns,min((row+1)*self.columns,len(self.doc.plan.pages)))
                if 12+row*self.row_h+274>top and 12+row*self.row_h<bottom]
        keep=set(wanted)
        for i in list(self.cells):
            if i not in keep:self.canvas.delete(f'page:{i}');self.cells.pop(i);self.photos.pop(i,None)
        requests=[];selected=set(self.doc.selection)
        for i in wanted:
            ref=self.doc.plan.pages[i];source=self.doc.sources[ref.source_id]
            request=self.service.request(ref,source,self.doc.plan.sheet);requests.append(request)
            if i not in self.cells:self.draw_cell(i,request['key'],source,ref)
            self.canvas.itemconfigure(self.cells[i]['border'],outline=self.host.colors['accent'] if i in selected else '#dce2eb',width=2 if i in selected else 1)
            data=self.service.get(request['key'])
            if data is not None and self.cells[i].get('loaded')!=request['key']:
                self.set_image(i,request['key'],data)
        self.visible_indices=tuple(wanted);self.service.visible(requests)
        self.max_visible=max(self.max_visible,len(wanted));self.max_items=max(self.max_items,len(self.canvas.find_all()))
        if self.dragging:self.placeholder()
        self.paint_times.append((time.perf_counter()-start)*1000)

    def draw_cell(self,index,key,source,ref):
        x=12+(index%self.columns)*self.cell_w;y=12+(index//self.columns)*self.row_h;tag=f'page:{index}'
        c=self.canvas;border=c.create_rectangle(x,y,x+200,y+274,fill=self.host.colors['surface'],outline='#dce2eb',tags=tag)
        image=c.create_image(x+100,y+12,anchor='n',tags=tag)
        loading=c.create_text(x+100,y+115,text='预览中…',fill=self.host.colors['muted'],font=('Microsoft YaHei UI',9),tags=tag)
        c.create_text(x+12,y+244,text=f'第 {index+1} 页',anchor='w',fill=self.host.colors['ink'],font=('Microsoft YaHei UI',10),tags=tag)
        if ref.rotate:c.create_text(x+188,y+244,text=f'{ref.rotate}°',anchor='e',fill=self.host.colors['muted'],font=('Microsoft YaHei UI',9),tags=tag)
        name=source.name if len(source.name)<=18 else source.name[:15]+'…'
        c.create_text(x+12,y+263,text=f'{name} · 源 {ref.page_index+1}',anchor='w',fill=self.host.colors['muted'],font=('Microsoft YaHei UI',8),tags=tag)
        self.cells[index]={'border':border,'image':image,'loading':loading,'key':key,'loaded':None}

    def set_image(self,index,key,data):
        cell=self.cells[index]
        if data:
            with Image.open(BytesIO(data)) as im:self.photos[index]=ImageTk.PhotoImage(im,master=self.canvas)
            self.canvas.itemconfigure(cell['image'],image=self.photos[index]);self.canvas.itemconfigure(cell['loading'],text='')
        else:self.canvas.itemconfigure(cell['loading'],text='预览不可用')
        cell['loaded']=key

    def poll(self):
        if self.closed:return
        if self.service.receive():self.schedule()
        self.poll_timer=self.after(16,self.poll)

    def index_at(self,x,y):
        x=self.canvas.canvasx(x)-12;y=self.canvas.canvasy(y)-12
        if x<0 or y<0:return None
        col=int(x//self.cell_w);row=int(y//self.row_h);index=row*self.columns+col
        if col>=self.columns or x%self.cell_w>200 or y%self.row_h>274 or index>=len(self.doc.plan.pages):return None
        return index

    def press(self,event):
        self.canvas.focus_set()
        if self.host.busy:return
        i=self.index_at(event.x,event.y)
        if i is None:self.doc.select(());self.host.selection_changed();self.schedule();return
        selected=set(self.doc.selection);control=bool(event.state&4);shift=bool(event.state&1)
        if shift:selected=set(range(min(self.anchor,i),max(self.anchor,i)+1))|(selected if control else set())
        elif control:
            if i in selected:selected.remove(i)
            else:selected.add(i)
            self.anchor=i
        else:
            if i not in selected:selected={i}
            self.anchor=i
        self.doc.select(selected);self.host.selection_changed();self.schedule()
        self.drag_start=(event.x,event.y,i,self.doc.revision,tuple(self.doc.selection),not(control or shift))

    def boundary(self,x,y):
        x=max(12,min(self.canvas.winfo_width()-1,x))
        cy=self.canvas.canvasy(y)-12;row=max(0,int(cy//self.row_h));col=max(0,min(self.columns-1,int((x-12)//self.cell_w)))
        return max(0,min(len(self.doc.plan.pages),row*self.columns+col+(1 if (x-12)%self.cell_w>100 else 0)))

    def motion(self,event):
        if not self.drag_start:return
        x,y,_,revision,indices,_=self.drag_start
        if self.doc.revision!=revision:self.cancel_drag();return
        if not self.dragging and abs(event.x-x)+abs(event.y-y)<7:return
        self.dragging=True;self.drag_boundary=self.boundary(event.x,event.y);self.last_pointer=(event.x,event.y);self.placeholder()
        preview=getattr(self.host,'preview',None)
        if preview is not None:preview.highlight(event)
        if (event.y<24 or event.y>self.canvas.winfo_height()-24) and self.auto_timer is None:self.auto_timer=self.after(40,self.auto_scroll)

    def auto_scroll(self):
        self.auto_timer=None
        if not self.dragging:return
        x,y=self.last_pointer
        direction=-1 if y<24 else (1 if y>self.canvas.winfo_height()-24 else 0)
        if direction:
            self.canvas.yview_scroll(direction,'units');self.drag_boundary=self.boundary(x,y);self.schedule();self.placeholder()
            self.auto_timer=self.after(40,self.auto_scroll)

    def placeholder(self):
        self.canvas.delete('placeholder')
        if self.drag_boundary is None:return
        n=self.drag_boundary;x=12+(n%self.columns)*self.cell_w;y=12+(n//self.columns)*self.row_h
        if n==len(self.doc.plan.pages) and n and n%self.columns==0:
            x=12+(self.columns-1)*self.cell_w+204;y-=self.row_h
        self.canvas.create_rectangle(x-5,y,x+3,y+274,fill=self.host.colors['accent'],outline='',tags='placeholder')
        self.canvas.tag_raise('placeholder')

    def release(self,event):
        start=self.drag_start
        if start and self.dragging and start[3]==self.doc.revision:
            preview=getattr(self.host,'preview',None)
            if preview is not None and preview.drop(start[4],event):self.cancel_drag();return
            boundary=self.boundary(event.x,event.y);self.cancel_drag();self.doc.move(start[4],boundary);self.host.changed()
        else:
            if start and start[5]:self.doc.select((start[2],));self.host.selection_changed();self.schedule()
            self.cancel_drag()

    def cancel_drag(self):
        if self.auto_timer is not None:self.after_cancel(self.auto_timer);self.auto_timer=None
        self.canvas.delete('placeholder');self.drag_start=None;self.dragging=False;self.drag_boundary=None
        preview=getattr(self.host,'preview',None)
        if preview is not None:preview.canvas.itemconfigure(preview.mark,state='hidden')

    def shutdown(self):
        self.closed=True;self.input.leave();self.cancel_drag()
        for timer in (self.paint_timer,self.poll_timer,self.resize_timer):
            if timer is not None:self.after_cancel(timer)
        self.service.close();self.photos.clear()
