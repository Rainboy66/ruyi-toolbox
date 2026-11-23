"""Live output-sheet preview; source-grid and sheet-slot drags share OutputPlan."""
import base64,queue,tkinter as tk
from dataclasses import replace
from io import BytesIO
from tkinter import ttk
from PIL import Image,ImageTk
from window_paint import window_is_resizing
from .preview_service import PreviewService
from .preview_layout import sheet_view,swap_pages
from .rule_spec import page_count


class SheetPreview(ttk.Frame):
    def __init__(self,parent,host):
        super().__init__(parent);self.host=host;self.page=0;self.key=None;self.slots=();self.size=None
        self.bitmap=None;self.photo=None;self.bounds=None;self.closed=False;self.render_timer=None;self.paint_timer=None;self.drag_index=None;self.zoom=1.0
        self.photo_size=None;self.scroll_size=None
        self.viewport=None
        self.service=PreviewService()
        header=ttk.Frame(self);header.pack(fill='x',padx=12,pady=(6,8))
        ttk.Label(header,text='成品预览').pack(side='left')
        self.page_label=ttk.Label(header,text='0 / 0');self.page_label.pack(side='right',padx=8)
        self.next=ttk.Button(header,text='›',width=3,command=lambda:self.turn(1));self.next.pack(side='right')
        self.previous=ttk.Button(header,text='‹',width=3,command=lambda:self.turn(-1));self.previous.pack(side='right')
        footer=ttk.Frame(self);footer.pack(side='bottom',fill='x',padx=12,pady=6)
        ttk.Button(footer,text='−',width=3,command=lambda:self.set_zoom(.8)).pack(side='left')
        ttk.Button(footer,text='适合窗口',command=self.fit).pack(side='left',padx=4)
        ttk.Button(footer,text='+',width=3,command=lambda:self.set_zoom(1.25)).pack(side='left')
        self.message=ttk.Label(footer,text='');self.message.pack(side='right')
        viewport=ttk.Frame(self);viewport.pack(fill='both',expand=True,padx=8)
        self.canvas=tk.Canvas(viewport,bg=host.colors.get('soft','#edf1f6'),highlightthickness=0,takefocus=True)
        vertical=ttk.Scrollbar(viewport,orient='vertical',command=self.canvas.yview);vertical.pack(side='right',fill='y')
        horizontal=ttk.Scrollbar(viewport,orient='horizontal',command=self.canvas.xview);horizontal.pack(side='bottom',fill='x')
        self.canvas.configure(yscrollcommand=vertical.set,xscrollcommand=horizontal.set)
        self.canvas.pack(fill='both',expand=True)
        self.picture=self.canvas.create_image(0,0,anchor='nw')
        self.notice=self.canvas.create_text(100,100,text='添加 PDF 后查看拼版',fill=host.colors['muted'],font=('Microsoft YaHei UI',10))
        self.mark=self.canvas.create_rectangle(0,0,0,0,outline=host.colors['accent'],width=3,dash=(6,3),state='hidden')
        self.canvas.bind('<Configure>',self.schedule_resize)
        self.canvas.bind('<Button-1>',self.press);self.canvas.bind('<B1-Motion>',self.motion);self.canvas.bind('<ButtonRelease-1>',self.release)
        self.canvas.bind('<Control-MouseWheel>',lambda e:self.set_zoom(1.1 if e.delta>0 else 1/1.1))
        self.canvas.bind('<MouseWheel>',self.wheel)
        self.poll_timer=self.after(30,self.poll)

    def changed(self):
        if self.closed:return
        if self.render_timer is not None:self.after_cancel(self.render_timer)
        count=page_count(self.host.document.plan);self.page=max(0,min(self.page,max(0,count-1)))
        self.page_label.configure(text=f'{self.page+1 if count else 0} / {count}')
        self.previous.state(['!disabled'] if self.page else ['disabled']);self.next.state(['!disabled'] if self.page<count-1 else ['disabled'])
        self.key=f'{self.host.document.revision}:{self.page}'
        self.slots=();self.size=None;self.canvas.itemconfigure(self.mark,state='hidden');self.message.configure(text='正在更新…' if count else '')
        if not count:
            self.clear();return
        self.render_timer=self.after(100,self.request)

    def request(self):
        self.render_timer=None;doc=self.host.document
        local,index,size,slots,count=sheet_view(doc.plan,doc.sources,self.page)
        self.pending_view=(self.key,size,slots)
        self.service.request({'key':self.key,'revision':self.key if slots else f'{doc.revision}:all','page':index,
                              '_plan':local,'_sources':doc.sources.copy()})

    def clear(self):
        self.photo=None;self.photo_size=None;self.bounds=None
        if self.bitmap is not None:self.bitmap.close();self.bitmap=None
        self.canvas.itemconfigure(self.picture,image='');self.canvas.itemconfigure(self.notice,text='添加 PDF 后查看拼版')

    def poll(self):
        if self.closed:return
        for _ in range(3):
            try:result=self.service.results.get_nowait()
            except queue.Empty:break
            if result['key'] not in (self.key,None):continue
            if 'error' in result:
                self.clear();self.canvas.itemconfigure(self.notice,text='此方案暂不能预览');self.message.configure(text='请检查规则')
                self.host.status.set('成品预览：'+result['error']);continue
            with Image.open(BytesIO(base64.b64decode(result['png']))) as image:bitmap=image.convert('RGB')
            if self.bitmap is not None:self.bitmap.close()
            self.bitmap=bitmap;self.photo_size=None
            if self.pending_view[0]==self.key:_,self.size,self.slots=self.pending_view
            self.canvas.itemconfigure(self.notice,text='');self.message.configure(text='');self.schedule_paint()
        self.poll_timer=self.after(30,self.poll)

    def turn(self,step):
        count=page_count(self.host.document.plan);new=max(0,min(self.page+step,max(0,count-1)))
        if new!=self.page:self.page=new;self.changed()

    def fit(self):self.zoom=1.;self.canvas.xview_moveto(0);self.canvas.yview_moveto(0);self.schedule_paint()
    def wheel(self,event):
        if self.zoom>1:self.canvas.yview_scroll(-int(event.delta/120)*3,'units')
        else:self.turn(-1 if event.delta>0 else 1)
        return 'break'
    def set_zoom(self,factor):self.zoom=max(.5,min(3,self.zoom*factor));self.schedule_paint();return 'break'
    def schedule_resize(self,event=None):
        if event is not None:
            size=(event.width,event.height)
            if size==self.viewport:return
            self.viewport=size
        self.schedule_paint(16)

    def schedule_paint(self,delay=None):
        if self.closed:return
        if self.paint_timer is not None:
            self.after_cancel(self.paint_timer);self.paint_timer=None
        if not window_is_resizing(self):
            self.paint_timer=self.after_idle(self.paint) if delay is None else self.after(delay,self.paint)

    def paint(self):
        if self.paint_timer is not None:self.after_cancel(self.paint_timer)
        self.paint_timer=None
        if window_is_resizing(self):return
        w,h=self.canvas.winfo_width(),self.canvas.winfo_height()
        self.canvas.coords(self.notice,w/2,h/2)
        if self.bitmap is None:return
        scale=min(max(1,w-32)/self.bitmap.width,max(1,h-24)/self.bitmap.height)*self.zoom
        iw,ih=max(1,round(self.bitmap.width*scale)),max(1,round(self.bitmap.height*scale))
        if (iw,ih)!=self.photo_size:
            with self.bitmap.resize((iw,ih),Image.Resampling.BILINEAR) as resized:
                self.photo=ImageTk.PhotoImage(resized,master=self.canvas)
            self.canvas.itemconfigure(self.picture,image=self.photo)
            self.photo_size=(iw,ih)
        total_w,total_h=max(w,iw+32),max(h,ih+24)
        if (total_w,total_h)!=self.scroll_size:
            self.canvas.configure(scrollregion=(0,0,total_w,total_h))
            self.scroll_size=(total_w,total_h)
        x,y=(total_w-iw)/2,max(8,(total_h-ih)/2)
        if (x,y,iw,ih)!=self.bounds:
            self.canvas.coords(self.picture,x,y)
            self.bounds=(x,y,iw,ih)

    def slot_at(self,x_root,y_root):
        if not self.bounds or not self.size:return None
        x,y=x_root-self.canvas.winfo_rootx(),y_root-self.canvas.winfo_rooty();bx,by,bw,bh=self.bounds
        if not (0<=x<self.canvas.winfo_width() and 0<=y<self.canvas.winfo_height()):return None
        x,y=self.canvas.canvasx(x),self.canvas.canvasy(y)
        px=(x-bx)/bw*self.size[0];py=(1-(y-by)/bh)*self.size[1]
        for index,rect in self.slots:
            if rect[0]<=px<=rect[2] and rect[1]<=py<=rect[3]:return index

    def highlight(self,event):
        index=self.slot_at(event.x_root,event.y_root)
        self.canvas.itemconfigure(self.mark,state='hidden')
        if index is not None:
            rect=dict(self.slots)[index];bx,by,bw,bh=self.bounds;sw,sh=self.size
            self.canvas.coords(self.mark,bx+rect[0]/sw*bw,by+(1-rect[3]/sh)*bh,bx+rect[2]/sw*bw,by+(1-rect[1]/sh)*bh)
            self.canvas.itemconfigure(self.mark,state='normal');self.canvas.tag_raise(self.mark)
        return index

    def drop(self,indices,event):
        target=self.slot_at(event.x_root,event.y_root);self.canvas.itemconfigure(self.mark,state='hidden')
        if target is None:return False
        if self.host.busy:return True
        if len(indices)==1:changed=swap_pages(self.host.document,indices[0],target)
        else:changed=self.host.document.move(indices,target) is not False
        if changed:self.host.changed()
        return True

    def press(self,event):
        if self.host.busy:return
        self.drag_index=self.slot_at(event.x_root,event.y_root);self.canvas.focus_set()
    def motion(self,event):
        if self.drag_index is not None:self.highlight(event)
    def release(self,event):
        if self.drag_index is not None:self.drop((self.drag_index,),event)
        self.drag_index=None;self.canvas.itemconfigure(self.mark,state='hidden')

    def shutdown(self):
        self.closed=True;self.service.close()
        for timer in (self.render_timer,self.paint_timer,self.poll_timer):
            if timer is not None:self.after_cancel(timer)
        if self.bitmap is not None:self.bitmap.close()
        self.photo=None
