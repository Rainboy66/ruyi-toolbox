"""Small rule-chain editor attached only to the PDF organiser window."""
from dataclasses import replace
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk,filedialog,messagebox
from .model import sha256
from .presets import save_preset,load_preset,batch_apply

KINDS={'N-up 宫格':'nup','骑马钉小册子':'booklet','纵向长页':'concat','页码':'page_number','文字水印':'text','图片印章':'stamp','PDF 背景':'background'}
PAPERS={'A4':(595.275591,841.889764),'A4 横向':(841.889764,595.275591),'Letter':(612,792)}
MODES={'适应':'fit','填充':'fill','拉伸':'stretch'}
FORMATS={'页码 / 总数':'{page} / {total}','第…页，共…页':'第{page}页，共{total}页','仅页码':'{page}'}


class RulesWindow(tk.Toplevel):
    def __init__(self,host):
        super().__init__(host);self.host=host;self.title('PDF 拼版与规则');self.geometry('850x650');self.minsize(760,600)
        self.configure(bg=host.colors['bg'])
        self.steps=list(host.document.plan.rules)+[{'op':'overlay','overlay':dict(s)} for s in host.document.plan.overlays]
        self.base_sheet=host.document.plan.sheet
        self.vars={};self.kind=tk.StringVar(value='N-up 宫格')
        top=ttk.Frame(self,padding=12);top.pack(fill='x')
        for title,command in (('打开预设',self.open),('保存预设',self.save),('批量套用',self.batch)):
            ttk.Button(top,text=title,command=command).pack(side='left',padx=4)
        ttk.Label(top,text='步骤从上到下执行').pack(side='right')
        body=ttk.Frame(self,padding=(12,0));body.pack(fill='both',expand=True)
        left=ttk.Frame(body);left.pack(side='left',fill='y',padx=(0,14))
        self.list=tk.Listbox(left,width=24,exportselection=False,bg=host.colors['surface'],fg=host.colors['ink'],
                             selectbackground=host.colors['accent'],selectforeground='white',borderwidth=0,highlightthickness=1,
                             highlightbackground=host.colors.get('line','#dce2eb'));self.list.pack(fill='both',expand=True)
        self.list.bind('<<ListboxSelect>>',self.select)
        move=ttk.Frame(left);move.pack(fill='x',pady=8)
        for title,command in (('上移',lambda:self.move(-1)),('下移',lambda:self.move(1)),('移除',self.remove)):
            ttk.Button(move,text=title,width=6,command=command).pack(side='left')
        right=ttk.Frame(body);right.pack(fill='both',expand=True)
        self.kind_box=ttk.Combobox(right,values=tuple(KINDS),textvariable=self.kind,state='readonly');self.kind_box.pack(fill='x')
        self.kind_box.bind('<<ComboboxSelected>>',lambda e:self.form())
        actions=ttk.Frame(right);actions.pack(side='bottom',fill='x',pady=10)
        ttk.Button(actions,text='添加步骤',command=self.add).pack(side='left',padx=4)
        ttk.Button(actions,text='更新选中步骤',command=self.update_selected).pack(side='left',padx=4)
        field_view=ttk.Frame(right);field_view.pack(fill='both',expand=True)
        self.field_canvas=tk.Canvas(field_view,highlightthickness=0,bg=host.colors['bg'],width=1,height=1)
        field_bar=ttk.Scrollbar(field_view,orient='vertical',command=self.field_canvas.yview);field_bar.pack(side='right',fill='y')
        self.field_canvas.configure(yscrollcommand=field_bar.set);self.field_canvas.pack(fill='both',expand=True)
        self.fields=ttk.Frame(self.field_canvas,padding=(0,12));field_window=self.field_canvas.create_window(0,0,window=self.fields,anchor='nw')
        self.field_canvas.bind('<Configure>',lambda e:self.field_canvas.itemconfigure(field_window,width=e.width))
        self.fields.bind('<Configure>',lambda e:self.field_canvas.configure(scrollregion=(0,0,self.fields.winfo_reqwidth(),self.fields.winfo_reqheight())))
        self.field_canvas.bind('<MouseWheel>',self.field_wheel)
        footer=ttk.Frame(self,padding=12);footer.pack(side='bottom',fill='x',before=body)
        ttk.Label(footer,text='应用后，主网格显示源页；导出文件按规则拼版。').pack(side='left')
        ttk.Button(footer,text='应用到当前方案',command=self.apply).pack(side='right')
        self.form();self.refresh()

    def field_wheel(self,event):
        self.field_canvas.yview_scroll(-int(event.delta/120)*3,'units');return 'break'

    def row(self,key,label,value,choices=None):
        line=ttk.Frame(self.fields);line.pack(fill='x',pady=4)
        ttk.Label(line,text=label,width=16).pack(side='left')
        var=tk.StringVar(value=str(value));self.vars[key]=var
        control=ttk.Combobox(line,textvariable=var,values=choices,state='readonly') if choices else ttk.Entry(line,textvariable=var)
        control.pack(side='left',fill='x',expand=True)
        for widget in (line,*line.winfo_children()):widget.bind('<MouseWheel>',self.field_wheel)
        return line

    def form(self,values=None):
        for child in self.fields.winfo_children():child.destroy()
        self.vars={};s=values or {};kind=KINDS[self.kind.get()]
        if kind in ('nup','booklet','concat'):
            if kind=='nup':
                self.row('n','每页张数',s.get('n',4),('2','4','6','9'))
                self.row('order','排列顺序','列优先' if s.get('order')=='column' else 'Z 字',('Z 字','列优先'))
            if kind!='concat':self.row('size','纸张',next((k for k,v in PAPERS.items() if tuple(s.get('size',()))==v),'A4 横向' if kind=='booklet' or s.get('n')==2 else 'A4'),tuple(PAPERS))
            else:self.row('width','内容宽 mm',s.get('width',595.275591)*25.4/72)
            for key,label,default in (('margin','外边距 mm',24),('gap','页间距 mm',8),('gutter','装订留白 mm',12 if kind=='booklet' else 0)):
                self.row(key,label,f'{s.get(key,default)*25.4/72:.3f}')
            self.row('binding','装订方向','右' if s.get('binding')=='right' else '左',('左','右'))
            if kind=='booklet':self.row('duplex','双面翻转','长边' if s.get('duplex')=='long' else '短边',('短边','长边'))
            if kind=='nup':self.row('fit_mode','内容适配',next(k for k,v in MODES.items() if v==s.get('fit_mode','fit')),tuple(MODES))
            self.row('border','格子边框','有' if s.get('border') else '无',('无','有'))
        elif kind in ('page_number','text'):
            if kind=='page_number':
                self.row('format','页码样式',next((k for k,v in FORMATS.items() if v==s.get('format')), '页码 / 总数'),tuple(FORMATS))
                self.row('start','起始页码',s.get('start',1));self.row('inset','距纸边 mm',f'{s.get("inset",12)*25.4/72:.3f}')
                self.row('align','对齐',{'left':'左','center':'中','right':'右'}[s.get('align','right')],('左','中','右'))
            else:
                self.row('text','水印文字',s.get('text','如意'))
                self.row('angle','角度 °',s.get('angle',30));self.row('opacity','不透明度 %',s.get('opacity',.2)*100)
            self.row('font_size','字号 pt',s.get('font_size',10 if kind=='page_number' else 40));self.row('color','颜色 HEX',s.get('color','#334455'))
        else:
            line=self.row('path','图片文件' if kind=='stamp' else 'PDF 文件',s.get('path',''))
            ttk.Button(line,text='选择',command=lambda:self.pick(kind)).pack(side='right')
            if kind=='background':self.row('page_index','背景页码',s.get('page_index',0)+1)
            else:
                r=s.get('rect',(425.2,34,510.2,119))
                for key,label,value in zip(('x','y','width','height'),('左侧 mm','底部 mm','宽 mm','高 mm'),(r[0],r[1],r[2]-r[0],r[3]-r[1])):self.row(key,label,f'{value*25.4/72:.3f}')
            self.row('opacity','不透明度 %',s.get('opacity',1)*100)
            self.row('fit_mode','内容适配',next(k for k,v in MODES.items() if v==s.get('fit_mode','fit')),tuple(MODES))

    def pick(self,kind):
        path=filedialog.askopenfilename(parent=self,filetypes=[('PDF','*.pdf')] if kind=='background' else [('图片','*.png *.jpg *.jpeg *.webp')])
        if path:self.vars['path'].set(path)

    def spec(self):
        from .rule_spec import rule
        v={k:var.get() for k,var in self.vars.items()};kind=KINDS[self.kind.get()]
        if kind in ('nup','booklet','concat'):
            s={'op':kind,'margin':float(v['margin'])*72/25.4,'gap':float(v['gap'])*72/25.4,'gutter':float(v['gutter'])*72/25.4,
               'binding':'left' if v['binding']=='左' else 'right','border':v['border']=='有'}
            if kind!='concat':s['size']=PAPERS[v['size']]
            else:s['width']=float(v['width'])*72/25.4
            if kind=='nup':s.update(n=int(v['n']),order='z' if v['order']=='Z 字' else 'column',fit_mode=MODES[v['fit_mode']])
            if kind=='booklet':s['duplex']='short' if v['duplex']=='短边' else 'long'
        else:
            layer={'kind':kind}
            for key in ('font_size','angle'):
                if key in v:layer[key]=float(v[key])
            for key in ('text','color'):
                if key in v:layer[key]=v[key]
            if 'opacity' in v:layer['opacity']=float(v['opacity'])/100
            if kind=='page_number':layer.update(format=FORMATS[v['format']],start=int(v['start']),inset=float(v['inset'])*72/25.4,align={'左':'left','中':'center','右':'right'}[v['align']])
            if kind in ('stamp','background'):
                path=Path(v['path']).resolve(strict=True);layer.update(path=str(path),sha256=sha256(path),fit_mode=MODES[v['fit_mode']])
                if kind=='background':layer['page_index']=int(v['page_index'])-1
                else:
                    x,y,w,h=(float(v[k])*72/25.4 for k in ('x','y','width','height'));layer['rect']=(x,y,x+w,y+h)
            s={'op':'overlay','overlay':layer}
        return rule(s)

    def guarded(self,action):
        if self.host.busy:return
        try:action()
        except (ValueError,OSError,KeyError,TypeError) as error:messagebox.showerror('规则未应用',str(error),parent=self)

    def refresh(self,selected=None):
        self.list.delete(0,'end')
        for i,s in enumerate(self.steps):
            kind=s['overlay']['kind'] if s['op']=='overlay' else s['op']
            label=next(k for k,v in KINDS.items() if v==kind)
            if kind=='nup':label+=f' · {s.get("n",4)}'
            self.list.insert('end',f'{i+1}. {label}')
        if selected is not None:self.list.selection_set(selected)

    def select(self,event=None):
        if not self.list.curselection():return
        step=self.steps[self.list.curselection()[0]];kind=step['overlay']['kind'] if step['op']=='overlay' else step['op']
        self.kind.set(next(k for k,v in KINDS.items() if v==kind));self.form(step['overlay'] if step['op']=='overlay' else step)

    def add(self):self.guarded(lambda:(self.steps.append(self.spec()),self.refresh(len(self.steps)-1)))
    def update_selected(self):
        def action():
            if self.list.curselection():
                i=self.list.curselection()[0];self.steps[i]=self.spec();self.refresh(i)
        self.guarded(action)
    def move(self,direction):
        if self.host.busy or not self.list.curselection():return
        i=self.list.curselection()[0];j=i+direction
        if 0<=j<len(self.steps):self.steps[i],self.steps[j]=self.steps[j],self.steps[i];self.refresh(j)
    def remove(self):
        if not self.host.busy and self.list.curselection():self.steps.pop(self.list.curselection()[0]);self.refresh()
    def plan(self):return replace(self.host.document.plan,sheet=self.base_sheet,rules=tuple(self.steps),overlays=())
    def apply(self):
        def action():
            self.host.document.commit(self.plan(),self.host.document.selection);self.host.changed();self.host.status.set('导出规则已应用，网格显示源页。')
        self.guarded(action)
    def save(self):
        path=filedialog.asksaveasfilename(parent=self,defaultextension='.ruyiplan.json',initialfile='拼版规则.ruyiplan.json',filetypes=[('如意方案','*.ruyiplan.json')])
        if path:self.guarded(lambda:save_preset(self.plan(),path))
    def open(self):
        path=filedialog.askopenfilename(parent=self,filetypes=[('如意方案','*.ruyiplan.json')])
        if path:
            def action():
                preset=load_preset(path);self.steps=list(preset.rules)+[{'op':'overlay','overlay':dict(s)} for s in preset.overlays]
                self.base_sheet=preset.sheet;self.refresh()
            self.guarded(action)
    def batch(self):
        if self.host.busy:return
        files=filedialog.askopenfilenames(parent=self,title='选择独立套用预设的 PDF',filetypes=[('PDF','*.pdf')])
        if not files:return
        parent=filedialog.askdirectory(parent=self,title='保存批量结果')
        if not parent:return
        def action():
            plan=self.plan();directory=Path(parent)/('如意批量拼版_'+datetime.now().strftime('%Y%m%d_%H%M%S_%f'))
            self.host.job('正在批量套用规则…',lambda:batch_apply(files,plan,directory,self.host.progress,self.host.cancel),
                          lambda result:self.host.status.set(f'批量完成：{sum(r["ok"] for r in result)}份成功，{sum(not r["ok"] for r in result)}份失败；详见批量结果.json'))
        self.guarded(action)


def open_rules(host):
    window=getattr(host,'rules_window',None)
    if window is None or not window.winfo_exists():host.rules_window=RulesWindow(host)
    else:window.deiconify();window.lift()
