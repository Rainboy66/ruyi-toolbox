"""PDF organizer page. All edits go through PlanDocument."""
from dataclasses import replace
from pathlib import Path
from datetime import datetime
import queue,threading,tkinter as tk
from tkinter import filedialog,messagebox,ttk
from .model import PlanDocument,Sheet,inspect_sources
from .exporter import export_pdf,split_pdf
from .grid import PageGrid

PAPERS={'原尺寸':None,'A4':(595.275591,841.889764),'A4 横向':(841.889764,595.275591),'Letter':(612,792)}
MODES={'适应':'fit','填充':'fill','拉伸':'stretch'}


class OrganizerPane(ttk.Frame):
    def __init__(self,parent,colors=None,show_home=None):
        super().__init__(parent,style='Organizer.TFrame')
        self.apply_theme(colors or {'bg':'#fafbfe','surface':'#ffffff','ink':'#242c3c','muted':'#6b7485','accent':'#374767'})
        self.document=PlanDocument();self.busy=False;self.closed=False;self.closing=False;self.dirty=False;self.plan_path=None
        self.events=queue.Queue();self.cancel=threading.Event();self.controls=[];self.poll_timer=None;self.last_output=None
        self.after_close=None
        self.body=ttk.Frame(self,style='Organizer.TFrame');self.body.pack(fill='both',expand=True)
        top=ttk.Frame(self.body,padding=(18,14,18,6),style='Organizer.TFrame');top.pack(fill='x')
        if show_home is not None:
            ttk.Button(top,text='‹  首页',style='Quiet.TButton',command=show_home).pack(side='left',padx=(0,18))
        ttk.Label(top,text='PDF 合并拼凑',font=('Microsoft YaHei UI',16,'bold'),style='Organizer.TLabel').pack(side='left')
        self.counter=ttk.Label(top,style='Organizer.TLabel');self.counter.pack(side='right')
        tools=ttk.Frame(self.body,padding=(18,5),style='Organizer.TFrame');tools.pack(fill='x')
        self.button(tools,'添加 PDF',self.add_files);self.button(tools,'打开方案',self.open_plan);self.button(tools,'保存方案',self.save_plan)
        self.button(tools,'拼版与规则',self.open_rules)
        self.redo_button=self.button(tools,'重做',lambda:self.history(True),side='right');self.undo_button=self.button(tools,'撤销',self.history,side='right')
        edit=ttk.Frame(self.body,padding=(18,5),style='Organizer.TFrame');edit.pack(fill='x')
        self.button(edit,'全选',self.select_all);self.button(edit,'清除选择',self.clear_selection)
        self.selection_buttons=[self.button(edit,'删除',lambda:self.edit('delete')),self.button(edit,'复制',lambda:self.edit('duplicate')),self.button(edit,'旋转 90°',lambda:self.edit('rotate'))]
        options=ttk.Frame(self.body,padding=(18,8),style='Organizer.TFrame');options.pack(fill='x')
        ttk.Label(options,text='纸张',style='Organizer.TLabel').pack(side='left')
        self.papers=PAPERS.copy()
        self.paper=tk.StringVar(value='原尺寸');self.paper_box=ttk.Combobox(options,textvariable=self.paper,values=tuple(self.papers),state='readonly',width=16);self.paper_box.pack(side='left',padx=(5,16));self.controls.append(self.paper_box)
        ttk.Label(options,text='适配',style='Organizer.TLabel').pack(side='left')
        self.mode=tk.StringVar(value='适应');box=ttk.Combobox(options,textvariable=self.mode,values=tuple(MODES),state='readonly',width=6);box.pack(side='left',padx=(5,16));self.controls.append(box)
        ttk.Label(options,text='页边距 mm',style='Organizer.TLabel').pack(side='left')
        self.margin=tk.StringVar(value='0');spin=ttk.Spinbox(options,textvariable=self.margin,from_=0,to=100,width=5);spin.pack(side='left',padx=(5,12));self.controls.append(spin)
        self.button(options,'应用到全部',self.apply_options)
        footer=ttk.Frame(self.body,padding=(18,10,18,14),style='Organizer.TFrame');footer.pack(side='bottom',fill='x')
        self.button(footer,'合并导出 PDF',self.export_all,side='right');self.extract_button=self.button(footer,'提取所选',self.extract,side='right')
        self.button(footer,'按页数拆分',self.split)
        self.split_count=tk.StringVar(value='20');split_spin=ttk.Spinbox(footer,textvariable=self.split_count,from_=1,to=10000,width=5);split_spin.pack(side='left',padx=(0,8));self.controls.append(split_spin)
        self.cancel_button=ttk.Button(footer,text='停止',command=lambda:self.cancel.set(),state='disabled');self.cancel_button.pack(side='left')
        self.status=tk.StringVar(value='添加 PDF 后，拖动页面调整顺序。')
        ttk.Label(self.body,textvariable=self.status,style='Organizer.TLabel',padding=(20,4)).pack(side='bottom',fill='x')
        quick=ttk.Frame(self.body,padding=(18,4),style='Organizer.TFrame');quick.pack(fill='x')
        ttk.Label(quick,text='拼页',style='Organizer.TLabel').pack(side='left')
        self.quick_mode=tk.StringVar(value='原页')
        selector=ttk.Combobox(quick,textvariable=self.quick_mode,values=('原页','2 合 1','4 合 1','6 合 1','9 合 1'),state='readonly',width=9)
        selector.pack(side='left',padx=8);selector.bind('<<ComboboxSelected>>',self.quick_layout);self.controls.append(selector)
        self.workspace=ttk.Panedwindow(self.body,orient='horizontal');self.workspace.pack(fill='both',expand=True,padx=8,pady=6)
        left=ttk.Frame(self.workspace);self.workspace.add(left,weight=1)
        ttk.Label(left,text='源页面',padding=(12,6)).pack(anchor='w')
        self.grid_view=PageGrid(left,self);self.grid_view.pack(fill='both',expand=True)
        from .preview_ui import SheetPreview
        self.preview=SheetPreview(self.workspace,self);self.workspace.add(self.preview,weight=2)
        self.shortcut_bindings={}
        for key,command in (('<Delete>',lambda:self.edit('delete')),('<Control-d>',lambda:self.edit('duplicate')),('<Key-r>',lambda:self.edit('rotate')),
                            ('<Control-z>',self.history),('<Control-y>',lambda:self.history(True)),('<Control-a>',self.select_all),('<Escape>',self.clear_selection)):
            self.shortcut_bindings[key]=self.winfo_toplevel().bind(key,lambda event,action=command:self.key(event,action),add='+')
        self.selection_changed();self.poll_timer=self.after(30,self.poll)

    def apply_theme(self, colors):
        if colors == getattr(self,'colors',None):
            return
        self.colors=colors
        style=ttk.Style(self)
        style.configure('Organizer.TFrame',background=colors['bg'])
        style.configure('Organizer.TLabel',background=colors['bg'],foreground=colors['ink'])
        if hasattr(self,'grid_view'):
            self.grid_view.canvas.configure(background=colors['bg'])
            self.grid_view.invalidate()
            self.preview.canvas.configure(background=colors.get('soft',colors['bg']))
            self.preview.canvas.itemconfigure(self.preview.notice,fill=colors['muted'])
            self.preview.canvas.itemconfigure(self.preview.mark,outline=colors['accent'])

    def on_hide(self):
        self.grid_view.cancel_drag()
        window=getattr(self,'rules_window',None)
        if window is not None and window.winfo_exists():
            window.withdraw()

    def button(self,parent,title,command,side='left'):
        button=ttk.Button(parent,text=title,command=command);button.pack(side=side,padx=(0,6));self.controls.append(button);return button

    def key(self,event,action):
        if not self.winfo_ismapped():
            return
        if isinstance(event.widget,(ttk.Entry,ttk.Spinbox,ttk.Combobox,tk.Entry)):return
        if not self.busy:action()
        return 'break'

    def select_all(self):self.document.select(range(len(self.document.plan.pages)));self.selection_changed();self.grid_view.schedule()
    def clear_selection(self):self.document.select(());self.selection_changed();self.grid_view.cancel_drag();self.grid_view.schedule()

    def selection_changed(self):
        from .rule_spec import page_count
        n=len(self.document.plan.pages);selected=len(self.document.selection)
        label=f'来源 {n} 页 · 导出 {page_count(self.document.plan)} 页' if self.document.plan.rules else f'共 {n} 页'
        self.counter.configure(text=f'{label} · 已选 {selected} 页')
        if hasattr(self,'quick_mode'):
            layouts=[r for r in self.document.plan.rules if r['op']!='overlay']
            self.quick_mode.set(f'{layouts[0].get("n",4)} 合 1' if len(layouts)==1 and layouts[0]['op']=='nup' else ('规则拼版' if layouts else '原页'))
        for button in self.selection_buttons+[self.extract_button]:button.state(['!disabled'] if selected and not self.busy else ['disabled'])
        self.undo_button.state(['!disabled'] if self.document.undo_stack and not self.busy else ['disabled'])
        self.redo_button.state(['!disabled'] if self.document.redo_stack and not self.busy else ['disabled'])

    def changed(self):
        self.dirty=True;self.selection_changed();self.grid_view.invalidate();self.preview.changed();self.status.set('顺序已更新。')

    def quick_layout(self,_event=None):
        if self.busy:return
        plan=self.document.plan
        if len([r for r in plan.rules if r['op']!='overlay'])>1 or any(r['op'] in ('booklet','concat') for r in plan.rules):
            self.status.set('当前使用多步规则，请在“拼版与规则”中调整。');return
        rules=tuple(r for r in plan.rules if r['op']!='nup')
        if self.quick_mode.get()!='原页':
            n=int(self.quick_mode.get().split()[0]);size=PAPERS['A4 横向' if n==2 else 'A4']
            rules=({'op':'nup','n':n,'size':size,'order':'z','margin':24,'gap':8,'border':False},)+rules
        if self.document.commit(replace(plan,rules=rules),self.document.selection):self.changed()

    def edit(self,operation):
        if not self.busy and self.document.selection:getattr(self.document,operation)();self.changed()

    def history(self,redo=False):
        if self.busy:return
        if (self.document.redo if redo else self.document.undo)():
            self.sync_options()
            self.changed()

    def sync_options(self):
        sheet=self.document.plan.sheet
        label=next((k for k,v in self.papers.items() if v==sheet.size),None)
        if label is None:
            label=f'自定 {sheet.size[0]*25.4/72:.1f}×{sheet.size[1]*25.4/72:.1f} mm'
            self.papers[label]=sheet.size;self.paper_box.configure(values=tuple(self.papers))
        self.paper.set(label);self.margin.set(f'{sheet.margin*25.4/72:g}')
        if self.document.plan.pages:self.mode.set(next(k for k,v in MODES.items() if v==self.document.plan.pages[0].fit_mode))

    def apply_options(self):
        if self.busy:return
        try:self.document.configure(Sheet(self.papers[self.paper.get()],0,float(self.margin.get())*72/25.4),MODES[self.mode.get()]);self.changed()
        except (ValueError,KeyError) as error:messagebox.showerror('页面设置未应用',str(error),parent=self)

    def open_rules(self):
        from .rules_ui import open_rules
        if not self.busy:open_rules(self)

    def job(self,title,function,done):
        if self.busy:return
        self.busy=True;self.cancel=threading.Event();self.status.set(title)
        for control in self.controls:control.state(['disabled'])
        self.cancel_button.state(['!disabled']);self.selection_changed()
        def worker():
            try:self.events.put(('done',function(),done))
            except Exception as error:self.events.put(('error',str(error),None))
        self.worker=threading.Thread(target=worker,daemon=True);self.worker.start()

    def poll(self):
        if self.closed:return
        latest=None
        for _ in range(40):
            try:kind,value,callback=self.events.get_nowait()
            except queue.Empty:break
            if kind=='progress':latest=value;continue
            self.busy=False
            for control in self.controls:control.state(['!disabled'])
            self.cancel_button.state(['disabled'])
            if kind=='error':
                self.status.set(value)
                if not self.closing:messagebox.showerror('操作未完成',value,parent=self)
            elif not self.cancel.is_set():callback(value)
            else:self.status.set('已停止。')
            if self.closed:return
            self.selection_changed();latest=None
            if self.closing:self.shutdown();return
        if latest is not None:self.status.set(f'正在处理 {latest[0]} / {latest[1]} 页')
        self.poll_timer=self.after(30,self.poll)

    def progress(self,a,b):self.events.put(('progress',(a,b),None))

    def add_files(self):
        paths=filedialog.askopenfilenames(parent=self,title='添加 PDF',filetypes=[('PDF 文件','*.pdf')])
        if paths:self.import_paths(paths)

    def import_paths(self,paths):
        def done(sources):self.document.import_sources(sources);self.changed();self.status.set(f'已添加 {len(sources)} 个文件。')
        self.job('正在读取 PDF…',lambda:inspect_sources(paths),done)

    def save_plan(self,after=None):
        name=filedialog.asksaveasfilename(parent=self,title='保存整理方案',defaultextension='.ruyiplan.json',initialfile='页面整理.ruyiplan.json',filetypes=[('如意页面方案','*.ruyiplan.json')])
        if not name:self.after_close=None;return
        snapshot=PlanDocument(self.document.plan,self.document.sources)
        def done(_):
            self.plan_path=name;self.dirty=False;self.status.set('方案已保存。')
            if after:after()
        self.job('正在保存方案…',lambda:snapshot.save(name),done)

    def open_plan(self):
        name=filedialog.askopenfilename(parent=self,title='打开整理方案',filetypes=[('如意页面方案','*.ruyiplan.json')])
        if not name:return
        def done(doc):
            self.document.commit(doc.plan,(),doc.sources);self.plan_path=name;self.dirty=False
            self.sync_options()
            self.grid_view.invalidate();self.status.set('方案已打开。')
            self.preview.changed()
        self.job('正在校验方案中的源文件…',lambda:PlanDocument.load(name),done)

    def export_all(self):self.export_indices(None)
    def extract(self):self.export_indices(self.document.selection)
    def export_indices(self,indices):
        if not self.document.plan.pages:return
        if indices is not None and not indices:return
        name=filedialog.asksaveasfilename(parent=self,title='导出 PDF',defaultextension='.pdf',initialfile='如意_页面整理.pdf',filetypes=[('PDF 文件','*.pdf')])
        if name:self.export_to(name,indices)

    def export_to(self,path,indices=None):
        plan=self.document.plan
        if indices is not None:plan=replace(plan,pages=tuple(plan.pages[i] for i in sorted(indices)))
        registry=self.document.sources.copy()
        from .rule_spec import page_count
        def done(output):self.last_output=str(output);self.status.set(f'已导出 {page_count(plan)} 页：{Path(output).name}')
        self.job('正在导出 PDF…',lambda:export_pdf(plan,registry,path,self.progress,self.cancel),done)

    def split(self):
        if not self.document.plan.pages:return
        try:count=int(self.split_count.get());assert count>0
        except (ValueError,AssertionError):messagebox.showerror('每份页数无效','请输入大于零的整数。',parent=self);return
        parent=filedialog.askdirectory(parent=self,title='选择拆分结果保存位置')
        if not parent:return
        directory=Path(parent)/('如意拆分_'+datetime.now().strftime('%Y%m%d_%H%M%S_%f'))
        plan=self.document.plan;registry=self.document.sources.copy()
        self.job('正在拆分…',lambda:split_pdf(plan,registry,directory,count,self.progress,self.cancel),lambda outputs:self.status.set(f'已保存 {len(outputs)} 份 PDF：{directory.name}'))

    def on_close(self):
        if self.busy:
            if messagebox.askyesno('停止并关闭','当前操作尚未结束，停止后关闭？',parent=self):self.closing=True;self.cancel.set();self.status.set('正在停止…')
            else:self.after_close=None
            return
        if self.dirty and self.document.plan.pages:
            answer=messagebox.askyesnocancel('保存方案','关闭前保存当前整理方案？',parent=self)
            if answer is None:self.after_close=None;return
            if answer:self.save_plan(after=self.shutdown);return
        self.shutdown()

    def shutdown(self):
        if self.closed:return
        self.closed=True;self.grid_view.shutdown();self.preview.shutdown()
        if self.poll_timer is not None:self.after_cancel(self.poll_timer)
        root=self.winfo_toplevel()
        for key,command in self.shortcut_bindings.items():
            root.unbind(key,command)
        callback=self.after_close
        self.destroy()
        if callback:self.master.after_idle(callback)
