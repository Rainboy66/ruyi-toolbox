"""Native desktop shell for the approved B08 panel design."""
from pathlib import Path
import math
from functools import lru_cache
import os
import sys
from time import perf_counter
import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageDraw, ImageFont, ImageTk, ImageOps
from desktop_integration import load_settings, save_settings, settings_path, system_dark, create_desktop_shortcut, window_scale
from id_photo_backgrounds import BACKGROUND_NAMES, resolve_background
from window_paint import RESIZE_ENDED, RESIZE_SETTLE_MS, track_window_resize, window_is_resizing

VERSION = '1.1.0'
HOME_RESIZE_THRESHOLD = 8
ASSETS = Path(getattr(sys, '_MEIPASS', Path(__file__).resolve().parent)) / 'assets'
PALETTES = {
    '雾蓝': ('#FAFBFE', '#374767', '#EDF1F8'), '石墨': ('#FAFAFB', '#424A57', '#EFF0F3'),
    '淡紫': ('#FBFAFE', '#695B89', '#F1EDF8'), '薄荷': ('#F8FCFA', '#3E6A5B', '#EAF3EE'),
    '浅粉': ('#FEFAFC', '#845F71', '#F8EDF2'), '暮霞': ('#FFFBF8', '#87684C', '#F8EEE4'),
}


@lru_cache(maxsize=192)
def font(size, bold=False):
    directory = Path(os.environ.get('SystemRoot', r'C:\Windows')) / 'Fonts'
    return ImageFont.truetype(str(directory / ('msyhbd.ttc' if bold else 'msyh.ttc')), max(1, round(size)))


class ScrollArea(ttk.Frame):
    """A scrollable page body; lists keep their own independent scroll position."""
    def __init__(self, master, padding=0):
        super().__init__(master)
        self.canvas = tk.Canvas(self, highlightthickness=0, borderwidth=0, background='#FAFBFE')
        self.scrollbar = ttk.Scrollbar(self, orient='vertical', command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self._position)
        self.canvas.pack(side='left', fill='both', expand=True)
        self.scrollbar.pack(side='right', fill='y')
        self.body = ttk.Frame(self.canvas, padding=padding)
        self.window = self.canvas.create_window(0, 0, window=self.body, anchor='nw')
        self._viewport = None
        self._scroll_region = None
        self._refresh_pending = None
        self.canvas.bind('<Configure>', self._resize)
        self.canvas.bind('<Map>', self._schedule_region)
        self.canvas.bind('<Destroy>', self._cancel_refresh)
        self.bind_all('<MouseWheel>', self._wheel, add='+')

    def _resize(self, event=None):
        if window_is_resizing(self):
            return
        size = ((event.width, event.height) if event is not None
                else (self.canvas.winfo_width(), self.canvas.winfo_height()))
        if size == self._viewport:
            return
        if self._viewport is None or size[0] != self._viewport[0]:
            self.canvas.itemconfigure(self.window, width=size[0])
        self._viewport = size
        self._schedule_region()

    def _schedule_region(self, _event=None):
        if self._refresh_pending is None:
            self._refresh_pending = self.after_idle(self._update_scroll_region)

    def refresh(self):
        if not self.canvas.winfo_ismapped():
            return
        if self._refresh_pending is not None:
            self.after_cancel(self._refresh_pending)
        self._refresh_pending = self.after(16, self._update_scroll_region)

    def _update_scroll_region(self):
        self._refresh_pending = None
        region = (0, 0, self.canvas.winfo_width(),
                  max(self.body.winfo_reqheight(), self.canvas.winfo_height()))
        if region != self._scroll_region:
            self._scroll_region = region
            self.canvas.configure(scrollregion=region)

    def _cancel_refresh(self, _event):
        if self._refresh_pending is not None:
            self.after_cancel(self._refresh_pending)
            self._refresh_pending = None

    def _position(self, first, last):
        self.scrollbar.set(first, last)

    def _wheel(self, event):
        widget = event.widget
        # Never consume another page's event, and leave Treeview's native wheel intact.
        if isinstance(widget, (ttk.Treeview, ttk.Combobox, ttk.Spinbox)):
            return
        while widget is not None and widget != self:
            widget = getattr(widget, 'master', None)
        if widget == self and self.canvas.winfo_ismapped():
            self.canvas.yview_scroll(-int(event.delta / 120), 'units')
            return 'break'


class PageHost(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master)
        self.app = app
        self.current = None
        self.pages = []
        self._size = None
        self._drag_canvas = None
        self._drag_photo = None
        self._drag_focus = None
        self._owner_window = self.winfo_toplevel()
        self._owner_window._ruyi_snapshot_hwnd = 0
        self.bind('<Configure>', self.resize_page)

    def add(self, pane, text=''):
        self.pages.append(pane)

    def select(self, pane=None):
        if pane is None:
            return str(self.current) if self.current else ''
        if isinstance(pane, int):
            pane = self.pages[pane]
        if self.current == pane:
            return
        if self.current:
            if hasattr(self.current, 'on_hide'):
                self.current.on_hide()
            self.current.place_forget()
        self.current = pane
        self._size = None
        self._owner_window._ruyi_snapshot_hwnd = (0 if pane is self.app.home else self.winfo_id())
        self.resize_page()

    def resize_page(self, event=None):
        if self.current is None:
            return
        if window_is_resizing(self):
            self._show_drag_snapshot()
            return
        size = ((event.width, event.height) if event is not None
                else (self.winfo_width(), self.winfo_height()))
        if size != self._size or self._drag_photo is not None:
            # Keep the current pane's native controls at their last size
            # throughout the live resize. The host clips them as needed.
            self.current.place(x=0, y=0, width=size[0], height=size[1])
            self._size = size
        if self._drag_photo is not None:
            self._drag_canvas.place_forget()
            self._drag_canvas.itemconfigure(self._drag_item, image='')
            self._drag_photo = None
            if self._drag_focus is not None and self._drag_focus.winfo_exists():
                self._drag_focus.focus_set()
            self._drag_focus = None
        image = getattr(self._owner_window, '_ruyi_resize_snapshot', None)
        if image is not None:
            self._owner_window._ruyi_resize_snapshot = None
            image.close()

    def _show_drag_snapshot(self):
        image = getattr(self._owner_window, '_ruyi_resize_snapshot', None)
        if image is None or self._drag_photo is not None:
            return
        self._owner_window._ruyi_resize_snapshot = None
        if self._drag_canvas is None:
            self._drag_canvas = tk.Canvas(self, highlightthickness=0, borderwidth=0)
            self._drag_item = self._drag_canvas.create_image(0, 0, anchor='nw')
        width,height=image.size
        with image:
            self._drag_photo = ImageTk.PhotoImage(image, master=self._drag_canvas)
        self._drag_canvas.configure(background=self.app.tokens['bg'])
        self._drag_canvas.itemconfigure(self._drag_item, image=self._drag_photo)
        self._drag_canvas.place(x=0, y=0, width=width, height=height)
        self._drag_focus = self.focus_get()
        self.current.place_forget()


class HomeCanvas(tk.Canvas):
    MODES = [('office', 'Office 转 PDF'), ('pdf', 'PDF 转 Office'), ('image', '图片压缩'), ('idphoto', '证件照换底'), ('organizer', 'PDF 合并拼凑')]
    def __init__(self, master, app):
        super().__init__(master, highlightthickness=0, borderwidth=0, takefocus=1)
        self.app = app
        self.hover = None
        self.active_mode = None
        self.hitboxes = {}
        self._resize_timer = None
        self._observed_size = None
        self._released = False
        self._hover_timer = None
        self._desired_hover = None
        self._cache_signature = None
        self._card_tags = {}
        self._tile_cache = {}
        self.logo = Image.open(ASSETS / 'ruyi-icon-256.png').convert('RGBA')
        self.icons = {key: Image.open(ASSETS / (key + '.png')).convert('RGBA') for key in ('files', 'arrow-right', 'arrow-left', 'image-down', 'contact-round', 'settings-2')}
        self.bind('<Configure>', self.schedule)
        self.bind('<Map>', self.on_map)
        self.bind('<Tab>', self.keyboard_next)
        self.bind('<Shift-Tab>', lambda e:self.keyboard_next(e,-1))
        self.bind('<Return>', self.keyboard_activate)
        self.bind('<space>', self.keyboard_activate)
        self.bind('<Destroy>', self.cleanup, add='+')
        self.photo = None

    def cleanup(self, event):
        if event.widget == self:
            for timer in (self._resize_timer, self._hover_timer):
                if timer is not None:
                    self.after_cancel(timer)
            self.logo.close()
            for image in self.icons.values():
                image.close()

    def schedule(self, event=None):
        size = self.winfo_width(), self.winfo_height()
        if event is not None and size == self._observed_size and self._cache_signature is not None:
            return
        self._observed_size = size
        if self._resize_timer is not None:
            self.after_cancel(self._resize_timer)
            self._resize_timer = None
        if window_is_resizing(self):
            self._released = False
            return
        if self._cache_signature is None:
            self._resize_timer = self.after_idle(self.refresh)
        elif self._released:
            self._resize_timer = self.after(16, self.refresh)
        elif event is None or self._size_changed(size):
            self._resize_timer = self.after(RESIZE_SETTLE_MS, self.refresh)

    def finish_live_resize(self):
        # A confirmed mouse release does not need a second quiet period.
        self._released = True
        self.schedule()

    def on_map(self, _event):
        # The window may have changed size while another page was open.
        self._released = True
        self.schedule()

    def _size_changed(self, size):
        return max(abs(new - old) for new, old in zip(size, self._cache_signature[:2])) >= HOME_RESIZE_THRESHOLD

    def request_hover(self, key):
        self._desired_hover = key
        if self._hover_timer is None:
            self._hover_timer = self.after(16, self.apply_hover)

    def apply_hover(self):
        if self._hover_timer is not None:
            self.after_cancel(self._hover_timer)
        self._hover_timer = None
        self.hover = self._desired_hover or self.active_mode
        if hasattr(self, '_surface'):
            self._surface.select(self.hover)

    def enter_card(self, key):
        if self.cget('cursor') != 'hand2':
            self.configure(cursor='hand2')
        self.request_hover(key)

    def leave_card(self, key):
        if self._desired_hover == key:
            self.configure(cursor='')
            self.request_hover(None)

    def activate(self, key):
        if key in dict(self.MODES):
            self.active_mode=key
            self.request_hover(None)
        if key == 'settings':
            self.app.show_settings()
        else:
            self.app.show_mode(key)

    def keyboard_next(self, _event, step=1):
        keys=[key for key,_name in self.MODES]+['settings']
        index=keys.index(self._desired_hover) if self._desired_hover in keys else (-1 if step==1 else 0)
        self.request_hover(keys[(index+step)%len(keys)])
        return 'break'

    def keyboard_activate(self, _event):
        key=self._desired_hover or self.active_mode
        if key:self.activate(key)
        return 'break'

    def _render(self, hover):
        width, height = self.winfo_width(), self.winfo_height()
        if width < 10 or height < 10:
            return
        t = self.app.tokens
        scale = min(width / 1024, height / 760)
        # Full-window drawing uses pixel fonts and antialiasing, so it remains crisp at DPI changes.
        ratio = 2
        canvas = Image.new('RGB', (width * ratio, height * ratio), t['bg'])
        d = ImageDraw.Draw(canvas)
        if self.app.settings.get('gradient'):
            from PIL import ImageColor
            start=ImageColor.getrgb(t['bg']);end=ImageColor.getrgb(t['hover'])
            for row in range(height*ratio):
                amount=(row/max(1,height*ratio-1))*.65
                color=tuple(round(a+(b-a)*amount) for a,b in zip(start,end))
                d.line((0,row,width*ratio,row),fill=color)
        def box(b): return tuple(round(x * ratio) for x in b)
        def roundrect(b, radius, fill, outline=None, stroke=1):
            d.rounded_rectangle(box(b), radius=round(radius * ratio), fill=fill, outline=outline, width=max(1, round(stroke * ratio)))
        def text(x, y, value, size, color, bold=False, anchor='mm'):
            d.text((round(x * ratio), round(y * ratio)), value, font=font(size * ratio, bold), fill=color, anchor=anchor)
        def paste_icon(key, x, y, w, h):
            source = self.icons[key]
            alpha = source.getchannel('A').resize((max(1, round(w * ratio)), max(1, round(h * ratio))), Image.Resampling.LANCZOS)
            image = Image.new('RGB', alpha.size, t['accent'])
            canvas.paste(image, (round(x * ratio), round(y * ratio)), alpha)
            image.close(); alpha.close()
        # Preserve the 1.0 homepage's edge curves, logo and card proportions.
        line = t['pattern']
        for index in range(9):
            offset = index * 23 * scale
            points = []
            for step in range(101):
                u = step / 100
                p0 = (-28 * scale, (278 * scale - offset))
                p1 = (-9 * scale, (114 * scale - offset * .4))
                p2 = ((102 - index) * scale, (6 - index * 3) * scale)
                p3 = ((318 - index * 26) * scale, -28 * scale)
                points.append(((1-u)**3*p0[0]+3*(1-u)**2*u*p1[0]+3*(1-u)*u*u*p2[0]+u**3*p3[0],
                               (1-u)**3*p0[1]+3*(1-u)**2*u*p1[1]+3*(1-u)*u*u*p2[1]+u**3*p3[1]))
            d.line([(round(x*ratio), round(y*ratio)) for x,y in points], fill=line, width=1)
            d.line([(round((width-x)*ratio), round((height-y)*ratio)) for x,y in points], fill=line, width=1)
        brand_y = height * .148
        logo_size = 74 * scale
        gap = 20 * scale
        word_font = font(34 * scale * ratio, True)
        word_width = d.textlength('如意', font=word_font) / ratio
        x = (width - logo_size - gap - word_width) / 2
        # Reuse the approved clean-alpha B08 icon; a second crop/mask leaves
        # pale horizontal seams on dark themes.
        cropped=self.logo.resize((round(logo_size*ratio),)*2,Image.Resampling.LANCZOS)
        canvas.paste(cropped,(round(x*ratio),round(brand_y*ratio)),cropped);cropped.close()
        text(x+logo_size+gap,brand_y+logo_size/2,'如意',34*scale,t['ink'],True,'lm')
        panel_width = min(width - 76 * scale, 916 * scale)
        panel_x = (width - panel_width)/2
        # Only reduce the logo-to-panel gap to accommodate the additional row.
        panel_top = height * .275
        padding, card_gap = 22*scale,16*scale
        card_height = 220*scale
        panel_height = 2*card_height+2*padding+card_gap
        roundrect((panel_x,panel_top,panel_x+panel_width,panel_top+panel_height),28*scale,t['surface'],t['line'])
        card_width = (panel_width-2*padding-3*card_gap)/4
        from PIL import ImageColor
        normal_rgb=ImageColor.getrgb(t['soft']);active_fill=t['hover']
        if sum(ImageColor.getrgb(active_fill))>=sum(normal_rgb):
            active_fill='#%02x%02x%02x'%tuple(round(channel*.72) for channel in normal_rgb)
        self.hitboxes = {}
        for index,(key,name) in enumerate(self.MODES):
            row,col=divmod(index,4)
            left=panel_x+padding+col*(card_width+card_gap)
            top=panel_top+padding+row*(card_height+card_gap)
            fill=active_fill if hover==key else t['soft']
            roundrect((left,top,left+card_width,top+card_height),15*scale,fill)
            center=left+card_width/2
            icon_y=top+card_height/2-37*scale
            if key in ('office','pdf'):
                start=center-46*scale
                paste_icon('files',start,icon_y,25*scale,29*scale)
                paste_icon('arrow-right' if key=='office' else 'arrow-left',start+33*scale,icon_y+2*scale,25*scale,25*scale)
                text(start+66*scale,icon_y+14*scale,'PDF',15*scale,t['accent'],anchor='lm')
            elif key=='organizer':
                for row in range(2):
                    for col in range(2):
                        left_icon=center-16*scale+col*19*scale;top_icon=icon_y-2*scale+row*19*scale
                        d.rounded_rectangle(box((left_icon,top_icon,left_icon+13*scale,top_icon+13*scale)),radius=max(1,round(2*scale*ratio)),outline=t['accent'],width=max(1,round(1.8*scale*ratio)))
            else:
                paste_icon('image-down' if key=='image' else 'contact-round',center-16*scale,icon_y-2*scale,32*scale,36*scale)
            text(center,top+card_height/2+32*scale,name,15*scale,t['ink'])
            self.hitboxes[key]=(left,top,left+card_width,top+card_height)
        settings_box=(width-66*scale,18*scale,width-24*scale,60*scale)
        if hover=='settings':
            roundrect(settings_box,9*scale,t['hover'])
        paste_icon('settings-2',width-54*scale,30*scale,20*scale,20*scale)
        self.hitboxes['settings']=settings_box
        text(width/2,height-29*scale,VERSION,12*scale,t['muted'])
        rendered=canvas.resize((width,height),Image.Resampling.LANCZOS)
        canvas.close()
        return rendered

    def refresh(self):
        if self._resize_timer is not None:
            self.after_cancel(self._resize_timer)
            self._resize_timer = None
        self._released = False
        if window_is_resizing(self):
            return
        width, height = self.winfo_width(), self.winfo_height()
        if min(width, height) < 10 or not self.winfo_ismapped():
            self._observed_size = None
            return
        tokens = self.app.tokens
        style = (tuple(tokens.items()), bool(self.app.settings.get('gradient')))
        previous = self._cache_signature
        if previous is not None and previous[2] == style and not self._size_changed((width, height)):
            return
        if not hasattr(self, "_surface"):
            from home_surface import HomeSurface
            self._surface = HomeSurface(self)
        start = perf_counter()
        artwork = self._surface.prepare(width, height, tokens['accent'])
        prepare_ms = (perf_counter() - start) * 1000
        self.draw((width, height, style), artwork, prepare_ms)

    def draw(self, signature, artwork, prepare_ms):
        start = perf_counter()
        self._surface.paint(signature, artwork)
        draw_ms = (perf_counter() - start) * 1000
        if draw_ms > 30 or draw_ms + prepare_ms > 30:
            width, height = signature[:2]
            log = settings_path().with_name('resize.log')
            try:
                log.parent.mkdir(parents=True, exist_ok=True)
                with log.open('a', encoding='utf-8') as stream:
                    stream.write(f'home {width}x{height} draw_ms={draw_ms:.2f} prepare_ms={prepare_ms:.2f}\n')
            except OSError:
                pass


class TaskPaneView(ttk.Frame):
    def __init__(self, app, mode):
        super().__init__(app.tabs, padding=(28,14,28,18))
        self.app,self.mode=app,mode
        self.files,self.outputs,self.photos={},{},[]
        self.details={};self.controls=[];self.locked=False;self.range_anchor=None
        self.columnconfigure(0,weight=1);self.rowconfigure(1,weight=1)
        header=ttk.Frame(self);header.grid(row=0,column=0,sticky='ew',pady=(0,18))
        ttk.Button(header,text='‹  首页',command=lambda:getattr(app,'show_home',lambda:None)(),style='Quiet.TButton').pack(side='left')
        title={'image':'图片压缩','office':'Office 转 PDF','pdf':'PDF 转 Office','idphoto':'证件照换底'}[mode]
        ttk.Label(header,text=title,style='Title.TLabel').pack(side='left',padx=18)
        ttk.Button(header,text='设置',command=lambda:getattr(app,'show_settings',lambda:None)(),style='Quiet.TButton').pack(side='right')
        self.scroll=ScrollArea(self);self.scroll.grid(row=1,column=0,sticky='nsew')
        content=self.scroll.body;content.columnconfigure(0,weight=1);content.rowconfigure(2,weight=1)
        options=ttk.Frame(content,style='Card.TFrame',padding=20);options.grid(row=0,column=0,sticky='ew',pady=(0,18))
        for c in range(3):options.columnconfigure(c,weight=1,uniform='option')
        def field(label,column,row=0):
            f=ttk.Frame(options,style='Card.TFrame');f.grid(row=row,column=column,sticky='ew',padx=(0,18) if column<2 else 0,pady=(0,12) if row==0 else 0)
            ttk.Label(f,text=label,style='Card.TLabel').pack(anchor='w',pady=(0,8))
            return f
        def combo(parent,var,values,width=18):
            w=ttk.Combobox(parent,textvariable=var,values=values,state='readonly',width=width)
            w.pack(fill='x');self.controls.append(w);return w
        def entry(parent,var,width=12):
            w=ttk.Entry(parent,textvariable=var,width=width);w.pack(fill='x');self.controls.append(w);return w
        if mode=='image':
            self.format_var=tk.StringVar(value='原格式');self.quality_var=tk.StringVar(value='82');self.max_var=tk.StringVar();self.target_var=tk.StringVar();self.lossless_var=tk.BooleanVar(value=False)
            self.format_box=combo(field('输出格式',0),self.format_var,['原格式','JPEG','PNG','WEBP']);self.format_box.bind('<<ComboboxSelected>>',self.format_changed)
            f=field('画质（1–100%）',1);self.quality_scale=ttk.Spinbox(f,textvariable=self.quality_var,from_=1,to=100,width=8);self.quality_scale.pack(fill='x');self.controls.append(self.quality_scale)
            entry(field('最长边（像素，可留空）',2),self.max_var)
            f=field('编码方式',0,1);w=ttk.Checkbutton(f,text='无损编码',variable=self.lossless_var,command=self.format_changed,style='Card.Ruyi.TCheckbutton');w.pack(anchor='w');self.controls.append(w)
            entry(field('目标大小（KB，可留空）',1,1),self.target_var)
            f=field('参数说明',2,1);self.hint=ttk.Label(f,text='画质是编码参数；100%不代表JPEG无损。',style='CardMuted.TLabel',wraplength=220);self.hint.pack(anchor='w')
        elif mode=='office':
            self.office_layout_var=tk.StringVar(value='自动匹配原编辑器')
            self.office_layout_box=combo(field('排版基准',0),self.office_layout_var,['自动匹配原编辑器','WPS','Microsoft Office','LibreOffice 兼容输出'])
            f=field('转换引擎',1);w=ttk.Button(f,text='选择 LibreOffice',command=self.select_engine);w.pack(anchor='w');self.controls.append(w)
            self.engine_label=ttk.Label(options,text='',style='CardMuted.TLabel',wraplength=800);self.engine_label.grid(row=2,column=0,columnspan=3,sticky='w');self.refresh_engine()
        elif mode=='pdf':
            self.target_var=tk.StringVar(value='Word');self.layout_var=tk.StringVar(value='自动还原')
            self.target_box=combo(field('输出格式',0),self.target_var,['Word','Excel','PPT']);self.target_box.bind('<<ComboboxSelected>>',self.pdf_target_changed)
            f=ttk.Frame(options,style='Card.TFrame');f.grid(row=0,column=1,columnspan=2,sticky='w',padx=8)
            self.layout_label=ttk.Label(f,text='版式',style='Card.TLabel')
            self.layout_box=ttk.Combobox(f,textvariable=self.layout_var,values=['自动还原'],state='readonly',width=22);self.controls.append(self.layout_box)
            self.pdf_hint=ttk.Label(options,text='Word 自动还原 · 正文和支持的表格可编辑，复杂公式局部保留为图片。',style='CardMuted.TLabel',wraplength=780);self.pdf_hint.grid(row=2,column=0,columnspan=3,sticky='w')
        else:
            self.custom_color=resolve_background('标准蓝').start;self.color_var=tk.StringVar(value='标准蓝');self.edge_var=tk.StringVar(value='自动精细');self.advanced_edge_var=tk.BooleanVar(value=False)
            self.color_box=combo(field('新背景',0),self.color_var,BACKGROUND_NAMES+('自定义',));self.color_box.bind('<<ComboboxSelected>>',self.color_changed)
            f=field('自定义颜色',1);self.color_button=ttk.Button(f,text='选择颜色…',command=self.choose_color,state='disabled');self.color_button.pack(anchor='w');self.controls.append(self.color_button)
            f=field('抠图',2);self.advanced_edge_button=ttk.Checkbutton(f,text='高级边缘设置',variable=self.advanced_edge_var,command=self.edge_options_changed,style='Card.Ruyi.TCheckbutton');self.advanced_edge_button.pack(anchor='w');self.controls.append(self.advanced_edge_button)
            self.edge_options=ttk.Frame(f,style='Card.TFrame');self.edge_box=ttk.Combobox(self.edge_options,textvariable=self.edge_var,values=['自动精细','自然发丝','去除余边'],state='readonly',width=12);self.edge_box.pack(side='left');self.controls.append(self.edge_box)
            info=ttk.Frame(options,style='Card.TFrame');info.grid(row=2,column=0,columnspan=3,sticky='ew',pady=(8,0));info.columnconfigure(1,weight=1)
            self.color_swatch=tk.Label(info,borderwidth=0);self.color_swatch.grid(row=0,column=0,rowspan=2,padx=(0,14),sticky='w')
            self.color_label=ttk.Label(info,style='CardMuted.TLabel',wraplength=650);self.color_label.grid(row=0,column=1,sticky='w')
            self.color_usage=ttk.Label(info,style='CardMuted.TLabel',wraplength=700);self.color_usage.grid(row=1,column=1,sticky='w');self.color_changed()
        toolbar=ttk.Frame(content);toolbar.grid(row=1,column=0,sticky='ew',pady=(0,10))
        add_text={'image':'添加图片','office':'添加办公文件','pdf':'添加 PDF','idphoto':'添加照片'}[mode]
        self.add_button=ttk.Button(toolbar,text=add_text,command=self.add_files);self.add_button.pack(side='left');self.controls.append(self.add_button)
        for label,command in [('清空',self.clear),('全选',self.select_all),('取消选择',self.clear_selection)]:
            w=ttk.Button(toolbar,text=label,command=command,style='Quiet.TButton');w.pack(side='left',padx=(8,0));self.controls.append(w)
        self.selection_label=ttk.Label(toolbar,style='Muted.TLabel');self.selection_label.pack(side='right',padx=6)
        table=ttk.Frame(content);table.grid(row=2,column=0,sticky='nsew');table.columnconfigure(0,weight=1);table.rowconfigure(0,weight=1)
        self.tree=ttk.Treeview(table,columns=('file','status','result'),show='tree headings',height=8,selectmode='extended')
        self.tree.heading('#0',text='选择');self.tree.column('#0',width=50,minwidth=50,stretch=False,anchor='center')
        for key,label,width in [('file','文件',260),('status','状态',80),('result','结果 / 说明',420)]:
            self.tree.heading(key,text=label);self.tree.column(key,width=width,minwidth=65,stretch=key!='status')
        self.tree.grid(row=0,column=0,sticky='nsew');self.file_scrollbar=ttk.Scrollbar(table,orient='vertical',command=self.tree.yview);self.file_scrollbar.grid(row=0,column=1,sticky='ns');self.tree.configure(yscrollcommand=self.file_scrollbar.set)
        hs=ttk.Scrollbar(table,orient='horizontal',command=self.tree.xview);hs.grid(row=1,column=0,sticky='ew');self.tree.configure(xscrollcommand=hs.set)
        for event,callback in [('<<TreeviewSelect>>',self.selection_changed),('<Double-1>',self.show_details),('<Control-a>',self.select_all),('<Control-A>',self.select_all),('<Escape>',self.clear_selection),('<Control-space>',self.toggle_focused)]:self.tree.bind(event,callback)
        for event,callback in [('<Control-Up>',lambda _e:self.move_selection(-1,preserve=True)),('<Control-Down>',lambda _e:self.move_selection(1,preserve=True)),('<Shift-Up>',lambda _e:self.move_selection(-1,extend=True)),('<Shift-Down>',lambda _e:self.move_selection(1,extend=True))]:self.tree.bind(event,callback)
        self.tree.bind('<Button-1>',self.selection_anchor_changed,add='+')
        if mode in ('image','idphoto'):
            previews=ttk.Frame(content);previews.grid(row=3,column=0,sticky='ew',pady=(18,0));previews.columnconfigure((0,1),weight=1,uniform='preview')
            for column,title,attr in [(0,'原图预览','before'),(1,'结果预览','after')]:
                f=ttk.Frame(previews,style='Card.TFrame',padding=14,height=160);f.grid(row=0,column=column,sticky='nsew',padx=(0,10) if column==0 else (10,0));f.grid_propagate(False);f.columnconfigure(0,weight=1);f.rowconfigure(0,weight=1)
                label=ttk.Label(f,text=title,anchor='center',style='Card.TLabel');label.grid(row=0,column=0,sticky='nsew');setattr(self,attr,label)
        footer=ttk.Frame(self);footer.grid(row=2,column=0,sticky='ew',pady=(16,0));footer.columnconfigure(1,weight=1)
        ttk.Label(footer,text='保存位置').grid(row=0,column=0,padx=(0,12));self.output_var=tk.StringVar();w=ttk.Entry(footer,textvariable=self.output_var);w.grid(row=0,column=1,sticky='ew');self.controls.append(w)
        w=ttk.Button(footer,text='浏览…',command=self.browse_output);w.grid(row=0,column=2,padx=(10,0));self.controls.append(w)
        actions=ttk.Frame(footer);actions.grid(row=1,column=0,columnspan=3,sticky='ew',pady=(16,0))
        self.start_button=ttk.Button(actions,text='开始处理',command=lambda:app.start(self),style='Accent.TButton');self.start_button.pack(side='left')
        self.all_button=ttk.Button(actions,text='处理全部',command=lambda:app.start(self,process_all=True));self.controls.extend([self.start_button,self.all_button])
        ttk.Button(actions,text='打开输出文件夹',command=self.open_output,style='Quiet.TButton').pack(side='right')
        self.selection_changed()


class SettingsPane(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master,padding=(28,14,28,24));self.app=app
        self.columnconfigure(0,weight=1);self.rowconfigure(1,weight=1)
        head=ttk.Frame(self);head.grid(row=0,column=0,sticky='ew',pady=(0,20));ttk.Button(head,text='‹  返回',command=lambda:app.go_back(),style='Quiet.TButton').pack(side='left');ttk.Label(head,text='设置',style='Title.TLabel').pack(side='left',padx=18)
        self.scroll=ScrollArea(self);self.scroll.grid(row=1,column=0,sticky='nsew');body=self.scroll.body;body.columnconfigure(0,weight=1)
        self.appearance=tk.StringVar(value=app.settings['appearance']);self.palette=tk.StringVar(value=app.settings['palette']);self.gradient=tk.BooleanVar(value=app.settings['gradient'])
        card=ttk.Frame(body,style='Card.TFrame',padding=24);card.grid(row=0,column=0,sticky='ew',pady=(0,18));card.columnconfigure(1,weight=1)
        ttk.Label(card,text='外观',style='CardHeading.TLabel').grid(row=0,column=0,columnspan=2,sticky='w',pady=(0,20))
        ttk.Label(card,text='主题',style='Card.TLabel').grid(row=1,column=0,sticky='w',padx=(0,30));row=ttk.Frame(card,style='Card.TFrame');row.grid(row=1,column=1,sticky='w')
        self.choice_buttons=[]
        for name in ('浅色','深色','跟随系统'):
            button=ttk.Button(row,text=name,command=lambda n=name:self.choose(self.appearance,n),width=10)
            button.pack(side='left',padx=(0,12));self.choice_buttons.append((button,self.appearance,name))
        ttk.Label(card,text='配色',style='Card.TLabel').grid(row=2,column=0,sticky='w',pady=(20,0));row=ttk.Frame(card,style='Card.TFrame');row.grid(row=2,column=1,sticky='w',pady=(20,0))
        for index,name in enumerate(PALETTES):
            button=ttk.Button(row,text=name,command=lambda n=name:self.choose(self.palette,n),width=10)
            button.grid(row=index//3,column=index%3,sticky='ew',padx=(0,12),pady=5);self.choice_buttons.append((button,self.palette,name))
        ttk.Checkbutton(card,text='首页渐变底色',variable=self.gradient,command=self.changed,style='Card.Ruyi.TCheckbutton').grid(row=3,column=1,sticky='w',pady=(20,0))
        card=ttk.Frame(body,style='Card.TFrame',padding=24);card.grid(row=1,column=0,sticky='ew',pady=(0,18));card.columnconfigure(0,weight=1)
        ttk.Label(card,text='桌面快捷方式',style='CardHeading.TLabel').grid(row=0,column=0,sticky='w');self.shortcut_button=ttk.Button(card,text='一键创建桌面快捷方式',command=self.shortcut);self.shortcut_button.grid(row=0,column=1,padx=12)
        self.notice=ttk.Label(card,text='',style='CardMuted.TLabel',wraplength=700);self.notice.grid(row=1,column=0,columnspan=2,sticky='w',pady=(12,0))
        card=ttk.Frame(body,style='Card.TFrame',padding=24);card.grid(row=2,column=0,sticky='ew');card.columnconfigure(0,weight=1)
        ttk.Label(card,text='转换引擎',style='CardHeading.TLabel').grid(row=0,column=0,sticky='w');ttk.Button(card,text='选择 LibreOffice',command=self.choose_engine).grid(row=0,column=1,padx=12)
        self.engine_status=ttk.Label(card,text='',style='CardMuted.TLabel',wraplength=780);self.engine_status.grid(row=1,column=0,columnspan=2,sticky='w',pady=(18,0))
        ttk.Label(body,text=f'如意  {VERSION} · 正式版',style='Muted.TLabel').grid(row=3,column=0,pady=26)
        self.refresh_selections()

    def choose(self, variable, value):
        variable.set(value);self.changed()

    def refresh_selections(self):
        for button,variable,name in self.choice_buttons:
            button.configure(style='Selected.TButton' if variable.get()==name else 'TButton')

    def changed(self):
        self.app.settings.update(appearance=self.appearance.get(),palette=self.palette.get(),gradient=self.gradient.get())
        self.app.apply_theme()
        self.refresh_selections()
        try:save_settings(self.app.settings)
        except OSError:messagebox.showerror('设置未保存','无法保存偏好，请检查本机应用配置目录的写入权限。',parent=self)

    def shortcut(self):
        try:
            path=create_desktop_shortcut();self.notice.configure(text=f'已创建：{path}')
        except Exception as error:messagebox.showerror('快捷方式未创建',str(error),parent=self)
        self.scroll.refresh()

    def choose_engine(self):
        if self.app.busy:
            messagebox.showinfo('正在处理','当前任务结束后可以调整转换引擎。',parent=self);return
        self.app.office_pane.select_engine();self.refresh()

    def refresh(self):
        self.app.office_pane.refresh_engine();self.engine_status.configure(text=self.app.office_pane.engine_label.cget('text'))
        self.scroll.refresh()


def apply_theme(app):
    settings=app.settings;dark=settings['appearance']=='深色' or (settings['appearance']=='跟随系统' and system_dark())
    bg,accent,tint=PALETTES[settings['palette']]
    t={'bg':'#171D28' if dark else bg,'surface':'#222B39' if dark else '#FFFFFF','soft':'#283242' if dark else '#F7F9FC',
       'hover':'#354257' if dark else tint,'line':'#354154' if dark else '#E4E8F0','ink':'#EAF0F8' if dark else '#242C3C',
       'muted':'#AFBACB' if dark else '#6B7485','accent':'#C5D7F4' if dark else accent,'pattern':'#2C384B' if dark else '#E0E7F1'}
    app.tokens=t;app._system_dark=dark
    style=app.style
    for name in ('TFrame','TLabel','TCheckbutton','TRadiobutton','Ruyi.TCheckbutton'):
        style.configure(name,background=t['bg'],foreground=t['ink'])
    style.configure('Card.TFrame',background=t['surface'])
    for name in ('Card.TLabel','CardHeading.TLabel','CardMuted.TLabel','Card.TRadiobutton','Card.Ruyi.TCheckbutton'):
        style.configure(name,background=t['surface'],foreground=t['ink'])
    style.configure('CardHeading.TLabel',font=('Microsoft YaHei UI',11,'bold'))
    style.configure('Title.TLabel',font=('Microsoft YaHei UI',17,'bold'))
    style.configure('Muted.TLabel',foreground=t['muted']);style.configure('CardMuted.TLabel',foreground=t['muted'],font=('Microsoft YaHei UI',9))
    style.configure('TButton',padding=(16,9),background=t['surface'],foreground=t['ink'],bordercolor=t['line'],lightcolor=t['surface'],darkcolor=t['surface'],relief='flat')
    style.map('TButton',background=[('active',t['hover']),('disabled',t['soft'])],foreground=[('disabled',t['muted'])])
    style.configure('Quiet.TButton',background=t['bg'],borderwidth=0)
    style.configure('Selected.TButton',background=t['accent'],foreground='#182130' if dark else '#FFFFFF',borderwidth=0)
    style.map('Selected.TButton',background=[('active',t['accent'])],foreground=[('active','#182130' if dark else '#FFFFFF')])
    style.configure('Accent.TButton',background=t['accent'],foreground='#182130' if dark else '#FFFFFF',borderwidth=0,padding=(22,11))
    style.map('Accent.TButton',background=[('active',t['accent']),('disabled',t['line'])],foreground=[('disabled',t['muted'])])
    for name in ('TEntry','TCombobox','TSpinbox'):
        style.configure(name,fieldbackground=t['surface'],background=t['surface'],foreground=t['ink'],insertcolor=t['ink'],bordercolor=t['line'],lightcolor=t['line'],darkcolor=t['line'],padding=8,arrowsize=12)
        style.map(name,fieldbackground=[('readonly',t['surface']),('disabled',t['soft'])],foreground=[('readonly',t['ink']),('disabled',t['muted'])],selectbackground=[('!disabled',t['accent'])],selectforeground=[('!disabled','#182130' if dark else '#FFFFFF')])
    style.configure('Treeview',background=t['surface'],fieldbackground=t['surface'],foreground=t['ink'],rowheight=40,bordercolor=t['line'],font=('Microsoft YaHei UI',10))
    style.configure('Treeview.Heading',background=t['soft'],foreground=t['muted'],relief='flat',padding=(10,10))
    style.map('Treeview',background=[('selected',t['hover'])],foreground=[('selected',t['ink'])])
    style.configure('TScrollbar',background=t['line'],troughcolor=t['bg'],bordercolor=t['bg'],arrowcolor=t['muted'],arrowsize=13)
    style.configure('TProgressbar',background=t['accent'],troughcolor=t['bg'],borderwidth=0,thickness=3)
    for name in ('Card.TRadiobutton','Card.Ruyi.TCheckbutton','Ruyi.TCheckbutton'):
        style.map(name,background=[('active',t['surface'] if name.startswith('Card') else t['bg'])],foreground=[('disabled',t['muted'])])
    app.root.configure(background=t['bg'])
    app.root._ruyi_background=int.from_bytes(bytes.fromhex(t['bg'][1:]), 'little')
    app.root.option_add('*TCombobox*Listbox.background',t['surface']);app.root.option_add('*TCombobox*Listbox.foreground',t['ink'])
    def recolor(widget):
        if isinstance(widget, HomeCanvas) or widget is getattr(app, 'organizer_pane', None):
            return
        if isinstance(widget, ScrollArea):widget.refresh()
        if isinstance(widget,tk.Canvas):widget.configure(background=t['bg'])
        for child in widget.winfo_children():recolor(child)
    recolor(app.root)
    if hasattr(app,'home'):app.home.refresh()
    organizer=getattr(app,'organizer_pane',None)
    if organizer is not None and not organizer.closed:organizer.apply_theme(t)


def build_shell(app, task_class):
    from PIL import ImageDraw
    root=app.root
    app.settings=load_settings()
    if app.settings.get('libreoffice'):
        from office_engine import find_libreoffice
        app.engine=find_libreoffice(app.settings['libreoffice']) or app.engine
    root.title('如意  '+VERSION)
    scale=window_scale(root)
    app.ui_scale=scale
    width=min(round(1024*scale),root.winfo_screenwidth()-round(80*scale))
    height=min(round(760*scale),root.winfo_screenheight()-round(100*scale))
    root.geometry(f'{width}x{height}')
    root.minsize(min(round(680*scale),width),min(round(480*scale),height))
    app.style=ttk.Style(root);app.style.theme_use('clam');app.style.configure('.',font=('Microsoft YaHei UI',10))
    app.check_images=[]
    edge=max(18,round(16*root.winfo_fpixels('1i')/96))
    for checked in (False,True):
        picture=Image.new('RGBA',(edge+8,edge),(0,0,0,0));draw=ImageDraw.Draw(picture)
        draw.rounded_rectangle((1,1,edge-2,edge-2),radius=3,fill='#374767' if checked else '#FFFFFF',outline='#7B889A',width=1)
        if checked:draw.line([(edge*.23,edge*.51),(edge*.43,edge*.72),(edge*.78,edge*.28)],fill='white',width=max(2,round(edge*.12)))
        app.check_images.append(ImageTk.PhotoImage(picture,master=root));picture.close()
    app.style.element_create('Ruyi.indicator','image',app.check_images[0],('selected',app.check_images[1]))
    checkbox_layout=[('Checkbutton.padding',{'sticky':'nswe','children':[('Ruyi.indicator',{'side':'left','sticky':''}),('Checkbutton.label',{'side':'left','sticky':'nswe'})]})]
    app.style.layout('Ruyi.TCheckbutton',checkbox_layout);app.style.layout('Card.Ruyi.TCheckbutton',checkbox_layout)
    apply_theme(app)
    app.tabs=PageHost(root,app);app.tabs.pack(fill='both',expand=True)
    app.home=HomeCanvas(app.tabs,app)
    app.image_pane=task_class(app,'image');app.office_pane=task_class(app,'office');app.pdf_pane=task_class(app,'pdf');app.idphoto_pane=task_class(app,'idphoto')
    app.panes=[app.image_pane,app.office_pane,app.pdf_pane,app.idphoto_pane]
    for pane in app.panes:app.tabs.add(pane)
    app.settings_pane=SettingsPane(app.tabs,app)
    app.organizer_pane=None
    app.bottom=ttk.Frame(root,padding=(28,0,28,12));app.status=ttk.Label(app.bottom,text='选择文件后开始处理。',style='Muted.TLabel',wraplength=760);app.status.pack(side='left',fill='x',expand=True)
    app.stop=ttk.Button(app.bottom,text='停止处理',command=app.cancel.set,state='disabled');app.stop.pack(side='right')
    app.progress=ttk.Progressbar(root,mode='determinate')
    app._back_page=app.home
    def show(page):
        app.tabs.select(page)
        if page in (app.home,app.settings_pane,app.organizer_pane):
            app.bottom.pack_forget();app.progress.pack_forget()
        else:
            app.progress.pack(side='bottom',fill='x');app.bottom.pack(side='bottom',fill='x')
    app.show_home=lambda:show(app.home)
    def show_mode(mode):
        if mode=='organizer':
            if app.organizer_pane is None or app.organizer_pane.closed:
                from pdf_organizer.ui import OrganizerPane
                app.organizer_pane=OrganizerPane(app.tabs,app.tokens,app.show_home)
            show(app.organizer_pane)
        else:
            show(next(p for p in app.panes if p.mode==mode))
    app.show_mode=show_mode
    def show_settings():
        if app.tabs.current!=app.settings_pane:app._back_page=app.tabs.current
        app.settings_pane.refresh();show(app.settings_pane)
    app.show_settings=show_settings
    app.go_back=lambda:show(app._back_page)
    app.apply_theme=lambda:apply_theme(app)
    app.icon_photos=[]
    for size in (16,20,24,30,32,40,48,60,64,80,128,160,256):
        with Image.open(ASSETS/f'ruyi-icon-{size}.png') as icon_image:
            app.icon_photos.append(ImageTk.PhotoImage(icon_image,master=root))
    root.iconphoto(True,*app.icon_photos)
    try:root.iconbitmap(str(ASSETS/'ruyi.ico'))
    except tk.TclError:pass
    from desktop_integration import install_window_icons
    install_window_icons(root,ASSETS/'ruyi.ico')
    app.show_home()
    def resize_ended(_event):
        app.tabs.resize_page()
        if app.tabs.current is app.home:
            app.home.finish_live_resize()
        elif app.organizer_pane is not None and not app.organizer_pane.closed and app.tabs.current is app.organizer_pane:
            app.organizer_pane.grid_view.schedule_resize()
            app.organizer_pane.preview.schedule_resize()
        elif hasattr(app.tabs.current,'scroll'):
            app.tabs.current.scroll._resize()
    root.bind(RESIZE_ENDED,resize_ended,add='+')
    track_window_resize(root)
    root.bind('<Alt-Home>',lambda _e:app.show_home())
    root.bind('<Alt-Left>',lambda _e:app.go_back() if app.tabs.current==app.settings_pane else app.show_home())
    for index,mode in enumerate(('office','pdf','image','idphoto','organizer'),1):root.bind(f'<Alt-Key-{index}>',lambda _e,m=mode:app.show_mode(m))
    monitor_timer=None
    def monitor_system():
        nonlocal monitor_timer
        if app.settings['appearance']=='跟随系统' and app._system_dark!=system_dark():app.apply_theme()
        monitor_timer=root.after(2000,monitor_system)
    def stop_monitor(event):
        if event.widget is root and monitor_timer is not None:
            root.after_cancel(monitor_timer)
    monitor_timer=root.after(2000,monitor_system)
    root.bind('<Destroy>',stop_monitor,add='+')
