"""如意 1.1.0: local Windows image, Office and PDF toolbox."""
import os
import json
import queue
import sys
if __name__ == '__main__' and len(sys.argv) == 2 and sys.argv[1] == '--pdf-thumbnail-worker':
    from pdf_organizer.thumb_worker import main
    main()
    raise SystemExit(0)
if __name__ == '__main__' and len(sys.argv) == 2 and sys.argv[1] == '--pdf-preview-worker':
    from pdf_organizer.preview_worker import main
    main()
    raise SystemExit(0)
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import colorchooser, filedialog, messagebox, ttk

from PIL import Image, ImageTk, ImageDraw
from image_engine import convert_image, load_image
from office_engine import ConversionCancelled, OfficeSession, convert_office, find_libreoffice, select_native_engine
from native_office import find_powerpoint, find_word, find_excel, find_wps
from pdf_engine import convert_pdf
from id_photo_engine import replace_background
from id_photo_backgrounds import BACKGROUNDS, BACKGROUND_NAMES, resolve_background, make_background
from window_paint import WindowPaintGate
from desktop_ui import TaskPaneView, build_shell, VERSION
from desktop_integration import enable_dpi_awareness
enable_dpi_awareness()

OFFICE_LAYOUT_SUITES = {'自动匹配原编辑器': 'auto', 'WPS': 'wps',
                       'Microsoft Office': 'microsoft', 'LibreOffice 兼容输出': 'libreoffice'}


def size_text(count):
    return f'{count / 1024:.1f} KB' if count < 1024**2 else f'{count / 1024**2:.2f} MB'


class TaskPane(TaskPaneView):
    def __init__(self, app, mode):
        super().__init__(app, mode)

    def job_items(self, process_all=False):
        selected = set(self.tree.selection())
        return [(row, self.files[row]) for row in self.tree.get_children()
                if row in self.files and (process_all or not selected or row in selected)]

    def selection_changed(self, _event=None):
        total = len(self.files)
        chosen = set(self.tree.selection())
        selected = sum(row in self.files for row in chosen)
        for row in self.files:
            if hasattr(self.app, 'check_images'):
                self.tree.item(row, image=self.app.check_images[int(row in chosen)])
            else:
                self.tree.item(row, text='☑' if row in chosen else '☐')
        self.selection_label.configure(text=f'共 {total} 个文件 · 已选 {selected} 个')
        text = f'处理所选（{selected}）' if selected else f'处理全部（{total}）'
        self.start_button.configure(text=text, state='disabled' if self.locked or self.app.busy or not total else 'normal')
        self.all_button.configure(state='disabled' if self.locked or self.app.busy or not total else 'normal')
        if 0 < selected < total:
            self.all_button.pack(side='left', after=self.start_button, padx=(8, 0))
        else:
            self.all_button.pack_forget()
        self.preview()

    def select_all(self, _event=None):
        self.tree.selection_set(self.tree.get_children())
        self.selection_changed()
        return 'break'

    def clear_selection(self, _event=None):
        self.tree.selection_remove(self.tree.selection())
        self.range_anchor = None
        self.selection_changed()
        return 'break'

    def selection_anchor_changed(self, event):
        row = self.tree.identify_row(event.y)
        if row in self.files:
            self.tree.focus_set()
        if self.tree.identify_column(event.x) == '#0' and row in self.files:
            self.tree.focus(row)
            self.toggle_focused()
            return 'break'
        if not event.state & 1:
            self.range_anchor = row or None

    def toggle_focused(self, _event=None):
        row = self.tree.focus()
        if row in self.files:
            if row in self.tree.selection():
                self.tree.selection_remove(row)
            else:
                self.tree.selection_add(row)
            self.range_anchor = row
            self.selection_changed()
        return 'break'

    def move_selection(self, step, *, extend=False, preserve=False):
        rows = self.tree.get_children()
        if not rows:
            return 'break'
        focus = self.tree.focus()
        current = rows.index(focus) if focus in rows else 0
        next_index = max(0, min(len(rows)-1, current+step))
        row = rows[next_index]
        self.tree.focus(row)
        self.tree.see(row)
        if extend:
            if self.range_anchor not in rows:
                self.range_anchor = rows[current]
            begin, end = sorted((rows.index(self.range_anchor), next_index))
            self.tree.selection_set(rows[begin:end+1])
        elif not preserve:
            self.tree.selection_set(row)
            self.range_anchor = row
        else:
            self.range_anchor = None
        self.selection_changed()
        return 'break'

    def format_changed(self, _event=None):
        png = self.format_var.get() == 'PNG'
        lossless = self.lossless_var.get()
        self.quality_scale.configure(state='disabled' if png or lossless or self.app.busy else 'normal')
        self.hint.configure(text='无损保存缩放后的像素；减小尺寸仍会减少细节。目标过小会提示，不会暗中缩图。' if png or lossless else
                            '画质 1–100 是编码参数，JPEG 100% 非无损。目标 KB 自动降低画质，保持所设尺寸。')
        self.scroll.refresh()

    def pdf_target_changed(self, _event=None):
        word = self.target_var.get() == 'Word'
        choices = ['自动还原'] if word else ['可编辑重建', '原版外观（图片）']
        self.layout_box.configure(values=choices)
        if word:
            self.layout_label.pack_forget()
            self.layout_box.pack_forget()
        else:
            self.layout_label.pack(side='left', padx=(14, 0))
            self.layout_box.pack(side='left', padx=10)
        if self.layout_var.get() not in choices:
            self.layout_var.set(choices[0])
        self.pdf_hint.configure(text=(
            'Word 自动匹配原稿分页、字体、段落及表格；文字和支持的表格可编辑，扫描内容保留为图片。'
            if word else '重建可编辑文字、表格、图片和线条；复杂版式请核对。原版外观按整页图片保留，页内文字不可单独编辑；扫描页不执行 OCR。'))
        self.scroll.refresh()

    def edge_options_changed(self):
        if self.advanced_edge_var.get():
            self.edge_options.pack(side='left', padx=5)
        else:
            self.edge_options.pack_forget()
            self.edge_var.set('自动精细')
        self.scroll.refresh()

    def color_changed(self, _event=None):
        self.color_button.configure(state='normal' if self.color_var.get() == '自定义' and not self.app.busy else 'disabled')
        background = self.selected_background()
        image = make_background((56, 28), background)
        try:
            self.color_swatch_photo = ImageTk.PhotoImage(image, master=self)
        finally:
            image.close()
        self.color_swatch.configure(image=self.color_swatch_photo, background=background.start)
        def rgb_text(color):
            return ', '.join(str(int(color[i:i + 2], 16)) for i in (1, 3, 5))
        if background.gradient:
            color_text = f'{background.start} → {background.end} · 上浅下深\nRGB {rgb_text(background.start)} → {rgb_text(background.end)}'
        else:
            color_text = f'{background.start} · RGB {rgb_text(background.start)}'
        self.color_label.configure(text=color_text)
        reference = background.usage if background.code != 'CUSTOM' else '自定义背景色'
        self.color_usage.configure(text=f'{reference}；以受理方要求为准。默认自动精细抠图，保留原尺寸与肤色。')
        self.scroll.refresh()

    def selected_background(self):
        value = self.custom_color if self.color_var.get() == '自定义' else self.color_var.get()
        return resolve_background(value)

    def choose_color(self):
        selected = colorchooser.askcolor(color=self.selected_background().start, parent=self, title='选择证件照背景色')
        if selected and selected[1]:
            self.color_var.set('自定义')
            self.custom_color = resolve_background(selected[1]).start
            self.color_changed()

    def add_files(self):
        if self.app.busy:
            return
        formats = {'image': [('静态图片', '*.jpg *.jpeg *.png *.webp')],
                   'office': [('办公文档', '*.doc *.docx *.xls *.xlsx *.ppt *.pptx')],
                   'pdf': [('PDF 文件', '*.pdf')],
                   'idphoto': [('照片', '*.jpg *.jpeg *.png *.webp')]}[self.mode]
        selected = filedialog.askopenfilenames(parent=self, title='选择待处理文件', filetypes=formats + [('全部文件', '*.*')])
        self.add_paths(selected)

    def add_paths(self, selected):
        if self.app.busy:
            return
        for filename in selected:
            path = Path(filename).resolve()
            if path in self.files.values():
                continue
            row = self.tree.insert('', 'end', values=(path.name, '等待', ''))
            self.files[row] = path
        if selected and not self.output_var.get():
            self.output_var.set(str(Path(selected[0]).parent / '如意输出'))
        self.selection_changed()

    def clear(self):
        if self.app.busy:
            return
        for row in self.tree.get_children():
            self.tree.delete(row)
        self.files.clear()
        self.outputs.clear()
        self.details.clear()
        self.range_anchor = None
        self.selection_changed()

    def show_details(self, _event=None):
        selection = self.tree.selection()
        if not selection:
            return
        row = selection[0]
        notes = self.details.get(row) or self.tree.set(row, 'result')
        if not notes:
            return
        popup = tk.Toplevel(self)
        popup.title('转换详情')
        popup.geometry('780x420')
        # Long paths and unspaced CJK runs are single "words" to Tk, so word
        # wrap leaves them wider than the window while only a vertical bar
        # exists: the tail of the message becomes unreachable. Char wrap keeps
        # every character visible and the layout still scrolls vertically.
        text = tk.Text(popup, wrap='char', padx=12, pady=12)
        scroll = ttk.Scrollbar(popup, command=text.yview)
        text.configure(yscrollcommand=scroll.set)
        scroll.pack(side='right', fill='y')
        text.pack(fill='both', expand=True)
        text.insert('1.0', notes)
        text.configure(state='disabled')
        popup.bind('<Escape>', lambda _e: popup.destroy())

    def browse_output(self):
        path = filedialog.askdirectory(parent=self, title='选择输出文件夹')
        if path:
            self.output_var.set(path)

    def open_output(self):
        path = Path(self.output_var.get()) if self.output_var.get() else None
        if path and path.is_dir():
            os.startfile(path)
        else:
            messagebox.showinfo('输出文件夹', '完成转换后，可在这里打开结果文件夹。', parent=self)

    def select_engine(self):
        path = filedialog.askopenfilename(parent=self, title='选择 LibreOffice 的 soffice.exe 或 soffice.com',
                                           filetypes=[('LibreOffice', 'soffice.exe soffice.com')])
        if path:
            self.app.engine = find_libreoffice(path)
        self.refresh_engine()
        if hasattr(self.app, 'settings'):
            from desktop_integration import save_settings
            self.app.settings['libreoffice'] = str(self.app.engine) if self.app.engine else ''
            try:
                save_settings(self.app.settings)
            except OSError:
                messagebox.showerror('设置未保存', '无法保存转换引擎位置。', parent=self)

    def refresh_engine(self):
        wps = [('文字', bool(find_wps('word'))), ('表格', bool(find_wps('sheet'))),
               ('演示', bool(find_wps('slides')))]
        microsoft = [('Word', bool(find_word())), ('Excel', bool(find_excel())),
                     ('PowerPoint', bool(find_powerpoint()))]
        def status(name, applications):
            ready = '、'.join(label for label, available in applications if available)
            missing = '、'.join(label for label, available in applications if not available)
            return name + '：' + ('已检测到 ' + ready if ready else '') + (
                ('；' if ready else '') + missing + ' 无导出接口' if missing else '')
        message = '\n'.join([status('WPS', wps), status('Microsoft Office', microsoft),
                             'LibreOffice：' + ('兼容输出已就绪' if self.app.engine else '未检测到')])
        self.engine_label.configure(text=message)
        self.scroll.refresh()

    def lock(self, busy):
        self.locked = busy
        for control in self.controls:
            state = 'disabled' if busy else ('readonly' if isinstance(control, ttk.Combobox) else 'normal')
            control.configure(state=state)
        if self.mode == 'image':
            self.format_changed()
        elif self.mode == 'idphoto':
            self.color_changed()
        self.selection_changed()

    def preview(self, _event=None):
        if self.mode not in ('image', 'idphoto'):
            return
        self.photos.clear()
        selection = self.tree.selection()
        row = selection[0] if selection else None
        for label, path, caption in [(self.before, self.files.get(row), '原图'), (self.after, self.outputs.get(row), '结果')]:
            if not path:
                label.configure(image='', text=caption + '预览')
                continue
            try:
                image = load_image(path)
                image.thumbnail((360, 125), Image.Resampling.LANCZOS)
                photo = ImageTk.PhotoImage(image, master=self)
                image.close()
                self.photos.append(photo)
                label.configure(image=photo, text=caption, compound='left')
            except Exception:
                label.configure(image='', text=caption + '暂不可预览')


class RuyiApp:
    def __init__(self, root):
        self.root = root
        self.busy = False
        self.closing = False
        self.engine = find_libreoffice()
        self.office_session = None
        self.job_started = None
        self.job_name = ''
        self.cancel = threading.Event()
        self.events = queue.Queue()
        self.paint_gate = WindowPaintGate(root)
        build_shell(self, TaskPane)
        root.protocol('WM_DELETE_WINDOW', self.close)
        self.poll_timer = root.after(100, self.poll)
        root.bind('<Destroy>', self.stop_polling, add='+')

    def stop_polling(self, event):
        if event.widget is self.root and self.poll_timer is not None:
            self.root.after_cancel(self.poll_timer)
            self.poll_timer = None

    def start(self, pane, process_all=False):
        if self.busy:
            return
        if not pane.files:
            messagebox.showinfo('添加文件', '请先添加要处理的文件。', parent=self.root)
            return
        jobs = pane.job_items(process_all=process_all)
        options = {}
        if pane.mode == 'image':
            try:
                maximum = int(pane.max_var.get()) if pane.max_var.get().strip() else None
                if maximum is not None and not 100 <= maximum <= 12000:
                    raise ValueError
            except ValueError:
                messagebox.showerror('尺寸不正确', '最长边请输入 100 到 12000 的整数，或留空。', parent=self.root)
                return
            try:
                quality = int(pane.quality_var.get())
                target = int(pane.target_var.get()) if pane.target_var.get().strip() else None
                if not 1 <= quality <= 100 or (target is not None and not 1 <= target <= 102400):
                    raise ValueError
            except ValueError:
                messagebox.showerror('参数不正确', '画质请输入 1–100 的整数，目标大小请输入 1–102400 KB，或留空。', parent=self.root)
                return
            options = dict(output_format='ORIGINAL' if pane.format_var.get() == '原格式' else pane.format_var.get(),
                           quality=quality, max_side=maximum, target_kb=target,
                           lossless=pane.lossless_var.get(), only_smaller=True)
        elif pane.mode == 'office':
            self.engine = self.engine or find_libreoffice()
            pane.refresh_engine()
            layout_suite = OFFICE_LAYOUT_SUITES[pane.office_layout_var.get()]
            native_choices = [select_native_engine(path, layout_suite) for _row, path in jobs]
            can_use_compatible = layout_suite in ('auto', 'libreoffice') and bool(self.engine)
            if not can_use_compatible and not any(native_choices):
                selected = pane.office_layout_var.get()
                message = (f'当前文件未找到可用的“{selected}”导出接口。请安装或修复对应办公软件后重试。'
                           if layout_suite in ('wps', 'microsoft') else
                           '当前文件未找到可用的导出引擎。请安装 WPS / Microsoft Office，或选择 LibreOffice 进行兼容输出。')
                if layout_suite == 'libreoffice':
                    message = '请安装 LibreOffice，或点击“选择 LibreOffice”指定 soffice.exe 后重试。'
                messagebox.showinfo('导出引擎不可用', message, parent=self.root)
                return
        if not pane.output_var.get().strip():
            messagebox.showinfo('保存位置', '请选择保存位置。', parent=self.root)
            return
        output = Path(pane.output_var.get()).expanduser().resolve()
        try:
            output.mkdir(parents=True, exist_ok=True)
        except OSError as error:
            messagebox.showerror('无法写入', str(error), parent=self.root)
            return
        self.busy = True
        self.active_rows = tuple(row for row, _path in jobs)
        self.cancel.clear()
        for row, _path in jobs:
            pane.outputs.pop(row, None)
            pane.details.pop(row, None)
        for item in self.panes:
            item.lock(True)
        self.stop.configure(state='normal')
        self.progress.configure(maximum=len(jobs), value=0)
        for row, path in jobs:
            pane.tree.item(row, values=(path.name, '等待', ''))
        self.status.configure(text=f'共 {len(jobs)} 个文件，正在本机处理…')
        engine = self.engine
        office_layout_suite = OFFICE_LAYOUT_SUITES[pane.office_layout_var.get()] if pane.mode == 'office' else 'auto'
        pdf_target = {'Word': 'docx', 'Excel': 'xlsx', 'PPT': 'pptx'}.get(
            pane.target_var.get() if pane.mode == 'pdf' else '')
        pdf_layout = ('fidelity' if pdf_target == 'docx' else
                      ('visual' if pane.layout_var.get() == '原版外观（图片）' else 'editable')) if pane.mode == 'pdf' else 'editable'
        id_background = pane.selected_background() if pane.mode == 'idphoto' else None
        id_color = (id_background.start if id_background.code == 'CUSTOM' else id_background.code) if id_background else None
        id_edge = pane.edge_var.get() if pane.mode == 'idphoto' else '自动精细'

        def work():
            succeeded = failed = processed = 0
            try:
                for row, path in jobs:
                    if self.cancel.is_set():
                        break
                    self.events.put(('running', pane, row, path.name))
                    try:
                        if pane.mode == 'office' and self.office_session is None:
                            self.office_session = OfficeSession()
                        if pane.mode == 'image':
                            result = convert_image(path, output, **options)
                        elif pane.mode == 'office':
                            result = convert_office(path, output, engine, self.cancel, session=self.office_session,
                                                    layout_suite=office_layout_suite)
                        elif pane.mode == 'pdf':
                            result = convert_pdf(path, output, pdf_target, self.cancel, layout_mode=pdf_layout)
                        else:
                            result = replace_background(path, output, id_color, self.cancel, edge_mode=id_edge)
                        if pane.mode == 'image':
                            delta = 100 * (1 - result.output_bytes / result.original_bytes)
                            change = f'减少 {delta:.1f}%' if delta >= 0 else f'增加 {-delta:.1f}%'
                            detail = f'{size_text(result.original_bytes)} → {size_text(result.output_bytes)}，{change} {result.note}'
                        elif pane.mode == 'office':
                            detail = f'PDF · {size_text(result.output_bytes)} · {result.elapsed_seconds:.1f} 秒 · {getattr(result,"engine_name","LibreOffice")}'
                            if getattr(result, 'engine_name', 'LibreOffice') == 'LibreOffice':
                                detail += ' · 兼容输出，请核对排版'
                            if result.resource_summary:
                                detail += ' · ' + result.resource_summary
                        elif pane.mode == 'pdf':
                            detail = f'{result.output_path.suffix[1:].upper()} · {size_text(result.output_bytes)} · {result.pages} 页 · {result.elapsed_seconds:.1f} 秒'
                        else:
                            output_path, output_bytes, elapsed = result
                            detail = f'JPG · {size_text(output_bytes)} · {elapsed:.1f} 秒'
                        if pane.mode == 'office':
                            notes = '导出引擎：'+getattr(result,'engine_name','LibreOffice')+'\n'+'\n'.join(result.resource_notes)
                            output_path = result.output_path
                        elif pane.mode == 'pdf':
                            notes = result.note
                            output_path = result.output_path
                        elif pane.mode == 'image':
                            notes = result.note
                            output_path = result.output_path
                        else:
                            color_description = id_background.start + ((' → ' + id_background.end) if id_background.gradient else '')
                            notes = f'背景：{id_background.name}（{color_description}）。\n抠图：{id_edge}。人像分割在本机完成；未上传照片。预设用途供参考，以受理方要求为准。'
                        self.events.put(('success', pane, row, output_path, detail, notes))
                        succeeded += 1
                    except ConversionCancelled:
                        self.events.put(('cancelled', pane, row))
                        break
                    except Exception as error:
                        self.events.put(('failed', pane, row, str(error)))
                        failed += 1
                    processed += 1
                    self.events.put(('progress', processed))
            finally:
                self.events.put(('done', pane, succeeded, failed, self.cancel.is_set()))

        threading.Thread(target=work, daemon=False).start()

    def poll(self):
        self.poll_timer = None
        try:
            while True:
                event = self.events.get_nowait()
                kind = event[0]
                if kind == 'progress':
                    self.progress.configure(value=event[1])
                elif kind == 'done':
                    _, pane, succeeded, failed, stopped = event
                    self.busy = False
                    self.job_started = None
                    for item in self.panes:
                        item.lock(False)
                    self.stop.configure(state='disabled')
                    for row in self.active_rows:
                        if pane.tree.set(row, 'status') == '等待':
                            pane.tree.set(row, 'status', '未处理')
                    self.active_rows = ()
                    self.status.configure(text=f'{"已停止" if stopped else "处理完成"}：成功 {succeeded} 个，失败 {failed} 个。结果保存在所选文件夹。')
                    pane.preview()
                    if self.closing:
                        if self.office_session:
                            self.office_session.close()
                        self.root.destroy()
                        return
                else:
                    pane, row = event[1:3]
                    if kind == 'running':
                        self.job_started = time.monotonic()
                        self.job_name = event[3]
                        pane.tree.set(row, 'status', '处理中')
                        self.status.configure(text='正在处理：' + event[3])
                    elif kind == 'success':
                        pane.tree.set(row, 'status', '已完成')
                        pane.tree.set(row, 'result', event[4])
                        pane.outputs[row] = event[3]
                        pane.details[row] = event[5]
                    elif kind == 'failed':
                        pane.tree.set(row, 'status', '失败')
                        pane.tree.set(row, 'result', event[3])
                    elif kind == 'cancelled':
                        pane.tree.set(row, 'status', '已停止')
        except queue.Empty:
            pass
        if self.busy and self.job_started:
            self.status.configure(text=f'正在处理：{self.job_name} · 已用 {int(time.monotonic()-self.job_started)} 秒')
        self.poll_timer = self.root.after(100, self.poll)

    def close(self):
        organizer = self.organizer_pane
        if organizer is not None and organizer.winfo_exists():
            organizer.after_close = self.close
            organizer.on_close()
            return
        if self.busy:
            if messagebox.askyesno('停止并退出', '正在处理文件。停止后退出？\n图片会在当前一张处理结束后停止，Office 转换会立即停止。', parent=self.root):
                self.closing = True
                self.cancel.set()
                self.status.configure(text='正在停止并清理，请稍候…')
        else:
            if self.office_session:
                self.office_session.close()
            self.root.destroy()


if __name__ == '__main__':
    if len(sys.argv) == 3 and sys.argv[1] == '--verify-conversions':
        from hashlib import sha256
        job = json.loads(Path(sys.argv[2]).read_text(encoding='utf-8-sig'))
        report = {'ok': False, 'results': []}
        report_path = Path(job['report'])
        report_path.parent.mkdir(parents=True, exist_ok=True)
        try:
            output = Path(job['output'])
            output.mkdir(parents=True, exist_ok=True)
            pdf_jobs = list(job.get('pdf_jobs', []))
            if job.get('pdf'): pdf_jobs.insert(0, {'source': job['pdf'], 'target': 'docx'})
            for pdf_job in pdf_jobs:
                source = Path(pdf_job['source'])
                before = sha256(source.read_bytes()).hexdigest()
                target = pdf_job.get('target', 'docx')
                layout_mode = pdf_job.get('layout_mode', 'editable')
                result = convert_pdf(source, output, target, layout_mode=layout_mode)
                if before != sha256(source.read_bytes()).hexdigest():
                    raise ValueError('PDF 原文件发生变化。')
                report['results'].append({'kind': 'pdf-' + target, 'layout_mode': layout_mode, 'path': str(result.output_path),
                                          'pages': result.pages, 'note': result.note,
                                          'seconds': result.elapsed_seconds, 'source_unchanged': True})
            for office_job in job.get('office_jobs', []):
                source=Path(office_job['source'])
                before=sha256(source.read_bytes()).hexdigest()
                layout_suite=office_job.get('layout_suite',job.get('layout_suite','auto'))
                timeout=float(office_job.get('timeout',job.get('timeout',120)))
                result=convert_office(source,output,layout_suite=layout_suite,timeout=timeout)
                if before != sha256(source.read_bytes()).hexdigest():
                    raise ValueError('Office原文件发生变化。')
                report['results'].append({'kind':'office-pdf','path':str(result.output_path),
                                          'layout_suite':layout_suite,'timeout':timeout,
                                          'engine':result.engine_name,'notes':result.resource_notes,
                                          'seconds':result.elapsed_seconds,'source_unchanged':True})
            if job.get('photo'):
                source = Path(job['photo'])
                before = sha256(source.read_bytes()).hexdigest()
                for color in job.get('colors', ['标准蓝', '正红']):
                    background = resolve_background(color)
                    path, byte_count, elapsed = replace_background(source, output, color,
                                                                  edge_mode=job.get('edge_mode', '自动精细'))
                    report['results'].append({'kind': 'id-photo', 'color': color, 'path': str(path),
                                              'background_code': background.code, 'background_name': background.name,
                                              'start': background.start, 'end': background.end,
                                              'bytes': byte_count, 'seconds': elapsed})
                if before != sha256(source.read_bytes()).hexdigest():
                    raise ValueError('照片原文件发生变化。')
                report['photo_source_unchanged'] = True
            report['ok'] = True
        except Exception as error:
            report['error'] = repr(error)
        report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
        raise SystemExit(0 if report['ok'] else 1)
    if len(sys.argv) == 3 and sys.argv[1] == '--self-check':
        report_path = Path(sys.argv[2])
        report_path.parent.mkdir(parents=True, exist_ok=True)
        try:
            from importlib.metadata import version
            import pdfplumber
            import docx
            import openpyxl
            import pptx
            import onnxruntime
            import pypdfium2
            import pypdf
            import reportlab
            from pdf_organizer import composition, exporter, overlays, presets, rules_ui
            from id_photo_engine import _session
            matting_session = _session()
            root = tk.Tk()
            root.withdraw()
            RuyiApp(root)
            root.update_idletasks()
            root.destroy()
            report_path.write_text(json.dumps({'ok': True, 'version': VERSION,
                                               'default_office_layout_suite': 'auto',
                                               'native_layout_interfaces': {
                                                   'wps_word': bool(find_wps('word')), 'wps_sheet': bool(find_wps('sheet')),
                                                   'wps_slides': bool(find_wps('slides')), 'word': bool(find_word()),
                                                   'excel': bool(find_excel()), 'powerpoint': bool(find_powerpoint())},
                                               'default_word_mode': 'fidelity', 'default_matting_mode': '自动精细',
                                               'backgrounds': [{'code': item.code, 'name': item.name,
                                                                'start': item.start, 'end': item.end}
                                                               for item in BACKGROUNDS],
                                               'matting_model': 'MODNet FP32',
                                               'matting_providers': matting_session.get_providers(),
                                               'pdfplumber': pdfplumber.__version__,
                                               'python-docx': docx.__version__,
                                               'openpyxl': openpyxl.__version__, 'python-pptx': pptx.__version__,
                                               'onnxruntime': onnxruntime.__version__, 'pypdfium2': version('pypdfium2'),
                                               'pypdf': pypdf.__version__, 'reportlab': reportlab.Version,
                                               'pdf_organizer_modules_loaded': True},
                                              ensure_ascii=False), encoding='utf-8')
        except Exception as error:
            report_path.write_text(json.dumps({'ok': False, 'error': repr(error)}, ensure_ascii=False), encoding='utf-8')
            raise
        raise SystemExit(0)
    root = tk.Tk()
    # Build and lay out the complete light surface while unmapped. Showing a
    # newly-created Tk HWND before its first themed paint can expose a black
    # frame when launched/restored from the taskbar on Windows.
    root.withdraw()
    application = RuyiApp(root)
    if len(sys.argv) == 3 and sys.argv[1] == '--verify-ui':
        from ui_verify import prepare_verification
        prepare_verification(root, application, sys.argv[2])
    root.update_idletasks()
    root.after_idle(root.deiconify)
    root.mainloop()
