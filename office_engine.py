"""Isolated LibreOffice jobs for DOCX/XLSX/PPTX to PDF."""
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import zipfile
import zlib
import re
import olefile
from contextlib import contextmanager
import warnings
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree as ET

from file_output import publish_file
from external_resources import prepare_ooxml
from native_office import find_powerpoint, export_powerpoint, find_word, export_word, find_excel, export_excel, find_wps, export_wps

FILTERS = {'.docx': 'writer_pdf_Export', '.xlsx': 'calc_pdf_Export', '.pptx': 'impress_pdf_Export'}
FILTERS.update({'.doc': 'writer_pdf_Export', '.xls': 'calc_pdf_Export', '.ppt': 'impress_pdf_Export'})
MARKERS = {'.docx': 'word/document.xml', '.xlsx': 'xl/workbook.xml', '.pptx': 'ppt/presentation.xml'}
OFFICE_KINDS = {'.doc': 'word', '.docx': 'word', '.xls': 'sheet', '.xlsx': 'sheet', '.ppt': 'slides', '.pptx': 'slides'}


def document_suite(source: Path) -> str | None:
    """The last editor is a hint, not proof of the user's preferred renderer."""
    if source.suffix.lower() not in MARKERS:
        return None
    try:
        with zipfile.ZipFile(source) as archive:
            info=archive.getinfo('docProps/app.xml')
            if info.file_size>1024*1024:return None
            root=ET.fromstring(archive.read(info))
        app=' '.join(n.text or '' for n in root.iter() if n.tag.rsplit('}',1)[-1]=='Application').lower()
        if 'wps' in app or 'kingsoft' in app:return 'wps'
        if 'microsoft' in app:return 'microsoft'
    except (OSError,KeyError,zipfile.BadZipFile,ET.ParseError,RuntimeError,NotImplementedError,zlib.error):
        pass
    return None


def select_native_engine(source: Path, suite: str = 'auto'):
    """Keep the chosen layout engine stable; a failure must not silently switch it."""
    if suite not in ('auto','wps','microsoft','libreoffice'):
        raise ValueError('未知的排版基准。')
    kind=OFFICE_KINDS.get(source.suffix.lower())
    if not kind or suite=='libreoffice':return None
    choices=(suite,) if suite!='auto' else (('microsoft','wps') if document_suite(source)=='microsoft' else ('wps','microsoft'))
    for choice in choices:
        available=find_wps(kind) if choice=='wps' else {'word':find_word,'sheet':find_excel,'slides':find_powerpoint}[kind]()
        if available:return choice,kind
    return None


class ConversionCancelled(Exception):
    pass


@dataclass(frozen=True)
class OfficeResult:
    output_path: Path
    output_bytes: int
    elapsed_seconds: float
    resource_summary: str = ''
    resource_notes: tuple[str, ...] = ()
    engine_name: str = 'LibreOffice'


def cleanup_workspace(folder: Path) -> None:
    # Windows may keep child-process file handles alive briefly after taskkill.
    folder = folder.resolve()
    if folder.parent != Path(tempfile.gettempdir()).resolve() or not folder.name.startswith('ruyi-'):
        raise ValueError('拒绝清理非如意临时目录。')
    for attempt in range(80):
        try:
            shutil.rmtree(folder)
            return
        except FileNotFoundError:
            return
        except OSError:
            if attempt == 79:
                warnings.warn(f'临时目录仍被占用，留待系统清理：{folder}', RuntimeWarning)
                return
            time.sleep(.1)


@contextmanager
def job_workspace():
    folder = Path(tempfile.mkdtemp(prefix='ruyi-office-'))
    try:
        yield folder
    finally:
        cleanup_workspace(folder)


def find_libreoffice(explicit: str | Path | None = None) -> Path | None:
    candidates = []
    chosen = explicit or os.environ.get('RUYI_LIBREOFFICE')
    if chosen:
        candidates.append(Path(chosen))
    else:
        for variable in ('ProgramFiles', 'ProgramFiles(x86)'):
            if os.environ.get(variable):
                candidates.append(Path(os.environ[variable]) / 'LibreOffice/program/soffice.com')
        local = os.environ.get('LOCALAPPDATA')
        if local:
            candidates.extend(sorted((Path(local) / 'Ruyi/engines').glob('LibreOffice-*/program/soffice.com'), reverse=True))
        command = shutil.which('soffice.com') or shutil.which('soffice')
        if command:
            candidates.append(Path(command))
    for candidate in candidates:
        if candidate.suffix.lower() == '.exe' and candidate.with_suffix('.com').is_file():
            candidate = candidate.with_suffix('.com')
        if candidate.is_file():
            return candidate.resolve()
    return None


def validate_office(source: Path) -> None:
    suffix = source.suffix.lower()
    if suffix not in FILTERS:
        raise ValueError('当前支持 DOC/DOCX、XLS/XLSX、PPT/PPTX。')
    if source.stat().st_size > 100 * 1024 * 1024:
        raise ValueError('当前版本支持 100 MB 以内的 Office 文件。')
    if suffix in ('.doc', '.xls', '.ppt'):
        try:
            with olefile.OleFileIO(source) as archive:
                expected = {'.doc': ('WordDocument',), '.xls': ('Workbook', 'Book'), '.ppt': ('PowerPoint Document',)}[suffix]
                if not any(archive.exists(name) for name in expected):
                    raise ValueError('文件内容与扩展名不符，或文档已损坏。')
                names = [part.lower() for path in archive.listdir(streams=True, storages=True) for part in path]
                if any(name in ('vba', '_vba_project', '_vba_project_cur', 'macros') for name in names):
                    raise ValueError('当前版本不处理含宏的文档。')
                if 'encryptedpackage' in names or 'encryptioninfo' in names:
                    raise ValueError('请先解除文档加密后再转换。')
        except OSError as error:
            raise ValueError('旧版 Office 文件损坏或内容与扩展名不符。') from error
        return
    try:
        with zipfile.ZipFile(source) as archive:
            entries = archive.infolist()
            if len(entries) > 10000 or sum(x.file_size for x in entries) > 500 * 1024 * 1024:
                raise ValueError('文档解压后过大，当前版本暂不处理。')
            names = set(archive.namelist())
            if MARKERS[suffix] not in names or '[Content_Types].xml' not in names:
                raise ValueError('文件内容与扩展名不符，或文档已损坏。')
            if any(x.lower().endswith('vbaproject.bin') for x in names):
                raise ValueError('当前版本不处理含宏的文档，请先另存为不含宏的文件。')
            for item in entries:
                if item.filename.endswith('.rels'):
                    if item.file_size > 2 * 1024 * 1024:
                        raise ValueError('文档关系信息过大，暂不处理。')
                    ET.fromstring(archive.read(item))
    except (zipfile.BadZipFile, ET.ParseError, RuntimeError) as error:
        raise ValueError('文档损坏或已加密，请在 Office 中检查并另存。') from error


def _stop_process(process: subprocess.Popen) -> None:
    if process.poll() is not None:
        return
    if os.name == 'nt':
        subprocess.run(['taskkill.exe', '/PID', str(process.pid), '/T', '/F'],
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                       creationflags=subprocess.CREATE_NO_WINDOW, timeout=15)
    else:
        process.kill()
    try:
        process.wait(timeout=10)
    except subprocess.TimeoutExpired:
        process.kill()
        process.wait(timeout=5)


class OfficeSession:
    """Reuse one isolated profile for sequential jobs; never use the user's profile."""
    def __init__(self):
        self.folder = Path(tempfile.mkdtemp(prefix='ruyi-session-'))
        self.profile = self.folder / 'profile'
        user = self.profile / 'user'
        user.mkdir(parents=True)
        (user / 'registrymodifications.xcu').write_text(
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<oor:items xmlns:oor="http://openoffice.org/2001/registry">'
            '<item oor:path="/org.openoffice.Office.Common/Security/Scripting">'
            '<prop oor:name="MacroSecurityLevel" oor:op="fuse"><value>3</value></prop>'
            '<prop oor:name="DisableMacrosExecution" oor:op="fuse"><value>true</value></prop>'
            # Native charts are also embedded objects; disabling all active content hides them.
            '<prop oor:name="DisableActiveContent" oor:op="fuse"><value>false</value></prop>'
            '<prop oor:name="DisableOLEAutomation" oor:op="fuse"><value>true</value></prop>'
            '<prop oor:name="BlockUntrustedRefererLinks" oor:op="fuse"><value>true</value></prop>'
            '</item>'
            '<item oor:path="/org.openoffice.Office.Writer/Content/Update">'
            '<prop oor:name="Link" oor:op="fuse"><value>2</value></prop>'
            '</item>'
            '<item oor:path="/org.openoffice.Office.Calc/Content/Update">'
            '<prop oor:name="Link" oor:op="fuse"><value>1</value></prop>'
            '</item></oor:items>', encoding='utf-8')

    def close(self):
        cleanup_workspace(self.folder)

    def __enter__(self):
        return self

    def __exit__(self, *_):
        self.close()


def _fit_plain_workbook(path: Path) -> None:
    """Give a plain data sheet a printable page setup on the temporary copy.

    Spreadsheet applications default to fitting everything on one page, so a wide
    raw-data sheet is shrunk into a corner and every cell is truncated to ###.
    Only workbooks without charts, pivots, drawings or media are rewritten:
    openpyxl re-saving richer files can drop those parts, so they keep the
    author's own layout instead.
    """
    try:
        with zipfile.ZipFile(path) as package:
            names = set(package.namelist())
        blocked = ('xl/charts/', 'xl/pivotCache/', 'xl/drawings/', 'xl/media/',
                   'xl/pivotTables/', 'xl/slicers/')
        if any(name.startswith(blocked) for name in names):
            return
        import openpyxl
        from openpyxl.worksheet.properties import PageSetupProperties
        book = openpyxl.load_workbook(path)
        try:
            changed = False
            for sheet in book.worksheets:
                setup = sheet.page_setup
                if (getattr(sheet, 'print_area', None) or setup.scale or
                        setup.fitToWidth or setup.fitToHeight):
                    continue
                sheet.page_setup.orientation = 'landscape'
                sheet.page_setup.fitToWidth = 1
                sheet.page_setup.fitToHeight = 0
                sheet.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
                changed = True
            if changed:
                book.save(path)
        finally:
            book.close()
    except Exception:
        return


def convert_office(source: Path, output_dir: Path, executable: Path | None = None,
                   cancel: threading.Event | None = None, timeout: float = 120,
                   session: OfficeSession | None = None, *, _depth=0, _origin=None, prefer_native=True, layout_suite='auto') -> OfficeResult:
    if session is None:
        with OfficeSession() as isolated:
            return convert_office(source, output_dir, executable, cancel, timeout, isolated, _depth=_depth, _origin=_origin, prefer_native=prefer_native,layout_suite=layout_suite)
    source = Path(source).resolve(strict=True)
    output_dir = Path(output_dir).resolve(strict=True)
    validate_office(source)
    engine = find_libreoffice(executable)
    native = select_native_engine(source,layout_suite) if prefer_native and _depth==0 else None
    if prefer_native and _depth==0 and layout_suite in ('wps','microsoft') and not native:
        raise ValueError(('WPS' if layout_suite=='wps' else 'Microsoft Office')+' 的本机导出接口不可用；未更换排版引擎，请安装或修复对应办公软件。')
    if engine is None and not native:
        raise ValueError('未找到 LibreOffice。请安装后重新检测，或手动选择 soffice.exe。')
    if cancel and cancel.is_set():
        raise ConversionCancelled('已停止。')
    start = time.monotonic()
    def check_cancel():
        if cancel and cancel.is_set():
            raise ConversionCancelled('已停止当前 Office 转换。')
        if time.monotonic() - start > timeout:
            raise TimeoutError(f'转换超过 {timeout:g} 秒，已停止。')
    report = None
    with job_workspace() as folder:
        workspace = Path(folder)
        profile = session.profile
        local_input = workspace / ('source' + source.suffix.lower())
        rendered_objects = 0
        def render_object_impl(path, data):
            nonlocal rendered_objects
            check_cancel()
            rendered_objects += 1
            if rendered_objects > 6:
                raise ValueError('超过本文件 6 个外链对象的自动预览上限')
            object_dir = workspace / f'object-{rendered_objects}'
            object_dir.mkdir()
            copied = object_dir / ('resource' + path.suffix.lower())
            copied.write_bytes(data)
            if path.suffix.lower() in FILTERS:
                result = convert_office(copied, object_dir, engine, cancel,
                                        min(40, max(1, timeout - (time.monotonic()-start))), session,
                                        _depth=_depth+1, _origin=path)
                pdf_path = result.output_path
            elif path.suffix.lower() == '.pdf':
                pdf_path = copied
            else:
                raise ValueError('对象类型没有静态预览适配器')
            import pypdfium2 as pdfium
            from PIL import Image, ImageChops
            from io import BytesIO
            try:
                with pdfium.PdfDocument(pdf_path) as pdf:
                    page = pdf[0]
                    width, height = page.get_size()
                    bitmap = page.render(scale=min(2, 1800/max(width,height)))
                    try:
                        picture = bitmap.to_pil().convert('RGB')
                        white = Image.new('RGB', picture.size, 'white')
                        difference = ImageChops.difference(picture, white)
                        bounds = difference.getbbox()
                        if bounds:
                            left,top,right,bottom=bounds
                            cropped=picture.crop((max(0,left-12),max(0,top-12),min(picture.width,right+12),min(picture.height,bottom+12)))
                            picture.close()
                            picture=cropped
                        difference.close()
                        white.close()
                        stream = BytesIO()
                        picture.save(stream, format='PNG')
                        picture.close()
                        return stream.getvalue()
                    finally:
                        bitmap.close()
                        page.close()
            except (pdfium.PdfiumError, IndexError) as error:
                raise ValueError('对象预览无法渲染') from error
        def render_object(path, data):
            try:
                return render_object_impl(path, data)
            except ConversionCancelled:
                raise
            except Exception as error:
                raise ValueError(f'对象预览失败：{error}') from error
        if source.suffix.lower() in MARKERS:
            report = prepare_ooxml(source, local_input, render_object if _depth == 0 else None, check_cancel, _origin)
        else:
            shutil.copyfile(source, local_input)
        if source.suffix.lower() == '.xlsx' and _depth == 0:
            _fit_plain_workbook(local_input)
        check_cancel()
        output = workspace / 'out'
        output.mkdir()
        generated = output / 'source.pdf'
        if native:
            try:
                suite,kind=native
                if source.suffix.lower() == '.docx':
                    # Native Word can evaluate fields on open. Relationship
                    # preprocessing alone does not neutralize DDE/LINK fields.
                    with zipfile.ZipFile(local_input) as package:
                        instructions=[]
                        for name in package.namelist():
                            if name.startswith('word/') and name.endswith('.xml'):
                                root=ET.fromstring(package.read(name))
                                for node in root.iter():
                                    tag=node.tag.rsplit('}',1)[-1]
                                    if tag=='instrText':instructions.append(node.text or '')
                                    elif tag=='fldSimple':instructions.append(node.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}instr',''))
                        if re.search(r'\b(?:INCLUDEPICTURE|INCLUDETEXT|LINK|DDEAUTO|DDE|DATABASE)\b',''.join(instructions),re.I):
                            raise ValueError('文档含外部资源域，当前原生导出尚不能安全保留；请在办公软件中将其转为静态内容')
                exporter = export_wps if suite=='wps' else {'word':export_word,'sheet':export_excel,'slides':export_powerpoint}[kind]
                engine_name = {'word':'WPS 文字','sheet':'WPS 表格','slides':'WPS 演示'}[kind] if suite=='wps' else {'word':'Microsoft Word','sheet':'Microsoft Excel','slides':'Microsoft PowerPoint'}[kind]
                extra={'kind':kind} if suite=='wps' else {}
                exporter(local_input, generated, workspace, check_cancel,
                         timeout=max(1,timeout-(time.monotonic()-start)),**extra)
                if report is not None:
                    report.notes.append('排版基准：'+engine_name+'。请使用同一办公软件的打印版式核对；已有打印区域或缩放设置的工作表沿用原设置，未设置的纯数据表按页宽横向适配。')
                    recorded=document_suite(source)
                    if layout_suite=='auto' and recorded and recorded!=suite:
                        report.notes.append('原编辑器与本次引擎不同，字体、表格或分页可能有差异；可在排版基准中指定原编辑器。')
                target = publish_file(generated, output_dir, source.stem, '.pdf')
                return OfficeResult(target,target.stat().st_size,time.monotonic()-start,
                                    report.summary if report else '',
                                    tuple(report.notes) if report else (), engine_name)
            except ConversionCancelled:
                raise
            except (OSError,ValueError,TimeoutError) as error:
                check_cancel()
                if generated.exists():generated.unlink()
                raise ValueError(str(error)+'；为保留所选排版基准，本次未切换备用引擎。可检查对应办公软件后重试。') from error
        args = [str(engine), f'-env:UserInstallation={profile.as_uri()}', '--headless',
                '--nologo', '--nodefault', '--norestore', '--nofirststartwizard',
                '--convert-to', f'pdf:{FILTERS[source.suffix.lower()]}', '--outdir', str(output), str(local_input)]
        with (workspace / 'conversion.log').open('w+b') as log:
            # PyInstaller's DLL directory must not leak into the external engine.
            bundled = os.name == 'nt' and getattr(sys, 'frozen', False)
            if bundled:
                import ctypes
                ctypes.windll.kernel32.SetDllDirectoryW(None)
            try:
                process = subprocess.Popen(args, stdout=log, stderr=log, stdin=subprocess.DEVNULL,
                                           cwd=workspace, creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0)
            finally:
                if bundled:
                    ctypes.windll.kernel32.SetDllDirectoryW(sys._MEIPASS)
            try:
                while process.poll() is None:
                    if cancel and cancel.is_set():
                        raise ConversionCancelled('已停止当前 Office 转换。')
                    if time.monotonic() - start > timeout:
                        raise TimeoutError(f'转换超过 {timeout:g} 秒，已终止。请检查文档内容或缩小文件。')
                    time.sleep(0.1)
            finally:
                _stop_process(process)
            if cancel and cancel.is_set():
                raise ConversionCancelled('已停止，未保存输出。')
            generated = output / 'source.pdf'
            if process.returncode != 0 or not generated.is_file():
                raise ValueError('LibreOffice 未生成 PDF。请检查文档是否加密、损坏，或能否在 LibreOffice 中打开。')
            with generated.open('rb') as pdf:
                if pdf.read(5) != b'%PDF-':
                    raise ValueError('转换输出不是有效 PDF，未保存。')
            target = publish_file(generated, output_dir, source.stem, '.pdf')
    notes=list(report.notes) if report else []
    notes.append('本次为LibreOffice兼容输出；图片、字体、表格、线条及分页可能与WPS或Microsoft Office不同，请核对排版。')
    return OfficeResult(target, target.stat().st_size, time.monotonic() - start,
                        report.summary if report else '', tuple(notes))
