"""Local PDF to editable Office or explicitly selected page-image output."""
from dataclasses import dataclass
from pathlib import Path
import tempfile
import threading
import time

from file_output import publish_file
from office_engine import cleanup_workspace


@dataclass(frozen=True)
class PdfResult:
    output_path: Path
    output_bytes: int
    elapsed_seconds: float
    note: str
    pages: int


def _validate_pdf(source: Path):
    if source.stat().st_size > 200 * 1024 * 1024:
        raise ValueError('当前版本支持 200 MB 以内的 PDF。')
    with source.open('rb') as handle:
        if handle.read(5) != b'%PDF-':
            raise ValueError('文件内容不是有效 PDF。')
    import pypdfium2 as pdfium
    try:
        document = pdfium.PdfDocument(source)
        count = len(document)
        document.close()
    except Exception as error:
        raise ValueError('PDF 无法打开，可能已损坏或受密码保护。') from error
    if count < 1 or count > 500:
        raise ValueError('PDF 页数需在 1–500 页之间。')
    return count


def convert_pdf(source: Path, output_dir: Path, target: str,
                cancel: threading.Event | None = None, progress=None, *, layout_mode='editable') -> PdfResult:
    source = Path(source).resolve(strict=True)
    output_dir = Path(output_dir).resolve(strict=True)
    target = target.lower()
    if target not in {'docx', 'xlsx', 'pptx'}:
        raise ValueError('目标格式仅支持 Word、Excel 或 PowerPoint。')
    if layout_mode not in {'editable', 'fidelity', 'visual'}:
        raise ValueError('请选择有效的转换设置。')
    if target != 'docx' and layout_mode == 'fidelity':
        raise ValueError('自动还原为 Word 专用转换设置。')
    pages = _validate_pdf(source)
    started = time.monotonic()

    def check_cancel():
        if cancel and cancel.is_set():
            from office_engine import ConversionCancelled
            raise ConversionCancelled('已停止当前 PDF 转换。')

    workspace = Path(tempfile.mkdtemp(prefix='ruyi-pdf-'))
    try:
        generated = workspace / f'converted.{target}'
        if layout_mode == 'visual' and target != 'docx':
            from pdf_faithful import rebuild_visual_pdf
            rebuild_visual_pdf(source, generated, target, pages, check_cancel, progress)
            note = '原版外观模式：每页为高清图片，保留原稿外观；图片内文字、表格不可直接编辑。'
        elif target == 'docx':
            from pdf_docx_fidelity import rebuild_pdf_docx_fidelity
            layout = rebuild_pdf_docx_fidelity(source, generated, pages, check_cancel, progress)
            note = (f'自动还原：保留 {pages} 页原稿的页面尺寸、文字行与段落位置；'
                    f'重建 {layout.paragraphs} 个可编辑段落、{layout.tables} 个可编辑表格（{layout.cells} 个单元格），保留 '
                    f'{layout.formula_regions} 处公式混排区域及 {layout.images} 张图片。\n'
                    '普通正文、表题及重建表格可编辑；复杂公式以局部图片保留在原段落或单元格中。扫描内容不含OCR识别。')
            if layout.directory_entries:
                note += f'已恢复 {layout.directory_entries} 个可编辑目录条目，页码保留原PDF文字。'
            if layout.toc_entries:
                note += '源文件带结构标签的目录已建立Word目录域，更新前请核对标题关联与分页。'
            if layout.warnings: note += '\n'+'\n'.join(layout.warnings[:5])
            if layout.scanned_pages:
                note += f'其中 {layout.scanned_pages} 个扫描页按整页图片保留，不执行 OCR。'
        elif target == 'xlsx':
            from pdf_xlsx_layout import rebuild_pdf_xlsx
            layout = rebuild_pdf_xlsx(source, generated, pages, check_cancel, progress)
            note = f'已重建 {layout.tables} 个表格、{layout.cells} 个单元格、保留 {layout.images} 张图片；Excel 字距与自由版式可能有差异。'
            if layout.scanned_pages: note += f'{layout.scanned_pages} 个扫描页保留图像，不执行 OCR。'
            if layout.warnings: note += '；'.join(layout.warnings[:3])
        else:
            from pdf_pptx_layout import rebuild_pdf_pptx
            layout = rebuild_pdf_pptx(source, generated, pages, check_cancel, progress)
            note = (f'已保留 {layout.text_boxes} 个可编辑文本框、{layout.tables} 个可编辑表格（{layout.cells} 个单元格）、'
                    f'{layout.images} 张可独立移动的图片和 {layout.shapes} 个线条图形；保留源叠放顺序和透明度。'
                    'PDF不包含原PPT的动画或图表数据，矢量图表保留为可编辑图形；复杂版式和缺失字体可能有差异。')
            if layout.scanned_pages: note += f'{layout.scanned_pages} 个扫描页保留图像，不执行 OCR。'
            if layout.raster_fragments: note += f'{layout.raster_fragments} 个复杂局部按图片保留。'
        check_cancel()
        if not generated.is_file() or generated.stat().st_size < 100:
            raise ValueError('转换引擎未生成有效文件。')
        from office_preview import remove_office_thumbnail
        remove_office_thumbnail(generated)
        check_cancel()
        output = publish_file(generated, output_dir, source.stem, '.' + target)
        return PdfResult(output, output.stat().st_size, time.monotonic() - started, note, pages)
    finally:
        cleanup_workspace(workspace)


def _pdf_to_docx(source, generated, pages, check_cancel, progress):
    try:
        from pdf_docx_layout import rebuild_pdf_docx
        return rebuild_pdf_docx(source, generated, pages, check_cancel, progress)
    except ImportError as error:
        raise ValueError('Word 转换组件未安装，请重装如意。') from error

