"""Explicit image-based fidelity mode. Never substitute for editable output."""
from io import BytesIO


def rebuild_visual_pdf(source, generated, target, pages, check_cancel, progress=None):
    import pypdfium2 as pdfium
    document = pdfium.PdfDocument(source)
    try:
        if target == 'docx':
            from docx import Document
            from docx.shared import Pt
            from docx.enum.section import WD_SECTION_START
            from docx.oxml import OxmlElement
            from docx.oxml.ns import qn
            output = Document()
            output._element.body.remove(output.paragraphs[0]._p) if output.paragraphs else None
        elif target == 'xlsx':
            from openpyxl import Workbook
            from openpyxl.drawing.image import Image
            output = Workbook()
            output.remove(output.active)
        else:
            from pptx import Presentation
            from pptx.util import Pt
            output = Presentation()
            first = document[0]
            width, height = first.get_size()
            first.close()
            factor = min(1.0, 4032 / max(width, height))
            slide_w, slide_h = width * factor, height * factor
            output.slide_width, output.slide_height = Pt(slide_w), Pt(slide_h)
        # Keep image streams alive through save (openpyxl reads them lazily).
        streams = []
        for index in range(pages):
            check_cancel()
            page = document[index]
            try:
                width, height = page.get_size()
                # 200 dpi for office print, bounded memory for large PDF pages.
                scale = min(200 / 72, (24_000_000 / max(1, width * height)) ** .5)
                bitmap = page.render(scale=scale)
                try:
                    stream = BytesIO()
                    picture = bitmap.to_pil()
                    try:
                        picture.save(stream, format='PNG')
                    finally:
                        picture.close()
                finally:
                    bitmap.close()
            finally:
                page.close()
            stream.seek(0)
            streams.append(stream)
            if target == 'docx':
                section = output.sections[0] if index == 0 else output.add_section(WD_SECTION_START.NEW_PAGE)
                factor = min(1.0, 1584 / max(width, height))
                placed_w, placed_h = width * factor, height * factor
                section.page_width, section.page_height = Pt(placed_w), Pt(placed_h)
                section.top_margin = section.bottom_margin = Pt(0)
                section.left_margin = section.right_margin = Pt(0)
                para = output.add_paragraph()
                para.paragraph_format.space_before = para.paragraph_format.space_after = Pt(0)
                para.paragraph_format.line_spacing = Pt(1)
                inline = para.add_run().add_picture(stream, width=Pt(placed_w), height=Pt(placed_h))._inline
                anchor = OxmlElement('wp:anchor')
                for key, value in {'distT':'0','distB':'0','distL':'0','distR':'0','simplePos':'0',
                                   'relativeHeight':'0','behindDoc':'0','locked':'0','layoutInCell':'0','allowOverlap':'1'}.items():
                    anchor.set(key, value)
                simple = OxmlElement('wp:simplePos'); simple.set('x','0'); simple.set('y','0'); anchor.append(simple)
                for axis in ('H','V'):
                    position = OxmlElement('wp:position' + axis); position.set('relativeFrom','page')
                    offset = OxmlElement('wp:posOffset'); offset.text='0'; position.append(offset); anchor.append(position)
                for tag in ('extent','effectExtent'):
                    node=inline.find(qn('wp:'+tag))
                    if node is not None: anchor.append(node)
                anchor.append(OxmlElement('wp:wrapNone'))
                for tag in ('docPr','cNvGraphicFramePr'):
                    node=inline.find(qn('wp:'+tag))
                    if node is not None: anchor.append(node)
                anchor.append(inline.find(qn('a:graphic')))
                inline.getparent().replace(inline,anchor)
            elif target == 'xlsx':
                sheet=output.create_sheet(f'第{index+1}页')
                sheet.sheet_view.showGridLines=False
                picture=Image(stream)
                factor=min(1.0, 792/max(width,height))
                placed_w,placed_h=width*factor,height*factor
                picture.width,picture.height=placed_w*96/72,placed_h*96/72
                sheet.add_image(picture,'A1')
                sheet.column_dimensions['A'].width=placed_w/5.25
                # Excel row height maximum is 409.5 pt.
                rows=max(1, int(placed_h/300)+1)
                for row in range(1,rows+1): sheet.row_dimensions[row].height=placed_h/rows
                sheet.print_area=f'A1:A{rows}'
                sheet.page_setup.paperSize=sheet.PAPERSIZE_A4
                sheet.page_setup.orientation='landscape' if placed_w>placed_h else 'portrait'
                sheet.page_setup.fitToWidth=sheet.page_setup.fitToHeight=1
                sheet.sheet_properties.pageSetUpPr.fitToPage=True
                from openpyxl.worksheet.page import PageMargins
                sheet.page_margins=PageMargins(left=0,right=0,top=0,bottom=0,header=0,footer=0)
                sheet['A1']=''
            else:
                slide=output.slides.add_slide(output.slide_layouts[6])
                fit=min(slide_w/width,slide_h/height)
                placed_w,placed_h=width*fit,height*fit
                slide.shapes.add_picture(stream,Pt((slide_w-placed_w)/2),Pt((slide_h-placed_h)/2),
                                         width=Pt(placed_w),height=Pt(placed_h))
            if progress: progress(index+1,pages)
        check_cancel()
        output.save(generated)
    finally:
        document.close()
