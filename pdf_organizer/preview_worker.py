"""Isolated composition/PDFium process. Sends only a bounded PNG to the Tk thread."""
from .thumb_worker import pipe_streams


def main():
    import base64,json
    from io import BytesIO
    import pypdfium2 as pdfium
    from .preview_layout import unpack
    from .exporter import build_writer
    inp,out=pipe_streams();cached=None;pdf=None
    try:
        for line in inp:
            r=json.loads(line)
            if r.get('op')=='close':break
            try:
                if cached!=r['revision']:
                    if pdf is not None:pdf.close();pdf=None
                    plan,sources=unpack(r)
                    writer=build_writer(plan,sources)
                    try:stream=BytesIO();writer.write(stream)
                    finally:writer.close()
                    pdf=pdfium.PdfDocument(stream.getvalue());cached=r['revision']
                page=pdf[r['page']]
                try:
                    w,h=page.get_size();bitmap=page.render(scale=min(2,1200/max(w,h)),draw_annots=True)
                    try:
                        im=bitmap.to_pil().convert('RGB');buf=BytesIO();im.save(buf,format='PNG');im.close()
                    finally:bitmap.close()
                finally:page.close()
                result={'key':r['key'],'png':base64.b64encode(buf.getvalue()).decode('ascii')}
            except Exception as error:result={'key':r['key'],'error':str(error)}
            out.write(json.dumps(result)+'\n');out.flush()
    finally:
        if pdf is not None:pdf.close()


if __name__=='__main__':main()
