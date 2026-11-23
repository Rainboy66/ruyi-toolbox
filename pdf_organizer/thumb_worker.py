"""Single PDFium process, separate from all desktop/conversion threads."""
from collections import OrderedDict
from io import BytesIO
import base64,json,sys


def pipe_streams():
    """pythonw / PyInstaller --windowed set stdio to None despite inherited pipes."""
    if sys.stdin is not None and sys.stdout is not None:return sys.stdin,sys.stdout
    import ctypes,io,msvcrt,os
    kernel=ctypes.WinDLL('kernel32',use_last_error=True)
    kernel.GetStdHandle.argtypes=[ctypes.c_uint32];kernel.GetStdHandle.restype=ctypes.c_void_p
    streams=[]
    for code,mode,flags in ((-10,'r',os.O_RDONLY),(-11,'w',os.O_WRONLY)):
        handle=kernel.GetStdHandle(code & 0xffffffff)
        if not handle or handle==ctypes.c_void_p(-1).value:raise OSError('缩略图进程缺少通信管道。')
        fd=msvcrt.open_osfhandle(handle,flags|os.O_BINARY)
        streams.append(io.open(fd,mode,encoding='utf-8',buffering=1))
    return tuple(streams)


def main():
    input_stream,output_stream=pipe_streams()
    import pypdfium2 as pdfium
    from PIL import Image
    from .model import SourceInfo,PageRef,Sheet,page_geometry
    documents=OrderedDict()
    try:
        for line in input_stream:
            request=json.loads(line)
            if request.get('op')=='close':break
            try:
                source=SourceInfo(**request['source']);ref=PageRef(**request['ref']);sheet=Sheet(**request['sheet'])
                from dataclasses import replace
                original_index=ref.page_index;ref=replace(ref,page_index=0)
                from pathlib import Path
                stat=Path(source.path).stat()
                if (stat.st_size,stat.st_mtime_ns)!=(source.byte_size,source.mtime_ns):raise ValueError('源文件已变化')
                if source.path not in documents:
                    doc=pdfium.PdfDocument(source.path);doc.init_forms();documents[source.path]=doc
                documents.move_to_end(source.path)
                while len(documents)>3:documents.popitem(last=False)[1].close()
                page=documents[source.path][original_index]
                try:
                    w,h=page.get_size();bitmap=page.render(scale=min(2,max(240/w,280/h)),draw_annots=True)
                    try:im=bitmap.to_pil().convert('RGB')
                    finally:bitmap.close()
                finally:page.close()
                if ref.crop is not None:
                    l,b,r,t,native,unit=source.boxes[ref.page_index]
                    matrices={0:(1,0,0,1,-l,-b),90:(0,-1,1,0,-b,r),180:(-1,0,0,-1,r,t),270:(0,1,-1,0,t,-l)}
                    aa,bb,c,d,e,f=matrices[native]
                    cl,cb,cr,ct=ref.crop
                    pts=[(aa*x+c*y+e,bb*x+d*y+f) for x,y in ((cl,cb),(cl,ct),(cr,cb),(cr,ct))]
                    vw,vh=(t-b,r-l) if native%180 else (r-l,t-b)
                    box=(round(min(x for x,y in pts)*im.width/vw),round((vh-max(y for x,y in pts))*im.height/vh),
                         round(max(x for x,y in pts)*im.width/vw),round((vh-min(y for x,y in pts))*im.height/vh))
                    cropped=im.crop(box);im.close();im=cropped
                if ref.rotate:
                    rotated=im.transpose({90:Image.Transpose.ROTATE_270,180:Image.Transpose.ROTATE_180,270:Image.Transpose.ROTATE_90}[ref.rotate]);im.close();im=rotated
                size,target,_=page_geometry(ref,source,sheet)
                ratio=min(180/size[0],220/size[1]);pw,ph=max(1,round(size[0]*ratio)),max(1,round(size[1]*ratio))
                result=Image.new('RGB',(pw,ph),'white');x,y,x2,y2=target
                tx,ty=round(x/size[0]*pw),round((size[1]-y2)/size[1]*ph)
                tw,th=max(1,round((x2-x)/size[0]*pw)),max(1,round((y2-y)/size[1]*ph))
                sx,sy=tw/im.width,th/im.height
                if ref.fit_mode!='stretch':sx=sy=(min if ref.fit_mode=='fit' else max)(sx,sy)
                resized=im.resize((max(1,round(im.width*sx)),max(1,round(im.height*sy))),Image.Resampling.LANCZOS);im.close()
                slot=Image.new('RGB',(tw,th),'white');slot.paste(resized,((tw-resized.width)//2,(th-resized.height)//2));resized.close()
                result.paste(slot,(tx,ty));slot.close();data=BytesIO();result.save(data,format='PNG');result.close()
                response={'key':request['key'],'png':base64.b64encode(data.getvalue()).decode('ascii')}
            except Exception as error:response={'key':request['key'],'error':str(error)}
            output_stream.write(json.dumps(response)+'\n');output_stream.flush()
    finally:
        for doc in documents.values():doc.close()


if __name__=='__main__':main()
