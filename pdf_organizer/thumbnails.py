"""Latest-viewport queue, bounded byte cache; never calls Tk or PDFium here."""
from collections import OrderedDict
import base64,json,os,queue,subprocess,sys,threading
from pathlib import Path


class ThumbnailService:
    def __init__(self):
        self.cache=OrderedDict();self.cache_bytes=0;self.max_cache=96;self.max_bytes=16*1024*1024
        self.condition=threading.Condition();self.pending=OrderedDict();self.results=queue.Queue();self.closed=False
        self.rendered=0;self.process=None
        self.outstanding=set();self.failed=None
        self.thread=threading.Thread(target=self._loop,daemon=True);self.thread.start()

    @staticmethod
    def request(ref,source,sheet):
        from dataclasses import asdict
        # Ship only this page's metadata, never copy/JSON-encode 500 boxes on Tk.
        info={'source_id':source.source_id,'path':source.path,'sha256':source.sha256,'byte_size':source.byte_size,
              'mtime_ns':source.mtime_ns,'boxes':[source.boxes[ref.page_index]]}
        body={'source':info,'ref':asdict(ref),'sheet':asdict(sheet)}
        key=json.dumps((source.path,source.sha256,source.mtime_ns,ref.page_index,ref.rotate,ref.crop,ref.fit_mode,ref.target_rect,sheet.size,sheet.margin),separators=(',',':'))
        return dict(body,key=key)

    def visible(self,requests):
        with self.condition:
            self.pending=OrderedDict((r['key'],r) for r in requests if r['key'] not in self.cache and r['key'] not in self.outstanding)
            if self.failed:
                for r in self.pending.values():self.results.put({'key':r['key'],'error':self.failed})
                self.pending.clear()
            self.condition.notify()

    def get(self,key):
        value=self.cache.get(key)
        if value is not None:self.cache.move_to_end(key)
        return value

    def receive(self,limit=3):
        results=[]
        for _ in range(limit):
            try:item=self.results.get_nowait()
            except queue.Empty:break
            key=item['key'];data=base64.b64decode(item['png']) if 'png' in item else b''
            with self.condition:self.outstanding.discard(key)
            if key in self.cache:self.cache_bytes-=len(self.cache.pop(key))
            self.cache[key]=data;self.cache_bytes+=len(data)
            while len(self.cache)>self.max_cache or self.cache_bytes>self.max_bytes:
                _,old=self.cache.popitem(last=False);self.cache_bytes-=len(old)
            self.rendered+=1;results.append(item)
        return results

    def _launch(self):
        if getattr(sys,'frozen',False):command=[sys.executable,'--pdf-thumbnail-worker']
        else:command=[sys.executable,'-B',str(Path(__file__).parents[1]/'app.py'),'--pdf-thumbnail-worker']
        env=os.environ.copy()
        if not getattr(sys,'frozen',False):env['PYTHONPATH']=os.pathsep.join(sys.path)
        return subprocess.Popen(command,stdin=subprocess.PIPE,stdout=subprocess.PIPE,stderr=subprocess.DEVNULL,
                                text=True,encoding='utf-8',bufsize=1,env=env,
                                creationflags=subprocess.CREATE_NO_WINDOW if os.name=='nt' else 0)

    def _loop(self):
        try:
            self.process=self._launch()
            while True:
                with self.condition:
                    self.condition.wait_for(lambda:self.closed or bool(self.pending))
                    if self.closed:break
                    key,request=self.pending.popitem(last=False)
                    self.outstanding.add(key)
                self.process.stdin.write(json.dumps(request)+'\n');self.process.stdin.flush()
                line=self.process.stdout.readline()
                if not line:raise RuntimeError('缩略图进程已退出')
                self.results.put(json.loads(line))
        except Exception as error:
            self.failed=str(error)
            with self.condition:pending=list(self.pending.values());self.pending.clear()
            if 'request' in locals():self.results.put({'key':request['key'],'error':str(error)})
            for request in pending:self.results.put({'key':request['key'],'error':str(error)})
        finally:
            if self.process is not None:
                try:
                    if self.process.poll() is None:
                        self.process.stdin.write('{"op":"close"}\n');self.process.stdin.flush();self.process.wait(timeout=.5)
                except (OSError,ValueError,subprocess.TimeoutExpired):
                    if self.process.poll() is None:self.process.terminate()
                for pipe in (self.process.stdin,self.process.stdout):
                    if pipe:pipe.close()

    def close(self):
        with self.condition:self.closed=True;self.pending.clear();self.condition.notify()
        # A pathological page cannot keep the desktop app alive indefinitely.
        def stop():
            self.thread.join(timeout=2)
            if self.thread.is_alive() and self.process and self.process.poll() is None:self.process.terminate()
        threading.Thread(target=stop,daemon=True).start()
