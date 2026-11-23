"""One replaceable pending request; no Tk or PDFium in the UI thread."""
import json,os,queue,subprocess,sys,threading
from pathlib import Path


class PreviewService:
    def __init__(self):
        self.condition=threading.Condition();self.pending=None;self.closed=False;self.process=None;self.results=queue.Queue()
        self.thread=threading.Thread(target=self.run,daemon=True);self.thread.start()

    def request(self,data):
        with self.condition:self.pending=data;self.condition.notify()

    def run(self):
        try:
            with self.condition:
                self.condition.wait_for(lambda:self.closed or self.pending is not None)
                if self.closed:return
            command=[sys.executable,'--pdf-preview-worker'] if getattr(sys,'frozen',False) else [sys.executable,'-B',str(Path(__file__).parents[1]/'app.py'),'--pdf-preview-worker']
            env=os.environ.copy()
            if not getattr(sys,'frozen',False):env['PYTHONPATH']=os.pathsep.join(sys.path)
            self.process=subprocess.Popen(command,stdin=subprocess.PIPE,stdout=subprocess.PIPE,stderr=subprocess.DEVNULL,text=True,encoding='utf-8',bufsize=1,env=env,creationflags=subprocess.CREATE_NO_WINDOW if os.name=='nt' else 0)
            while True:
                with self.condition:
                    self.condition.wait_for(lambda:self.closed or self.pending is not None)
                    if self.closed:break
                    r=self.pending;self.pending=None
                from dataclasses import asdict
                r=dict(r);r['plan']=asdict(r.pop('_plan'));r['sources']={key:asdict(source) for key,source in r.pop('_sources').items()}
                self.process.stdin.write(json.dumps(r)+'\n');self.process.stdin.flush()
                line=self.process.stdout.readline()
                if not line:raise RuntimeError('预览进程已退出')
                self.results.put(json.loads(line))
        except Exception as error:
            self.results.put({'key':r['key'] if 'r' in locals() else None,'error':str(error)})
        finally:
            if self.process is not None:
                try:
                    if self.process.poll() is None:self.process.terminate()
                    self.process.wait(timeout=3)
                except (OSError,subprocess.TimeoutExpired):pass
                for pipe in (self.process.stdin,self.process.stdout):
                    if pipe is not None:pipe.close()

    def close(self):
        with self.condition:self.closed=True;self.pending=None;self.condition.notify()
        if self.process is not None and self.process.poll() is None:
            try:self.process.terminate()
            except OSError:pass
