"""Explicit local QA fixtures; no automatic clicks, conversion or desktop writes."""
import json,time
from pathlib import Path


def prepare_verification(root, app, job_file):
    job=json.loads(Path(job_file).read_text(encoding='utf-8-sig'))
    pane=next(p for p in app.panes if p.mode==job['mode'])
    pane.add_paths(job.get('files',[]))
    pane.output_var.set(job['output'])
    report=Path(job['report']);report.parent.mkdir(parents=True,exist_ok=True)
    app.show_mode(job['mode'])
    def record():
        rows=[{'id':row,'values':pane.tree.item(row,'values'),'selected':row in pane.tree.selection()} for row in pane.tree.get_children()]
        value={'time':time.time(),'busy':app.busy,'status':app.status.cget('text'),'rows':rows,'list_scroll':pane.tree.yview(),'page_scroll':pane.scroll.canvas.yview()}
        report.write_text(json.dumps(value,ensure_ascii=False,indent=2),encoding='utf-8')
        with report.with_suffix('.jsonl').open('a',encoding='utf-8') as stream:stream.write(json.dumps(value,ensure_ascii=False)+'\n')
        root.after(500,record)
    root.after(500,record)
