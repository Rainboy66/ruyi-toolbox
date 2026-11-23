"""Presets are source-free OutputPlans. Batch files each get the same rules."""
from dataclasses import replace
from pathlib import Path
import json,tempfile
from .model import PlanDocument,OutputPlan,inspect_sources,verify_sources,sha256


def save_preset(plan,path):
    PlanDocument(replace(plan,pages=()),{}).save(path)


def load_preset(path):
    doc=PlanDocument.load(path)
    if doc.plan.pages or doc.sources:raise ValueError('请选择不含源页引用的规则预设。')
    from .overlays import verify_assets
    verify_assets(doc.plan)
    return doc.plan


def apply_preset(plan,preset):
    return replace(plan,sheet=preset.sheet,rules=preset.rules,overlays=preset.overlays)


def batch_apply(paths,preset,directory,progress=None,cancel=None):
    from .exporter import export_pdf
    directory=Path(directory)
    if directory.exists():raise ValueError('批量结果目录已存在，请选择新目录。')
    directory.mkdir(parents=True);results=[]
    paths=tuple(paths)
    for i,path in enumerate(paths):
        if cancel and cancel.is_set():break
        item={'source':str(path)}
        try:
            doc=PlanDocument();doc.import_sources(inspect_sources([path]));doc.plan=apply_preset(doc.plan,preset)
            name=f'{i+1:03}_{Path(path).stem}_拼版.pdf'
            out=export_pdf(doc.plan,doc.sources,directory/name,cancel=cancel)
            item.update(ok=True,output=str(out),sha256=sha256(out))
        except InterruptedError:break
        except Exception as error:item.update(ok=False,error=str(error))
        results.append(item)
        if progress:progress(i+1,len(paths))
    (directory/'批量结果.json').write_text(json.dumps(results,ensure_ascii=False,indent=2),encoding='utf-8')
    return results


def split_composed(plan,sources,directory,count,progress,cancel):
    from .exporter import export_pdf,split_pdf
    with tempfile.TemporaryDirectory(prefix='ruyi-compose-') as tmp:
        combined=Path(tmp)/'combined.pdf';export_pdf(plan,sources,combined,cancel=cancel)
        doc=PlanDocument();doc.import_sources(inspect_sources([combined]))
        outputs=split_pdf(doc.plan,doc.sources,directory,count,progress,cancel)
    verify_sources(plan,sources);return outputs
