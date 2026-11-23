"""Prepare an OOXML copy: local resources become static, missing ones get placeholders."""
from copy import deepcopy
from dataclasses import dataclass, field
from io import BytesIO
import os
import posixpath
import re
from pathlib import Path
from urllib.parse import unquote, urlsplit
import zipfile

from lxml import etree as X
from PIL import Image, ImageDraw, ImageFont, ImageOps

R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
V = 'urn:schemas-microsoft-com:vml'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
MAX_RESOURCE = 25 * 1024 * 1024
MAX_TOTAL = 150 * 1024 * 1024


def xml(data):
    root = X.fromstring(data, X.XMLParser(resolve_entities=False, load_dtd=False, no_network=True))
    if root.getroottree().docinfo.doctype:
        raise ValueError('文档 XML 含 DTD，暂不处理。')
    return root


def packed(root):
    return X.tostring(root, encoding='UTF-8', xml_declaration=True, standalone=True)


def local_name(node):
    return X.QName(node).localname if isinstance(node.tag, str) else ''


def local_resource_path(target, source):
    """Resolve against the original document, never against its temporary copy."""
    target = target.strip()
    if not target or '\x00' in target:
        raise ValueError('空路径或无效路径')
    if re.match(r'^[a-zA-Z]:[\\/]', target):
        value = target
    else:
        parsed = urlsplit(target)
        if parsed.scheme:
            if parsed.scheme.lower() != 'file' or parsed.netloc.lower() not in ('', 'localhost'):
                raise ValueError('网络或其他协议资源未自动读取')
            value = unquote(parsed.path)
            if re.match(r'^/[a-zA-Z]:[/\\]', value):
                value = value[1:]
        else:
            value = target
    if value.startswith(('\\\\', '//')):
        raise ValueError('网络共享或设备路径未自动读取')
    value = value.replace('\\', os.sep)
    # Reject drive-relative paths and NTFS alternate streams, but allow normal drive roots.
    if ':' in re.sub(r'^[a-zA-Z]:[/\\]', '', value):
        raise ValueError('不支持的路径形式')
    path = Path(value)
    if not path.is_absolute():
        path = source.parent / path
    if '%' in value and not path.is_file():
        path = Path(unquote(value))
        if not path.is_absolute():
            path = source.parent / path
    path = path.resolve()
    if str(path).startswith(('\\\\', '//')):
        raise ValueError('网络共享未自动读取')
    return path


def placeholder():
    image = Image.new('RGB', (960, 540), '#f1f5f9')
    draw = ImageDraw.Draw(image)
    draw.rectangle((3, 3, 956, 536), outline='#94a3b8', width=6)
    font_path = Path(os.environ.get('WINDIR', 'C:/Windows'))/'Fonts/msyh.ttc'
    if font_path.is_file():
        font = ImageFont.truetype(str(font_path), 38)
        text = '外部资源缺失或不可读取'
    else:
        font = ImageFont.load_default(size=32)
        text = 'External resource unavailable'
    draw.text((480, 240), text, font=font, fill='#475569', anchor='mm')
    draw.text((480, 315), 'RUYI - RESOURCE UNAVAILABLE', font=ImageFont.load_default(size=25), fill='#64748b', anchor='mm')
    out = BytesIO()
    image.save(out, format='PNG')
    image.close()
    return out.getvalue()


def image_bytes(data):
    with Image.open(BytesIO(data)) as image:
        if image.width * image.height > 50_000_000:
            raise ValueError('图片超过 5000 万像素')
        result = ImageOps.exif_transpose(image).convert('RGBA')
        output = BytesIO()
        result.save(output, format='PNG')
        result.close()
        return output.getvalue()


def fit_object_image(owner, rid, data):
    node=next((n for n in owner.iter() if local_name(n) in ('oleObj','OLEObject','oleObject') and rid in n.attrib.values()),None)
    frame=next((n for n in node.iterancestors() if local_name(n)=='graphicFrame'),None) if node is not None else None
    extent=next((n for n in frame.iter() if local_name(n)=='ext' and n.get('cx') and n.get('cy')),None) if frame is not None else None
    if extent is None:
        return data
    try:
        ratio=int(extent.get('cx'))/int(extent.get('cy'))
        width,height=(1600,max(1,round(1600/ratio))) if ratio>=1 else (max(1,round(1600*ratio)),1600)
        with Image.open(BytesIO(data)) as image:
            fitted=ImageOps.contain(image,(width,height),Image.Resampling.LANCZOS)
            canvas=Image.new('RGB',(width,height),'white')
            canvas.paste(fitted,((width-fitted.width)//2,(height-fitted.height)//2))
            output=BytesIO();canvas.save(output,format='PNG')
            fitted.close();canvas.close()
            return output.getvalue()
    except (ValueError,ZeroDivisionError):
        return data


@dataclass
class ResourceReport:
    loaded: int = 0
    degraded: int = 0
    notes: list[str] = field(default_factory=list)
    cached: int = 0

    @property
    def summary(self):
        if not self.loaded and not self.degraded:
            return ''
        if self.cached:
            return f'外链加载 {self.loaded} 项，保留缓存 {self.cached} 项，其他降级 {self.degraded-self.cached} 项（双击查看详情）'
        return f'外链加载 {self.loaded} 项，降级 {self.degraded} 项（双击查看详情）'


def owner_part(rels_name):
    if rels_name == '_rels/.rels':
        return None
    directory, name = posixpath.split(rels_name)
    return posixpath.join(posixpath.dirname(directory), name[:-5])


def replace_object(owner, rid, image_rid, preserve_preview=False):
    """Replace a linked OLE object's visible shape, preserving its position and extent."""
    matched = [node for node in owner.iter() if local_name(node) in ('oleObj', 'OLEObject', 'oleObject')
               and any(v == rid for k, v in node.attrib.items() if k.endswith('}id'))]
    for node in matched:
        frame = next((p for p in node.iterancestors() if local_name(p) == 'graphicFrame'), None)
        if frame is not None and frame.tag == f'{{{P}}}graphicFrame':
            pic = next((p for p in node.iter() if p.tag == f'{{{P}}}pic'), None)
            if pic is not None:
                pic = deepcopy(pic)
            else:
                pic = X.Element(f'{{{P}}}pic')
                nv = X.SubElement(pic, f'{{{P}}}nvPicPr')
                original = frame.find(f'{{{P}}}nvGraphicFramePr/{{{P}}}cNvPr')
                X.SubElement(nv, f'{{{P}}}cNvPr', id=original.get('id', '1') if original is not None else '1', name='如意外部资源')
                X.SubElement(nv, f'{{{P}}}cNvPicPr')
                X.SubElement(nv, f'{{{P}}}nvPr')
                fill = X.SubElement(pic, f'{{{P}}}blipFill')
                X.SubElement(fill, f'{{{A}}}blip')
                X.SubElement(X.SubElement(fill, f'{{{A}}}stretch'), f'{{{A}}}fillRect')
                X.SubElement(pic, f'{{{P}}}spPr')
            for blip in pic.iter(f'{{{A}}}blip'):
                if not preserve_preview or not blip.get(f'{{{R}}}embed'):
                    blip.attrib.pop(f'{{{R}}}link', None)
                    blip.set(f'{{{R}}}embed', image_rid)
            props = pic.find(f'{{{P}}}spPr')
            if props is None:
                props = X.SubElement(pic, f'{{{P}}}spPr')
            transform = frame.find(f'{{{P}}}xfrm')
            if transform is not None:
                for previous in list(props):
                    if local_name(previous) == 'xfrm':
                        props.remove(previous)
                transform = deepcopy(transform)
                transform.tag = f'{{{A}}}xfrm'
                props.insert(0, transform)
            frame.getparent().replace(frame, pic)
        elif node.tag.endswith('}OLEObject'):
            container = node.getparent()
            shapes = [p for p in container.iter() if p.tag == f'{{{V}}}shape']
            if not shapes:
                shapes = [X.SubElement(container, f'{{{V}}}shape', style='width:240pt;height:135pt')]
            for shape in shapes:
                imgs = list(shape.iter(f'{{{V}}}imagedata'))
                if not imgs:
                    imgs = [X.SubElement(shape, f'{{{V}}}imagedata')]
                for img in imgs:
                    if not preserve_preview or not img.get(f'{{{R}}}id'):
                        img.attrib.clear()
                        img.set(f'{{{R}}}id', image_rid)
            container.remove(node)
        else:
            # Spreadsheet OLE previews live in a separate drawing/VML part.
            node.getparent().remove(node)
    return bool(matched)


def freeze_chart_cache(owner):
    """Retain saved chart values without an external workbook or recalculation.

    Work on a copy first: an incomplete cache must not partially rewrite a chart.
    Numeric and string references become literals; styles and geometry stay intact.
    """
    chart_ns = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
    draft = deepcopy(owner)
    series = list(draft.iter(f'{{{chart_ns}}}ser'))
    if not series:
        return False
    for series_node in series:
        values = [n for n in series_node if local_name(n) in ('val', 'yVal')]
        if not values or not any(n.tag == f'{{{chart_ns}}}v' and n.text for value in values for n in value.iter()):
            return False
    for ref in list(draft.iter()):
        kind = local_name(ref)
        if ref.tag not in (f'{{{chart_ns}}}numRef', f'{{{chart_ns}}}strRef'):
            if kind == 'multiLvlStrRef':
                return False  # Keep an explicit fallback until multilevel literals are supported.
            continue
        cache = ref.find(f'{{{chart_ns}}}{"numCache" if kind == "numRef" else "strCache"}')
        if cache is None or not any(local_name(n) == 'v' and n.text is not None for n in cache.iter()):
            return False
        if local_name(ref.getparent()) == 'tx':
            value = next(n for n in cache.iter() if local_name(n) == 'v')
            ref.getparent().replace(ref, deepcopy(value))
            continue
        literal = deepcopy(cache)
        literal.tag = f'{{{chart_ns}}}{"numLit" if kind == "numRef" else "strLit"}'
        ref.getparent().replace(ref, literal)
    def series_format(series_node, point_index=None):
        # A linked label's format comes from the values, not the placeholder
        # formatCode="General" on c:dLbl/c:numFmt. Keep the stored number text
        # unchanged; only make its effective display format explicit.
        value = next((n for n in series_node if local_name(n) in ('val', 'yVal')), None)
        if value is None:
            return None
        cache = next((n for n in value.iter() if n.tag in
                      (f'{{{chart_ns}}}numCache', f'{{{chart_ns}}}numLit')), None)
        if cache is None:
            return None
        if point_index is not None:
            point = next((n for n in cache if local_name(n) == 'pt' and n.get('idx') == point_index), None)
            if point is not None:
                code = point.find(f'{{{chart_ns}}}formatCode')
                if code is not None and code.text:
                    return code.text
                if point.get('formatCode'):
                    return point.get('formatCode')
        code = cache.find(f'{{{chart_ns}}}formatCode')
        return code.text if code is not None and code.text else None

    for fmt in draft.iter(f'{{{chart_ns}}}numFmt'):
        parent = fmt.getparent()
        linked = fmt.get('sourceLinked', '1').lower() in ('1', 'true', 'on')
        if linked and local_name(parent) in ('dLbl', 'dLbls'):
            series_node = next((n for n in parent.iterancestors() if n.tag == f'{{{chart_ns}}}ser'), None)
            index = parent.find(f'{{{chart_ns}}}idx') if local_name(parent) == 'dLbl' else None
            point_index = index.get('val') if index is not None else None
            if series_node is not None:
                code = series_format(series_node, point_index)
            else:
                # Plot-level labels can inherit different formats per series.
                # Resolve only a common format; otherwise keep their link to
                # the now-local literals instead of inventing one format.
                label_group = parent if local_name(parent) == 'dLbls' else parent.getparent()
                plot = label_group.getparent()
                codes = {series_format(n, point_index) for n in plot if n.tag == f'{{{chart_ns}}}ser'}
                code = next(iter(codes)) if len(codes) == 1 else None
            if code:
                fmt.set('formatCode', code)
            else:
                continue
        fmt.set('sourceLinked', '0')
    owner.attrib.clear()
    owner.attrib.update(draft.attrib)
    owner[:] = list(draft)
    return True


def prepare_ooxml(source: Path, destination: Path, render_object=None, check_cancel=None, origin=None) -> ResourceReport:
    report = ResourceReport()
    with zipfile.ZipFile(source) as archive:
        parts = {item.filename: archive.read(item) for item in archive.infolist()}
    roots = {}
    def root(name):
        if name not in roots:
            roots[name] = xml(parts[name])
        return roots[name]
    types = root('[Content_Types].xml')
    remaining = MAX_TOTAL
    count = 0
    sequence = 0
    placeholder_data = None
    def missing():
        nonlocal placeholder_data
        if placeholder_data is None:
            placeholder_data = placeholder()
        return placeholder_data
    def resource(target):
        nonlocal remaining, count
        count += 1
        if count > 200:
            raise ValueError('外链资源数量超过 200 项')
        path = local_resource_path(target, origin or source)
        if not path.is_file():
            raise ValueError('文件不存在')
        size = path.stat().st_size
        if size > MAX_RESOURCE or size > remaining:
            raise ValueError('资源超出读取大小上限')
        with path.open('rb') as stream:
            data = stream.read(min(MAX_RESOURCE, remaining) + 1)
        if len(data) > min(MAX_RESOURCE, remaining):
            raise ValueError('资源超出读取大小上限')
        remaining -= len(data)
        return path, data
    def image_relation(rels, owner_name, data):
        nonlocal sequence
        sequence += 1
        name = f'ruyi_resources/image{sequence}.png'
        while name in parts:
            sequence += 1
            name = f'ruyi_resources/image{sequence}.png'
        parts[name] = data
        X.SubElement(types, f'{{{CT}}}Override', PartName='/'+name, ContentType='image/png')
        rid = f'ruyiImage{sequence}'
        ids = {r.get('Id') for r in rels}
        while rid in ids:
            rid += '_'
        X.SubElement(rels, f'{{{REL}}}Relationship', Id=rid, Type=R+'/image',
                     Target=posixpath.relpath(name, posixpath.dirname(owner_name)))
        return rid
    for rel_name in [n for n in parts if n.endswith('.rels')]:
        rels = root(rel_name)
        owner_name = owner_part(rel_name)
        owner = root(owner_name) if owner_name in parts and owner_name.endswith(('.xml', '.vml')) else None
        externals = [r for r in rels if r.get('TargetMode', '').lower() == 'external'
                     and not r.get('Type', '').endswith('/hyperlink')]
        externals.sort(key=lambda r: not r.get('Type', '').endswith('/image'))
        for relation in externals:
            if check_cancel:
                check_cancel()
            rid, target = relation.get('Id'), relation.get('Target', '')
            kind = relation.get('Type', '').rsplit('/', 1)[-1]
            label = f'{owner_name or "文档"} / {rid}: {target}'
            if owner is None:
                rels.remove(relation)
                report.degraded += 1
                report.notes.append(label+' → 移除无显示内容的外部关系')
                continue
            if kind == 'image':
                try:
                    _, data = resource(target)
                    data = image_bytes(data)
                    report.loaded += 1
                    report.notes.append(label+' → 已加载本地图片')
                except (OSError, ValueError, Image.DecompressionBombError) as error:
                    data = missing()
                    report.degraded += 1
                    report.notes.append(label+f' → 占位图（{error}）')
                new_id = image_relation(rels, owner_name, data)
                for node in owner.iter():
                    for key, value in list(node.attrib.items()):
                        if value == rid and key.startswith('{'):
                            if local_name(node) in ('blip', 'svgBlip') and key.endswith('}link'):
                                node.attrib.pop(key)
                                node.set(key.rsplit('}', 1)[0]+'}embed', new_id)
                            else:
                                node.set(key, new_id)
                rels.remove(relation)
            elif kind in ('oleObject', 'package') and any(local_name(n) in ('oleObj','OLEObject','oleObject') and rid in n.attrib.values() for n in owner.iter()):
                readable = False
                try:
                    path, data = resource(target)
                    readable = True
                    if render_object is None:
                        raise ValueError('嵌套对象使用静态预览')
                    data = render_object(path, data)
                    report.loaded += 1
                    report.notes.append(label+' → 已读取，使用首个打印页的静态预览')
                except (OSError, ValueError, TimeoutError) as error:
                    data = missing()
                    report.degraded += 1
                    report.notes.append(label+f' → {"优先使用已有预览，否则占位图" if readable else "占位图"}（{error}）')
                preserve = readable and data == missing()
                data = fit_object_image(owner, rid, data)
                new_id = image_relation(rels, owner_name, data)
                # Excel stores an object preview in a separate drawing and/or VML part.
                if owner_name.startswith('xl/'):
                    _replace_excel_preview(parts, roots, types, owner_name, owner, rid, data, image_relation, preserve=preserve)
                replace_object(owner, rid, new_id, preserve_preview=preserve)
                rels.remove(relation)
            elif kind in ('package', 'oleObject') and any(local_name(n)=='externalData' and rid in n.attrib.values() for n in owner.iter()):
                # A chart already stores its plotted series. Embed readable source workbooks,
                # but keep the saved series instead of promising an unverified recalculation.
                try:
                    path, data = resource(target)
                    if path.suffix.lower() != '.xlsx':
                        raise ValueError('图表数据源暂只内嵌 XLSX')
                    with zipfile.ZipFile(BytesIO(data)) as book:
                        if len(book.infolist())>10000 or sum(i.file_size for i in book.infolist())>500*1024*1024:
                            raise ValueError('图表数据源解压后过大')
                        if 'xl/workbook.xml' not in book.namelist() or any(n.lower().endswith('vbaproject.bin') for n in book.namelist()):
                            raise ValueError('数据源不是无宏 XLSX')
                    sequence += 1
                    name=f'ruyi_resources/workbook{sequence}.xlsx'
                    while name in parts:
                        sequence += 1
                        name=f'ruyi_resources/workbook{sequence}.xlsx'
                    parts[name]=data
                    X.SubElement(types,f'{{{CT}}}Override',PartName='/'+name,ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                    relation.attrib.pop('TargetMode',None)
                    relation.set('Target',posixpath.relpath(name,posixpath.dirname(owner_name)))
                    for node in owner.iter():
                        if local_name(node)=='externalData' and rid in node.attrib.values():
                            update=next((x for x in node if local_name(x)=='autoUpdate'),None)
                            if update is None:
                                update=X.SubElement(node,node.tag.rsplit('}',1)[0]+'}autoUpdate')
                            update.set('val','0')
                    report.loaded += 1
                    report.notes.append(label+' → 已内嵌数据源；图表显示保存时的缓存，不强制重算')
                except (OSError,ValueError,zipfile.BadZipFile) as error:
                    for node in list(owner.iter()):
                        if local_name(node)=='externalData' and rid in node.attrib.values():
                            node.getparent().remove(node)
                    rels.remove(relation)
                    report.degraded += 1
                    if freeze_chart_cache(owner):
                        report.cached += 1
                        report.notes.append(label+f' → 已保留原图表和保存时的数据缓存；未更新数据源（{error}）')
                    else:
                        _replace_chart(parts, roots, owner_name, missing(), image_relation)
                        report.notes.append(label+f' → 无可用完整图表缓存，使用占位图（{error}）')
            else:
                # Charts retain their embedded series cache; workbook formulas retain cached values.
                # Never leave an external fetch request for LibreOffice to follow implicitly.
                for node in list(owner.iter()):
                    if node is owner:
                        continue
                    if rid in node.attrib.values():
                        if local_name(node) == 'externalData':
                            node.getparent().remove(node)
                        else:
                            for key, value in list(node.attrib.items()):
                                if value == rid:
                                    node.attrib.pop(key)
                rels.remove(relation)
                report.degraded += 1
                report.notes.append(label+' → 保留文档已有缓存；不自动更新外部数据')
    # The native spreadsheet opener's UpdateLinks=0 does not disable data
    # connections or query/pivot refresh-on-open. Keep their saved contents.
    refresh_count = _disable_external_refresh(parts, roots)
    if refresh_count:
        report.notes.append(f'已禁用 {refresh_count} 处连接、查询或透视缓存自动刷新，使用文件保存时的数据')
    # Freeze external workbook formulas to stored results, including shared-formula followers.
    if any('/externalLinks/' in n for n in parts):
        _freeze_external_formulas(parts, roots)
    for name, value in roots.items():
        parts[name] = packed(value)
    with zipfile.ZipFile(destination, 'w', zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    return report


def _disable_external_refresh(parts, roots):
    """Disable refresh controls without touching cells, formats or cached data."""
    count = 0
    for name in parts:
        if not (name == 'xl/connections.xml' or
                (name.startswith('xl/queryTables/') and name.endswith('.xml')) or
                (name.startswith('xl/pivotCache/pivotCacheDefinition') and name.endswith('.xml'))):
            continue
        document = roots.get(name)
        if document is None:
            document = xml(parts[name])
        changed = False
        for node in document.iter():
            kind = local_name(node)
            settings = {}
            if name == 'xl/connections.xml' and kind == 'connection' and node.getparent() is document:
                settings = {'refreshOnLoad': '0', 'background': '0', 'interval': '0'}
            elif kind == 'queryTable' and node is document:
                settings = {'refreshOnLoad': '0', 'backgroundRefresh': '0', 'disableRefresh': '1'}
            elif kind == 'pivotCacheDefinition' and node is document:
                settings = {'refreshOnLoad': '0', 'backgroundQuery': '0', 'enableRefresh': '0'}
            # Preserve extension structure while disabling an explicitly saved
            # refresh switch, including producer-specific enableRefresh flags.
            for attribute in ('refreshOnLoad', 'backgroundRefresh', 'backgroundQuery', 'enableRefresh'):
                if attribute in node.attrib:
                    settings[attribute] = '0'
            if settings and any(node.get(key) != value for key, value in settings.items()):
                node.attrib.update(settings)
                changed = True
                count += 1
        if changed:
            roots[name] = document
    return count


def _freeze_external_formulas(parts, roots):
    for name in [n for n in parts if n.startswith('xl/worksheets/') and n.endswith('.xml')]:
        sheet = roots.setdefault(name, xml(parts[name]))
        shared = {f.get('si') for f in sheet.iter() if local_name(f) == 'f' and re.search(r'\[\d+\]', f.text or '') and f.get('si') is not None}
        for cell in [n for n in sheet.iter() if local_name(n) == 'c']:
            formula = next((n for n in cell if local_name(n) == 'f'), None)
            if formula is not None and (re.search(r'\[\d+\]', formula.text or '') or formula.get('si') in shared):
                cell.remove(formula)
                value = next((n for n in cell if local_name(n) == 'v'), None)
                if value is None or value.text is None:
                    cell.set('t', 'e')
                    if value is None:
                        value = X.SubElement(cell, cell.tag.rsplit('}',1)[0]+'}v')
                    value.text = '#N/A'


def _replace_chart(parts, roots, chart_name, data, add_image):
    for name in [n for n in parts if n.endswith('.rels')]:
        rels=roots.setdefault(name,xml(parts[name]))
        owner_name=owner_part(name)
        if owner_name not in parts:
            continue
        for rel in rels:
            if not rel.get('Type','').endswith('/chart'):
                continue
            target=rel.get('Target','')
            resolved=target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(posixpath.dirname(owner_name),target))
            if resolved != chart_name:
                continue
            owner=roots.setdefault(owner_name,xml(parts[owner_name]))
            for chart in list(owner.iter()):
                if local_name(chart)!='chart' or rel.get('Id') not in chart.attrib.values():
                    continue
                new_id=add_image(rels,owner_name,data)
                frame=next((n for n in chart.iterancestors() if local_name(n)=='graphicFrame'),None)
                if frame is not None and frame.tag==f'{{{P}}}graphicFrame':
                    fake=X.SubElement(chart.getparent(),f'{{{P}}}oleObj',{f'{{{R}}}id':'ruyiMissingChart'})
                    replace_object(owner,'ruyiMissingChart',new_id)
                elif frame is not None:
                    ns=X.QName(frame).namespace
                    pic=X.Element(f'{{{ns}}}pic')
                    nv=X.SubElement(pic,f'{{{ns}}}nvPicPr')
                    original=next((n for n in frame.iter() if local_name(n)=='cNvPr'),None)
                    if original is not None:nv.append(deepcopy(original))
                    else:X.SubElement(nv,f'{{{ns}}}cNvPr',id='1',name='缺失外链图表')
                    X.SubElement(nv,f'{{{ns}}}cNvPicPr')
                    fill=X.SubElement(pic,f'{{{ns}}}blipFill')
                    X.SubElement(fill,f'{{{A}}}blip',{f'{{{R}}}embed':new_id})
                    X.SubElement(X.SubElement(fill,f'{{{A}}}stretch'),f'{{{A}}}fillRect')
                    props=X.SubElement(pic,f'{{{ns}}}spPr')
                    transform=next((n for n in frame if local_name(n)=='xfrm'),None)
                    if transform is not None:
                        transform=deepcopy(transform);transform.tag=f'{{{A}}}xfrm';props.append(transform)
                    X.SubElement(X.SubElement(props,f'{{{A}}}prstGeom',prst='rect'),f'{{{A}}}avLst')
                    frame.getparent().replace(frame,pic)
                else:
                    # Word inline/anchored drawing: preserve the surrounding wp extent.
                    graphic=chart.getparent()
                    ns='http://schemas.openxmlformats.org/drawingml/2006/picture'
                    graphic.clear();graphic.set('uri',ns)
                    pic=X.SubElement(graphic,f'{{{ns}}}pic')
                    nv=X.SubElement(pic,f'{{{ns}}}nvPicPr')
                    X.SubElement(nv,f'{{{ns}}}cNvPr',id='0',name='缺失外链图表')
                    X.SubElement(nv,f'{{{ns}}}cNvPicPr')
                    fill=X.SubElement(pic,f'{{{ns}}}blipFill')
                    X.SubElement(fill,f'{{{A}}}blip',{f'{{{R}}}embed':new_id})
                    X.SubElement(X.SubElement(fill,f'{{{A}}}stretch'),f'{{{A}}}fillRect')
                    props=X.SubElement(pic,f'{{{ns}}}spPr')
                    transform=X.SubElement(props,f'{{{A}}}xfrm')
                    X.SubElement(transform,f'{{{A}}}off',x='0',y='0')
                    container=next((n for n in graphic.iterancestors() if local_name(n) in ('inline','anchor')),None)
                    extent=next((n for n in container if local_name(n)=='extent'),None) if container is not None else None
                    X.SubElement(transform,f'{{{A}}}ext',cx=extent.get('cx') if extent is not None else '3600000',cy=extent.get('cy') if extent is not None else '2000000')
                    X.SubElement(X.SubElement(props,f'{{{A}}}prstGeom',prst='rect'),f'{{{A}}}avLst')


def _replace_excel_preview(parts, roots, types, sheet_name, sheet, rid, data, add_image, preserve=False):
    objects = [n for n in sheet.iter() if local_name(n) == 'oleObject' and rid in n.attrib.values()]
    ids = {n.get('shapeId') for n in objects}
    found = False
    sheet_rels_name=posixpath.join(posixpath.dirname(sheet_name),'_rels',posixpath.basename(sheet_name)+'.rels')
    sheet_rels=roots.get(sheet_rels_name)
    allowed=set()
    if sheet_rels is not None:
        for rel in sheet_rels:
            if rel.get('Type','').endswith(('/drawing','/vmlDrawing')):
                target=rel.get('Target','')
                allowed.add(target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(posixpath.dirname(sheet_name),target)))
    for name in [n for n in parts if n in allowed and n.endswith(('.xml','.vml'))]:
        drawing = roots.setdefault(name, xml(parts[name]))
        rel_name = posixpath.join(posixpath.dirname(name), '_rels', posixpath.basename(name)+'.rels')
        if rel_name not in parts:
            continue
        rels = roots.setdefault(rel_name, xml(parts[rel_name]))
        for shape in drawing.iter():
            match = (local_name(shape) == 'shape' and shape.get('id','').split('_s')[-1] in ids)
            if local_name(shape) == 'pic':
                match = any(n.get('id') in ids or any(k.rsplit('}',1)[-1]=='spid' and v.split('_s')[-1] in ids for k,v in n.attrib.items()) for n in shape.iter())
            if not match:
                continue
            found = True
            for image in [n for n in shape.iter() if local_name(n) in ('blip','imagedata')]:
                if preserve and any(k.endswith(('}embed','}id')) for k in image.attrib):
                    continue
                new_id = add_image(rels, name, data)
                for key in list(image.attrib):
                    if key.endswith(('}link','}embed','}id')):
                        image.attrib.pop(key)
                image.set(f'{{{R}}}{"embed" if local_name(image)=="blip" else "id"}', new_id)
    if found:
        return
    # Some producers omit the OLE preview completely. Add an explicit sheet drawing.
    ns='http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
    rel_name=posixpath.join(posixpath.dirname(sheet_name),'_rels',posixpath.basename(sheet_name)+'.rels')
    rels=roots.setdefault(rel_name,xml(parts[rel_name]) if rel_name in parts else X.Element(f'{{{REL}}}Relationships'))
    link=next((n for n in sheet if local_name(n)=='drawing'),None)
    target=None
    if link is not None:
        rel=next((r for r in rels if r.get('Id')==link.get(f'{{{R}}}id')),None)
        if rel is not None:
            target=posixpath.normpath(posixpath.join(posixpath.dirname(sheet_name),rel.get('Target','')))
            if rel.get('Target','').startswith('/'):
                target=rel.get('Target').lstrip('/')
    if target not in parts and target not in roots:
        serial=1
        target=f'xl/drawings/ruyi{serial}.xml'
        while target in parts or target in roots:
            serial+=1;target=f'xl/drawings/ruyi{serial}.xml'
        drawing=X.Element(f'{{{ns}}}wsDr',nsmap={'xdr':ns,'a':A,'r':R})
        roots[target]=drawing
        new_id=f'ruyiDrawing{serial}'
        while any(r.get('Id')==new_id for r in rels):new_id+='_' 
        X.SubElement(rels,f'{{{REL}}}Relationship',Id=new_id,Type=R+'/drawing',Target=posixpath.relpath(target,posixpath.dirname(sheet_name)))
        if link is None:
            link=X.Element(sheet.tag.rsplit('}',1)[0]+'}drawing')
            later=next((n for n in sheet if local_name(n) in ('legacyDrawing','legacyDrawingHF','picture','oleObjects','controls','webPublishItems','tableParts','extLst')),None)
            if later is not None:sheet.insert(sheet.index(later),link)
            else:sheet.append(link)
        link.set(f'{{{R}}}id',new_id)
        X.SubElement(types,f'{{{CT}}}Override',PartName='/'+target,ContentType='application/vnd.openxmlformats-officedocument.drawing+xml')
    drawing=roots.setdefault(target,xml(parts[target]) if target in parts else roots[target])
    drawing_rels=posixpath.join(posixpath.dirname(target),'_rels',posixpath.basename(target)+'.rels')
    d_rels=roots.setdefault(drawing_rels,xml(parts[drawing_rels]) if drawing_rels in parts else X.Element(f'{{{REL}}}Relationships'))
    new_id=add_image(d_rels,target,data)
    anchor=X.SubElement(drawing,f'{{{ns}}}twoCellAnchor')
    original=next((n for obj in objects for n in obj.iter() if local_name(n)=='anchor'),None)
    for tag,row,col in [('from',2,0),('to',17,8)]:
        mark=X.SubElement(anchor,f'{{{ns}}}{tag}')
        saved=next((n for n in original if local_name(n)==tag),None) if original is not None else None
        for key,value in [('col',col),('colOff',0),('row',row),('rowOff',0)]:
            old=next((n for n in saved if local_name(n)==key),None) if saved is not None else None
            X.SubElement(mark,f'{{{ns}}}{key}').text=old.text if old is not None else str(value)
    pic=X.SubElement(anchor,f'{{{ns}}}pic')
    nv=X.SubElement(pic,f'{{{ns}}}nvPicPr')
    used=[int(n.get('id')) for n in drawing.iter() if local_name(n)=='cNvPr' and n.get('id','').isdigit()]
    X.SubElement(nv,f'{{{ns}}}cNvPr',id=str(max(used,default=0)+1),name='如意外部对象预览')
    X.SubElement(nv,f'{{{ns}}}cNvPicPr')
    fill=X.SubElement(pic,f'{{{ns}}}blipFill')
    X.SubElement(fill,f'{{{A}}}blip',{f'{{{R}}}embed':new_id})
    X.SubElement(X.SubElement(fill,f'{{{A}}}stretch'),f'{{{A}}}fillRect')
    props=X.SubElement(pic,f'{{{ns}}}spPr')
    X.SubElement(X.SubElement(props,f'{{{A}}}prstGeom',prst='rect'),f'{{{A}}}avLst')
    X.SubElement(anchor,f'{{{ns}}}clientData')
