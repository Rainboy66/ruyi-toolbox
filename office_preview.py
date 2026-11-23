"""Remove optional Office thumbnails without changing document content."""
import os
import posixpath
import tempfile
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree

REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
THUMB_TYPE = REL_NS + '/metadata/thumbnail'


def remove_office_thumbnail(office_file):
    office_file = Path(office_file)
    parser = etree.XMLParser(resolve_entities=False, no_network=True)
    temporary = None
    try:
        with ZipFile(office_file) as source:
            rels = etree.fromstring(source.read('_rels/.rels'), parser)
            types = etree.fromstring(source.read('[Content_Types].xml'), parser)
            removed = {n for n in source.namelist() if n.lower().startswith('docprops/thumbnail.')}
            found = False
            for relation in list(rels):
                if relation.get('Type') == THUMB_TYPE:
                    found = True
                    if relation.get('TargetMode') != 'External':
                        part = posixpath.normpath(relation.get('Target', '').lstrip('/'))
                        if part.startswith('docProps/'):
                            removed.add(part)
                    rels.remove(relation)
            if not found and not removed:
                return False
            for item in list(types):
                if item.tag == f'{{{CT_NS}}}Override' and item.get('PartName', '').lstrip('/') in removed:
                    types.remove(item)
            replacements = {
                '_rels/.rels': etree.tostring(rels, xml_declaration=True, encoding='UTF-8', standalone=True),
                '[Content_Types].xml': etree.tostring(types, xml_declaration=True, encoding='UTF-8', standalone=True),
            }
            fd, name = tempfile.mkstemp(prefix='no-preview-', suffix=office_file.suffix, dir=office_file.parent)
            os.close(fd)
            temporary = Path(name)
            with ZipFile(temporary, 'w', ZIP_DEFLATED) as target:
                for info in source.infolist():
                    if info.filename not in removed:
                        target.writestr(info, replacements.get(info.filename, source.read(info.filename)))
        os.replace(temporary, office_file)
        return True
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)
