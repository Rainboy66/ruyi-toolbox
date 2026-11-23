"""Keep subpaths from one PDF paint operation together, including holes."""
import ctypes


def _parts(commands):
    result=[]
    for command in commands:
        if command[0]=='m':result.append([])
        if result:result[-1].append(command)
    return result


def _polygon(commands):
    points=[];position=None
    for command in commands:
        kind=command[0]
        if kind in ('m','l'):
            position=command[1];points.append(position)
        elif kind=='c' and position:
            p0=position;p1,p2,p3=command[1:]
            for i in range(1,17):
                t=i/16;u=1-t
                points.append(tuple(u**3*p0[k]+3*u*u*t*p1[k]+3*u*t*t*p2[k]+t**3*p3[k] for k in (0,1)))
            position=p3
    return points


def _area(points):
    return sum(a[0]*b[1]-b[0]*a[1] for a,b in zip(points,points[1:]+points[:1]))/2


def _inside(point,polygon):
    x,y=point;inside=False
    for a,b in zip(polygon,polygon[1:]+polygon[:1]):
        if (a[1]>y)!=(b[1]>y) and x<(b[0]-a[0])*(y-a[1])/(b[1]-a[1])+a[0]:inside=not inside
    return inside


def _reverse(commands):
    start=commands[0][1];position=start;segments=[]
    for command in commands[1:]:
        if command[0]=='h':continue
        segments.append((position,command));position=command[-1]
    result=[('m',position)]
    for previous,command in reversed(segments):
        if command[0]=='l':result.append(('l',previous))
        elif command[0]=='c':result.append(('c',command[2],command[1],previous))
    if any(c[0]=='h' for c in commands):result.append(('h',))
    return result


def _winding(commands):
    parts=_parts(commands);polygons=[_polygon(part) for part in parts]
    areas=[_area(poly) for poly in polygons];result=[]
    for i,part in enumerate(parts):
        if not polygons[i]:continue
        depth=sum(abs(areas[j])>abs(areas[i]) and _inside(polygons[i][0],poly)
                  for j,poly in enumerate(polygons) if j!=i and poly)
        sign=1 if depth%2==0 else -1
        result.extend(_reverse(part) if areas[i]*sign<0 else part)
    return result


def compound_vectors(page,pdf_page):
    from pypdfium2 import raw
    originals=[*page.rects,*page.lines,*page.curves]
    claimed=set();merged=[]
    for obj in pdf_page.get_objects(filter=[raw.FPDF_PAGEOBJ_PATH]):
        count=raw.FPDFPath_CountSegments(obj)
        if count<2:continue
        segments=[raw.FPDFPath_GetPathSegment(obj,i) for i in range(count)]
        if sum(raw.FPDFPathSegment_GetType(s)==raw.FPDF_SEGMENT_MOVETO for s in segments)<2:continue
        matrices=[];current=obj
        while current is not None:
            matrices.append(current.get_matrix());current=current.container
        def point(segment):
            x,y=ctypes.c_float(),ctypes.c_float()
            if not raw.FPDFPathSegment_GetPoint(segment,ctypes.byref(x),ctypes.byref(y)):
                raise ValueError('Cannot read compound PDF path point')
            px,py=x.value,y.value
            for matrix in matrices:px,py=matrix.on_point(px,py)
            return (px,page.height-py)
        commands=[];i=0
        while i<count:
            segment=segments[i];kind=raw.FPDFPathSegment_GetType(segment)
            if kind==raw.FPDF_SEGMENT_MOVETO:commands.append(('m',point(segment)))
            elif kind==raw.FPDF_SEGMENT_LINETO:commands.append(('l',point(segment)))
            elif kind==raw.FPDF_SEGMENT_BEZIERTO and i+2<count:
                commands.append(('c',point(segment),point(segments[i+1]),point(segments[i+2])));i+=2
                segment=segments[i]
            if raw.FPDFPathSegment_GetClose(segment):commands.append(('h',))
            i+=1
        pieces=_parts(commands);matches=[]
        for part in pieces:
            first=part[0][1]
            candidates=[]
            for index,item in enumerate(originals):
                if index in claimed or index in matches or not item.get('path'):continue
                start=item['path'][0]
                if start[0]!='m':continue
                error=abs(start[1][0]-first[0])+abs(start[1][1]-first[1])
                if error<.08:candidates.append((error,index))
            if not candidates:break
            matches.append(min(candidates)[1])
        if len(matches)!=len(pieces) or len(matches)<2:continue
        # Keeping one source paint operation as one DrawingML path preserves
        # the nonzero winding relationship which pdfminer splits into objects.
        fill_mode,stroke=ctypes.c_int(),raw.FPDF_BOOL()
        if raw.FPDFPath_GetDrawMode(obj,ctypes.byref(fill_mode),ctypes.byref(stroke)) and fill_mode.value==1:
            commands=_winding(commands)
        source=[originals[n] for n in matches];record=dict(source[0])
        record.update(object_type='curve',path=commands,
                      x0=min(v['x0'] for v in source),x1=max(v['x1'] for v in source),
                      top=min(v['top'] for v in source),bottom=max(v['bottom'] for v in source),
                      _ppt_order=min(v.get('_ppt_order',0) for v in source))
        record['pts']=[p for command in commands for p in command[1:]]
        claimed.update(matches);merged.append(record)
    return [v for i,v in enumerate(originals) if i not in claimed]+merged
