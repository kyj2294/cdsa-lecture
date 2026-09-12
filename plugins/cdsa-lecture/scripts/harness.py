"""CDSA lecture delivery gate.

Deterministic structural checks + a record that visual review happened.
Visual judgement itself is the reviewer's job; this script only refuses to
release when the review is missing, stale, or incomplete.
"""
import argparse, copy, hashlib, json, math, os, pathlib, re, shutil, sqlite3, sys, tempfile, zipfile
import xml.etree.ElementTree as E

ROOT = pathlib.Path(__file__).resolve().parents[1]
SKILL = ROOT / 'skills/cdsappt'
sys.path.insert(0, str(SKILL / 'scripts'))
import jeju_template as jt
import shape_labels
import teaching_visuals
import visual_retrieval
import reuse_slide

P, A, R = jt.P, 'http://schemas.openxmlformats.org/drawingml/2006/main', jt.R
NS = {'p': P, 'a': A}

# Visual review items. Keep this list short: every item must be checked on every page by a human/agent.
CHECKS = ['template_match', 'text_fit', 'content_accuracy', 'visual_fit', 'learner_value']
COVER_CHECKS = ['cover_style']
ROLES = ('cover', 'content', 'transition')
ANSWER_KEYS = ('topic', 'audience', 'duration_minutes', 'institution', 'materials', 'cover_mood')
COURSE_LABEL = '2026년 AI 챔피언 고급과정 인증자 1차 보수교육'
GUEST_META = {'혁신의 누명': 'title', '리버스엔지니어링과 디자인싱킹': 'subtitle',
              'AI챔피언 전문인재 보수교육 in Jeju': 'title', '2026-06-05': 'date',
              '한국데이터사이언티스트협회 김태유 박사': 'presenter'}


def library_db():
    """User-built corpus location; override with CDSA_LECTURE_DB."""
    return pathlib.Path(os.environ.get('CDSA_LECTURE_DB') or SKILL / 'data/lecture-library.sqlite')


def sha(path):
    with pathlib.Path(path).open('rb') as f:
        return hashlib.file_digest(f, 'sha256').hexdigest()

def blobsha(data):
    return hashlib.sha256(data).hexdigest()

def load(p):
    return json.loads(p.read_text(encoding='utf-8-sig'))

def save(p, obj):
    p.write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding='utf-8')

def resolve(run, name):
    p = (run / name).resolve()
    if not p.is_relative_to(run.resolve()):
        raise ValueError('Run asset must be inside run folder: ' + str(name))
    return p

def bbox(s):
    x = s.find('p:spPr/a:xfrm', NS)
    if x is None:
        x = s.find('p:xfrm', NS)
    if x is None:
        return None
    o = x.find('a:off', NS); e = x.find('a:ext', NS)
    if o is None or e is None:
        return None
    return tuple(int(v) for v in (o.get('x'), o.get('y'), e.get('cx'), e.get('cy')))


def _matrix_mul(left,right):
    a,b,c,d,e,f=left;g,h,i,j,k,l=right
    return (a*g+c*h,b*g+d*h,a*i+c*j,b*i+d*j,a*k+c*l+e,b*k+d*l+f)


def _matrix_point(matrix,x,y):
    a,b,c,d,e,f=matrix
    return a*x+c*y+e,b*x+d*y+f


def _transform_matrix(xfrm,off,ch_off=None,ch_ext=None):
    ext=xfrm.find('a:ext',NS);cx=float(ext.get('cx'));cy=float(ext.get('cy'))
    ox=float(off.get('x'));oy=float(off.get('y'));mx=ox+cx/2;my=oy+cy/2
    angle=math.radians(float(xfrm.get('rot','0'))/60000.0);co=math.cos(angle);si=math.sin(angle)
    fx=-1.0 if xfrm.get('flipH') in ('1','true') else 1.0
    fy=-1.0 if xfrm.get('flipV') in ('1','true') else 1.0
    if ch_off is None:return (co*fx,si*fx,-si*fy,co*fy,mx-co*fx*mx+si*fy*my,my-si*fx*mx-co*fy*my)
    cw=float(ch_ext.get('cx'));ch=float(ch_ext.get('cy'))
    if cw<=0 or ch<=0:raise ValueError('Group child extent must be positive')
    ccx=float(ch_off.get('x'))+cw/2;ccy=float(ch_off.get('y'))+ch/2
    scale=(cx/cw,0,0,cy/ch,0,0)
    rotate_flip=(co*fx,si*fx,-si*fy,co*fy,0,0)
    return _matrix_mul((1,0,0,1,mx,my),_matrix_mul(rotate_flip,_matrix_mul(scale,(1,0,0,1,-ccx,-ccy))))


def _aabb(matrix,off,ext):
    x=float(off.get('x'));y=float(off.get('y'));w=float(ext.get('cx'));height=float(ext.get('cy'))
    points=[_matrix_point(matrix,px,py) for px,py in ((x,y),(x+w,y),(x,y+height),(x+w,y+height))]
    left=min(p[0] for p in points);top=min(p[1] for p in points)
    return left,top,max(p[0] for p in points)-left,max(p[1] for p in points)-top


def shape_boxes(container, transform=(1.0,0.0,0.0,1.0,0.0,0.0), path=()):
    """Yield actual canvas AABBs, including nested group rotation/scale/flip."""
    if container is None:return
    for shape in container:
        kind=shape.tag.split('}')[-1]
        if kind in ('nvGrpSpPr','grpSpPr'):continue
        sid,name=reuse_slide.direct_shape_identity(shape)
        if kind=='grpSp':
            xf=shape.find('p:grpSpPr/a:xfrm',NS)
            if xf is None:continue
            off,ext,ch_off,ch_ext=[xf.find('a:'+key,NS) for key in ('off','ext','chOff','chExt')]
            if any(value is None for value in (off,ext,ch_off,ch_ext)):
                raise ValueError('Group transform is incomplete: '+str(name or sid or '?'))
            outer=_matrix_mul(transform,_transform_matrix(xf,off))
            yield shape,_aabb(outer,off,ext),name or sid or '?',path
            nested=_matrix_mul(transform,_transform_matrix(xf,off,ch_off,ch_ext))
            yield from shape_boxes(shape,nested,path+(sid or '?',))
            continue
        xf=reuse_slide.xfrm_of(shape)
        if xf is None:continue
        off=xf.find('a:off',NS);ext=xf.find('a:ext',NS)
        if off is None or ext is None:continue
        local=_matrix_mul(transform,_transform_matrix(xf,off))
        yield shape,_aabb(local,off,ext),name or sid or '?',path


CORE_CONTENT_TYPES={
    'ppt/presentation.xml':'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
    'ppt/slides/':'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
    'ppt/slideMasters/':'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml',
    'ppt/slideLayouts/':'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml',
    'ppt/theme/':'application/vnd.openxmlformats-officedocument.theme+xml',
    'ppt/charts/':'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
}


def inspect_package_graph(files, fail):
    """Reject relationship, page-registration and core content-type ambiguity."""
    for name,data in files.items():
        if name.endswith('.xml') or name.endswith('.rels'):
            try:jt.xml(data)
            except Exception as exc:fail('PACKAGE-01','Malformed XML part '+name+': '+str(exc))
    for name,data in files.items():
        if not name.endswith('.rels'):continue
        rels=list(jt.xml(data));ids=[r.get('Id') for r in rels]
        if any(not rid for rid in ids) or len(ids)!=len(set(ids)):
            fail('PACKAGE-01','Duplicate or missing relationship ID in '+name)

    pres=jt.xml(files['ppt/presentation.xml'])
    prels=list(jt.xml(files['ppt/_rels/presentation.xml.rels']))
    relmap={r.get('Id'):r for r in prels}
    entries=list(pres.find('p:sldIdLst',NS) or [])
    slide_ids=[];slide_rids=[];ordered=[]
    for entry in entries:
        try:value=int(entry.get('id'));slide_ids.append(value)
        except (TypeError,ValueError):fail('PACKAGE-01','Invalid presentation slide ID');continue
        if value<256 or value>2147483647:fail('PACKAGE-01','Presentation slide ID is outside the PowerPoint range')
        rid=entry.get('{'+R+'}id');slide_rids.append(rid);rel=relmap.get(rid)
        if rel is None or rel.get('Type')!=R+'/slide':
            fail('PACKAGE-01','Presentation slide ID does not resolve to a slide relationship');continue
        target=jt.target('ppt/presentation.xml',rel);ordered.append(target)
        if target not in files:fail('PACKAGE-01','Registered slide part is missing: '+target)
    if len(slide_ids)!=len(set(slide_ids)):fail('PACKAGE-01','Duplicate presentation slide IDs')
    if len(slide_rids)!=len(set(slide_rids)):fail('PACKAGE-01','Duplicate presentation slide relationship IDs')
    if len(ordered)!=len(set(ordered)):fail('PACKAGE-01','Two presentation pages resolve to the same slide part')
    all_parts={n for n in files if re.fullmatch(r'ppt/slides/[^/]+\.xml',n)}
    for part in all_parts:
        shape_ids=[node.get('id') for node in jt.xml(files[part]).findall('.//p:cNvPr',NS)]
        try:numeric_shape_ids=[int(value) for value in shape_ids]
        except (TypeError,ValueError):numeric_shape_ids=[]
        if (len(numeric_shape_ids)!=len(shape_ids)
                or any(value<=0 or value>4294967295 for value in numeric_shape_ids)
                or len(numeric_shape_ids)!=len(set(numeric_shape_ids))):
            fail('PACKAGE-01','Slide shape IDs must be nonempty numeric and unique: '+part)
    slide_rel_targets=[jt.target('ppt/presentation.xml',r) for r in prels if r.get('Type')==R+'/slide']
    if set(ordered)!=all_parts or len(ordered)!=len(all_parts):
        fail('PACKAGE-01','Slide XML parts must exactly match the ordered presentation pages')
    if set(slide_rel_targets)!=all_parts or len(slide_rel_targets)!=len(all_parts):
        fail('PACKAGE-01','Presentation slide relationships must exactly match slide XML parts')
    orphan_rels=[]
    for name in files:
        match=re.fullmatch(r'ppt/slides/_rels/(.+\.xml)\.rels',name)
        if match and 'ppt/slides/'+match.group(1) not in all_parts:orphan_rels.append(name)
    if orphan_rels:fail('PACKAGE-01','Orphan slide relationship part: '+orphan_rels[0])

    ct=list(jt.xml(files['[Content_Types].xml']))
    overrides=[x for x in ct if x.tag=='{'+jt.CT+'}Override']
    part_names=[x.get('PartName','').lstrip('/') for x in overrides]
    if len(part_names)!=len(set(part_names)):fail('PACKAGE-01','Duplicate Content_Types Override PartName')
    override_map={x.get('PartName','').lstrip('/'):x.get('ContentType') for x in overrides}
    for name in files:
        expected=None
        if name=='ppt/presentation.xml':expected=CORE_CONTENT_TYPES[name]
        else:
            for prefix,content_type in CORE_CONTENT_TYPES.items():
                if prefix.endswith('/') and name.startswith(prefix) and name.endswith('.xml'):
                    if prefix=='ppt/charts/' and not re.fullmatch(r'ppt/charts/chart\d+(?:-r\d+)?\.xml',name):continue
                    expected=content_type;break
        if expected is not None and override_map.get(name)!=expected:
            fail('PACKAGE-01','Missing or wrong core content type for '+name)
    return ordered

def addtext(tree, sid, name, text, box, color='FFFFFF', size=950):
    sp = E.SubElement(tree, '{'+P+'}sp')
    nv = E.SubElement(sp, '{'+P+'}nvSpPr')
    E.SubElement(nv, '{'+P+'}cNvPr', id=str(sid), name=name)
    E.SubElement(nv, '{'+P+'}cNvSpPr'); E.SubElement(nv, '{'+P+'}nvPr')
    pr = E.SubElement(sp, '{'+P+'}spPr'); xf = E.SubElement(pr, '{'+A+'}xfrm')
    E.SubElement(xf, '{'+A+'}off', x=str(box[0]), y=str(box[1]))
    E.SubElement(xf, '{'+A+'}ext', cx=str(box[2]), cy=str(box[3]))
    tx = E.SubElement(sp, '{'+P+'}txBody')
    E.SubElement(tx, '{'+A+'}bodyPr'); E.SubElement(tx, '{'+A+'}lstStyle')
    pa = E.SubElement(tx, '{'+A+'}p'); r = E.SubElement(pa, '{'+A+'}r')
    rp = E.SubElement(r, '{'+A+'}rPr', sz=str(size))
    f = E.SubElement(rp, '{'+A+'}solidFill'); E.SubElement(f, '{'+A+'}srgbClr', val=color)
    E.SubElement(rp, '{'+A+'}latin', typeface='Noto Sans KR Black')
    E.SubElement(rp, '{'+A+'}ea', typeface='페이퍼로지 8 ExtraBold')
    E.SubElement(r, '{'+A+'}t').text = text
    return sp


# --------------------------------------------------------------------------- branding

def _check_logo(path):
    from PIL import Image
    im = Image.open(path)
    if im.format != 'PNG':
        raise ValueError('Logo must be PNG: ' + str(path))
    if 'A' not in im.getbands() or im.getchannel('A').getextrema()[0] == 255:
        raise ValueError('Logo has no transparent pixels: ' + str(path))
    return im

def brand(run, request):
    """Apply course title + institution logo to the Jeju master. This is the only permitted baseline change."""
    logo_spec = request.get('logo', {})
    mode = logo_spec.get('mode', 'placeholder')
    im = None
    if mode != 'none':
        logo = resolve(run, logo_spec['file']); im = _check_logo(logo)
    base = jt.read(jt.DEFAULT)
    if im is not None:
        base['ppt/media/ssj-institution.png'] = logo.read_bytes()
    light_im = None
    if mode != 'none' and logo_spec.get('light_file'):
        light = resolve(run, logo_spec['light_file']); light_im = _check_logo(light)
        base['ppt/media/ssj-institution-light.png'] = light.read_bytes()
    for part in list(base):
        if not re.fullmatch(r'ppt/slideLayouts/slideLayout\d+\.xml', part):
            continue
        e = jt.xml(base[part]); tree = e.find('p:cSld/p:spTree', NS); rr = jt.xml(base[jt.relpart(part)])
        layout_id = int(re.search(r'slideLayout(\d+)', part)[1])
        use_light = light_im is not None and layout_id in (10, 11, 12, 13)
        active_im = light_im if use_light else im
        found = False
        for sp in tree:
            ts = sp.findall('.//a:t', NS); full = ''.join(t.text or '' for t in ts)
            if COURSE_LABEL in full:
                ts[0].text = request['title']
                for t in ts[1:]: t.text = ''
                found = True
            elif layout_id == 11 and full in GUEST_META:
                key = GUEST_META[full]
                ts[0].text = request['title'] if key == 'title' else request.get(key, '')
                for t in ts[1:]: t.text = ''
        pics = [s for s in tree if s.tag == '{'+P+'}pic' and bbox(s) and bbox(s)[0] > 9*914400 and bbox(s)[1] < 0.5*914400]
        maxid = max([int(s.get('id', '0')) for s in e.findall('.//p:cNvPr', NS)] + [0])
        if pics:
            anchor = copy.deepcopy(pics[0]); boxes = [bbox(s) for s in pics]
            x = min(b[0] for b in boxes); y = min(b[1] for b in boxes)
            w = max(b[0]+b[2] for b in boxes) - x; h = max(b[1]+b[3] for b in boxes) - y
            for s in pics: tree.remove(s)
        elif mode != 'none':
            anchor = E.Element('{'+P+'}pic')
            nv = E.SubElement(anchor, '{'+P+'}nvPicPr')
            E.SubElement(nv, '{'+P+'}cNvPr', id=str(maxid+1), name='SSJ institution logo')
            E.SubElement(nv, '{'+P+'}cNvPicPr'); E.SubElement(nv, '{'+P+'}nvPr')
            fill = E.SubElement(anchor, '{'+P+'}blipFill'); E.SubElement(fill, '{'+A+'}blip')
            stretch = E.SubElement(fill, '{'+A+'}stretch'); E.SubElement(stretch, '{'+A+'}fillRect')
            pr = E.SubElement(anchor, '{'+P+'}spPr'); xf = E.SubElement(pr, '{'+A+'}xfrm')
            E.SubElement(xf, '{'+A+'}off'); E.SubElement(xf, '{'+A+'}ext')
            x, y, w, h = (int(v*914400) for v in (10.85, 0.06, 2.30, 0.34))
        for rel in list(rr):
            if rel.get('Id') == 'rIdSsjLogo':
                rr.remove(rel)
        if mode != 'none':
            scale = min(w/active_im.width, h/active_im.height)
            nw, nh = round(active_im.width*scale), round(active_im.height*scale)
            xf = anchor.find('p:spPr/a:xfrm', NS)
            xf.find('a:off', NS).attrib.update(x=str(x+w-nw), y=str(y+(h-nh)//2))
            xf.find('a:ext', NS).attrib.update(cx=str(nw), cy=str(nh))
            blip = anchor.find('.//a:blip', NS); blip.clear(); blip.set('{'+R+'}embed', 'rIdSsjLogo')
            crop = anchor.find('p:blipFill/a:srcRect', NS)
            if crop is not None: anchor.find('p:blipFill', NS).remove(crop)
            E.SubElement(rr, '{'+jt.PKG+'}Relationship', Id='rIdSsjLogo', Type=R+'/image',
                         Target='../media/ssj-institution-light.png' if use_light else '../media/ssj-institution.png')
            tree.append(anchor)
        if not found:
            addtext(tree, maxid+2, 'SSJ course name', request['title'], (450000, 60000, 7000000, 260000), color='333333')
        base[part] = jt.serialize(e); base[jt.relpart(part)] = jt.serialize(rr)
    ct = jt.xml(base['[Content_Types].xml'])
    if not any(c.get('Extension') == 'png' for c in ct):
        E.SubElement(ct, '{'+jt.CT+'}Default', Extension='png', ContentType='image/png')
    base['[Content_Types].xml'] = jt.serialize(ct)
    return base


# --------------------------------------------------------------------------- commands

def questions(req):
    missing = [k for k in ANSWER_KEYS if not str(req.get('answers', {}).get(k, {}).get('value', '')).strip()]
    if not req.get('title'):
        missing.append('title')
    return missing


def visible_text(files, slide_part):
    parts = [slide_part]
    try:
        layout, master = reuse_slide.layout_and_master(files, slide_part)
        parts.extend([layout, master])
    except Exception:
        pass
    values = []
    for part in parts:
        if part and part in files:
            values.extend((node.text or '') for node in jt.xml(files[part]).findall('.//a:t', NS))
    return '\n'.join(values)


def textfit_issues(slide, page):
    issues = []
    for shape in slide.findall('p:cSld/p:spTree/*', NS):
        if shape.find('p:txBody', NS) is None:
            continue
        body = shape.find('p:txBody/a:bodyPr', NS)
        if body is not None and (body.find('a:spAutoFit', NS) is not None or body.find('a:normAutofit', NS) is not None):
            continue
        box = bbox(shape)
        if not box or box[2] <= 0 or box[3] <= 0:
            continue
        left = int(body.get('lIns', '91440')) if body is not None else 91440
        right = int(body.get('rIns', '91440')) if body is not None else 91440
        top = int(body.get('tIns', '45720')) if body is not None else 45720
        bottom = int(body.get('bIns', '45720')) if body is not None else 45720
        width = max(1, box[2] - left - right); height = max(1, box[3] - top - bottom)
        used = 0.0
        for paragraph in shape.findall('.//a:p', NS):
            text = ''.join(node.text or '' for node in paragraph.findall('.//a:t', NS))
            if not text:
                continue
            sizes = [int(node.get('sz')) for node in paragraph.findall('.//a:rPr', NS) + paragraph.findall('.//a:defRPr', NS) if node.get('sz')]
            size = (max(sizes) if sizes else 1800) / 100 * 12700
            units = sum(1.0 if ('가' <= char <= '힣' or ord(char) > 255) else 0.55 for char in text)
            lines = max(1, math.ceil(units * size / width))
            used += lines * size * 1.2
        if used > height * 1.1:
            name = shape.find('.//p:cNvPr', NS)
            issues.append({'rule': 'TEXTFIT-01', 'page': page,
                           'message': 'Estimated text overflow: ' + (name.get('name', '?') if name is not None else '?')})
    return issues

def reuse_record(item):
    reuse = item.get('reuse')
    return isinstance(reuse, dict) and reuse.get('mode') == 'reuse'


def inspect_reuse_record(run, item, part, files, db_path, fail, page):
    """Validate tool-produced provenance while allowing later geometry/style edits."""
    reuse = item.get('reuse', {})
    if not reuse_record(item):
        return set()
    shapes = reuse.get('shapes')
    if not isinstance(shapes, list) or not shapes or len(shapes) != len(set(map(str, shapes))):
        fail('REUSE-01', 'reuse.shapes must contain unique candidate shape IDs', page)
        return set()
    tree = jt.xml(files[part]).find('p:cSld/p:spTree', NS)
    actual = []
    for shape in tree:
        if shape.tag.split('}')[-1] in ('nvGrpSpPr', 'grpSpPr'):
            continue
        sid, _ = reuse_slide.direct_shape_identity(shape)
        if sid is not None:
            actual.append(str(sid))
    if len(actual) != len(set(actual)):
        fail('PACKAGE-01', 'Duplicate top-level shape IDs', page)
    missing = sorted(set(map(str, shapes)) - set(actual))
    if missing:
        fail('REUSE-01', 'Recorded reused shapes are missing: ' + ', '.join(missing), page)
    source = str(reuse.get('source', ''))
    match = re.fullmatch(r'ssj-slide://(s-[0-9a-f]{20})', source)
    if match:
        try:
            with sqlite3.connect('file:' + db_path.resolve().as_posix() + '?mode=ro', uri=True) as db:
                if not db.execute('SELECT 1 FROM slides WHERE id=?', (match.group(1),)).fetchone():
                    fail('REUSE-01', 'Reused source slide is absent from the library', page)
        except (OSError, sqlite3.Error) as exc:
            fail('REUSE-01', 'Could not validate reused source: ' + str(exc), page)
    elif not source.strip():
        fail('REUSE-01', 'reuse.source is required', page)
    fit = reuse.get('fit')
    if not isinstance(fit, dict) or not isinstance(fit.get('scale'), (int, float)) or fit.get('scale', 0) <= 0:
        fail('REUSE-01', 'reuse.fit must record a positive scale', page)
    warnings=reuse.get('warnings',[])
    if isinstance(warnings,list):
        for warning in warnings:
            if str(warning).startswith('contrast risk:'):
                fail('CONTRAST-01',str(warning),page)
            elif str(warning).startswith('density risk:'):
                fail('DENSITY-01',str(warning),page)
    return set(map(str, shapes))

def _source_path(run,value):
    path=pathlib.Path(value)
    return path if path.is_absolute() else resolve(run,value)


def _db_source_fingerprint(db_path,sid,cache):
    if sid not in cache:
        with sqlite3.connect('file:'+db_path.resolve().as_posix()+'?mode=ro',uri=True) as db:
            cache[sid]=visual_retrieval.source_fingerprint(db,sid)
    return cache[sid]


def _cached_file_sha(path,cache):
    key=str(path.resolve())
    if key not in cache:cache[key]=sha(path)
    return cache[key]


def _cached_package(path,cache):
    key=str(path.resolve())
    if key not in cache:cache[key]=jt.read(path)
    return cache[key]


def inspect_master_graph(files, fail):
    """PowerPoint-specific master/layout identity and relationship invariants."""
    pres=jt.xml(files['ppt/presentation.xml']);prels=list(jt.xml(files['ppt/_rels/presentation.xml.rels']))
    rel_ids=[r.get('Id') for r in prels]
    if len(rel_ids)!=len(set(rel_ids)):fail('MASTER-02','Duplicate presentation relationship IDs')
    relmap={r.get('Id'):r for r in prels};entries=list(pres.find('p:sldMasterIdLst',NS) or [])
    master_ids=[];registered=[];themes=[];layout_ids=[]
    for entry in entries:
        try:mid=int(entry.get('id'));master_ids.append(mid)
        except (TypeError,ValueError):fail('MASTER-02','Invalid presentation master ID');continue
        rid=entry.get('{'+R+'}id');rel=relmap.get(rid)
        if mid<2147483648:fail('MASTER-02','Presentation master ID is outside the PowerPoint range')
        if rel is None or rel.get('Type')!=R+'/slideMaster':
            fail('MASTER-02','Presentation master ID does not resolve to a slideMaster relationship');continue
        master=jt.target('ppt/presentation.xml',rel);registered.append(master)
        if master not in files:
            fail('MASTER-02','Registered slide master part is missing: '+master);continue
        rp=jt.relpart(master)
        if rp not in files:
            fail('MASTER-02','Registered slide master relationships are missing: '+master);continue
        mrels=list(jt.xml(files[rp]));mids=[r.get('Id') for r in mrels]
        if len(mids)!=len(set(mids)):fail('MASTER-02','Duplicate relationship IDs in '+rp)
        mmap={r.get('Id'):r for r in mrels}
        theme=[jt.target(master,r) for r in mrels if r.get('Type')==R+'/theme']
        if len(theme)!=1:fail('MASTER-02','Each registered slide master must resolve exactly one theme: '+master)
        else:
            themes.extend(theme)
            if theme[0] not in files:fail('MASTER-02','Registered slide master theme is missing: '+theme[0])
        listed_layouts=[];layout_nodes=jt.xml(files[master]).findall('.//p:sldLayoutId',NS)
        if not layout_nodes:fail('MASTER-02','Registered slide master has no layouts: '+master)
        for lid in layout_nodes:
            try:value=int(lid.get('id'));layout_ids.append(value)
            except (TypeError,ValueError):fail('MASTER-02','Invalid slide layout ID in '+master);continue
            if value<2147483648:fail('MASTER-02','Slide layout ID is outside the PowerPoint range in '+master)
            rel=mmap.get(lid.get('{'+R+'}id'))
            if rel is None or rel.get('Type')!=R+'/slideLayout' or jt.target(master,rel) not in files:
                fail('MASTER-02','Slide layout ID does not resolve to a valid layout in '+master)
                continue
            layout=jt.target(master,rel);listed_layouts.append(layout)
            lrp=jt.relpart(layout)
            if lrp not in files:
                fail('MASTER-02','Slide layout relationships are missing: '+layout);continue
            lrels=list(jt.xml(files[lrp]));lids=[r.get('Id') for r in lrels]
            if len(lids)!=len(set(lids)):fail('MASTER-02','Duplicate relationship IDs in '+lrp)
            back=[jt.target(layout,r) for r in lrels if r.get('Type')==R+'/slideMaster']
            if back!=[master]:fail('MASTER-02','Slide layout must resolve back to its registered master: '+layout)
        related_layouts=[jt.target(master,r) for r in mrels if r.get('Type')==R+'/slideLayout']
        if len(listed_layouts)!=len(set(listed_layouts)) or sorted(listed_layouts)!=sorted(related_layouts):
            fail('MASTER-02','Master layout IDs and relationships do not match in '+master)
    if len(master_ids)!=len(set(master_ids)):fail('MASTER-02','Duplicate presentation master IDs')
    if len(layout_ids)!=len(set(layout_ids)):fail('MASTER-02','Duplicate slide layout IDs across masters')
    if len(registered)!=len(set(registered)):fail('MASTER-02','Duplicate registered slide master targets')
    if len(themes)!=len(set(themes)):fail('MASTER-02','Each registered master must use its own theme part')
    related_masters=[jt.target('ppt/presentation.xml',r) for r in prels if r.get('Type')==R+'/slideMaster']
    if sorted(registered)!=sorted(related_masters):fail('MASTER-02','Presentation master IDs and relationships do not match')
    slide_relmap={r.get('Id'):r for r in prels};live_masters=[]
    for slide_id in list(pres.find('p:sldIdLst',NS) or []):
        rel=slide_relmap.get(slide_id.get('{'+R+'}id'))
        if rel is None or rel.get('Type')!=R+'/slide':continue
        slide=jt.target('ppt/presentation.xml',rel)
        srp=jt.relpart(slide)
        if srp not in files:
            fail('MASTER-02','Live slide is missing its layout/master chain: '+slide);continue
        layout_rels=[r for r in jt.xml(files[srp]) if r.get('Type')==R+'/slideLayout']
        if len(layout_rels)!=1:
            fail('MASTER-02','Live slide must resolve exactly one slide layout: '+slide);continue
        layout=jt.target(slide,layout_rels[0]);lrp=jt.relpart(layout)
        if layout not in files or lrp not in files:
            fail('MASTER-02','Live slide layout is missing: '+layout);continue
        master_rels=[r for r in jt.xml(files[lrp]) if r.get('Type')==R+'/slideMaster']
        if len(master_rels)!=1:
            fail('MASTER-02','Live slide layout must resolve exactly one master: '+layout);continue
        master=jt.target(layout,master_rels[0])
        live_masters.append(master)
    if live_masters and not registered:fail('MASTER-02','A presentation with live slides must register at least one master')
    if set(live_masters)!=set(registered):
        fail('MASTER-02','Registered masters must exactly match masters reached by live slides')

def init(run):
    run.mkdir(parents=True, exist_ok=True)
    protected = [run/name for name in ('request.json','plan.json','assets.json','template.pptx','candidate.pptx')]
    if any(path.exists() for path in protected):
        raise ValueError('Run folder already contains cdsappt work files')
    save(run/'request.json', {'title': '', 'subtitle': '', 'presenter': '', 'date': '',
                              'answers': {k: {'value': ''} for k in ANSWER_KEYS},
                              'fixed_slide_count': None,
                              'logo': {'mode': 'placeholder', 'file': 'logo.png', 'light_file': 'logo-dark.png',
                                       'source': '', 'institution': ''}})
    save(run/'plan.json', {'slides': []})
    save(run/'assets.json', {'images': []})
    print('Initialised', run, '- fill request.json from what the user already said; ask only missing fields.')

def prepare(run):
    req = load(run/'request.json'); missing = questions(req)
    if missing:
        raise ValueError('Missing user answers: ' + ', '.join(missing))
    logo = req.setdefault('logo', {})
    mode = logo.get('mode', 'placeholder')
    if mode not in ('official', 'user', 'placeholder', 'none'):
        raise ValueError('logo.mode must be official, user, placeholder or none')
    institution = str(req['answers']['institution']['value']).strip()
    logo['institution'] = institution
    if mode in ('official', 'user'):
        if not logo.get('source'):
            raise ValueError('Official/user logo requires source')
        _check_logo(resolve(run, logo.get('file', '')))
    elif mode == 'placeholder':
        from PIL import Image, ImageDraw, ImageFont
        def wordmark(path, color):
            font = ImageFont.truetype('C:/Windows/Fonts/malgunbd.ttf', 34)
            probe = ImageDraw.Draw(Image.new('RGBA', (1, 1)))
            box = probe.textbbox((0, 0), institution, font=font)
            image = Image.new('RGBA', (max(40, box[2]-box[0]+24), 52), (0,0,0,0))
            ImageDraw.Draw(image).text((12, 5), institution, font=font, fill=color)
            image.save(path)
        logo['file'] = logo.get('file') or 'logo.png'
        logo['light_file'] = logo.get('light_file') or 'logo-dark.png'
        wordmark(resolve(run, logo['file']), (255,255,255,255))
        wordmark(resolve(run, logo['light_file']), (20,38,51,255))
        logo['source'] = 'generated institution wordmark'
    else:
        logo['file'] = ''
        logo.pop('light_file', None)
        logo['source'] = 'none'
    save(run/'request.json', req)
    files = brand(run, req)
    with zipfile.ZipFile(run/'template.pptx', 'w', zipfile.ZIP_DEFLATED) as z:
        for n, data in files.items(): z.writestr(n, data)
    logo_sha = sha(resolve(run, req['logo']['file'])) if mode != 'none' else None
    save(run/'baseline.json', {'request_sha256': sha(run/'request.json'), 'logo_sha256': logo_sha,
                               'template_sha256': sha(run/'template.pptx'), 'original_sha256': sha(jt.DEFAULT)})
    print('template.pptx ready')

def inspect_plan(plan):
    issues = []
    slides = plan.get('slides', [])
    if not slides:
        issues.append({'rule': 'PLAN-01', 'message': 'No planned slides'})
    for i, s in enumerate(slides, 1):
        reuse = s.get('reuse')
        if reuse is not None and not isinstance(reuse, dict):
            issues.append({'rule': 'PLAN-07', 'page': i, 'message': 'reuse must be an object'})
            reuse = {}
        mode = reuse.get('mode') if isinstance(reuse, dict) else None
        if mode is not None and mode != 'reuse':
            issues.append({'rule': 'PLAN-07', 'page': i, 'message': 'reuse.mode must be reuse'})
        layout = s.get('layout')
        if not isinstance(layout, int) or isinstance(layout, bool) or layout not in range(1, 14):
            issues.append({'rule': 'PLAN-04', 'page': i, 'message': 'layout must be 1..13'})
        if mode == 'reuse':
            if not isinstance(reuse.get('source'),str) or not reuse.get('source','').strip():
                issues.append({'rule':'PLAN-07','page':i,'message':'reuse must record its source'})
            if not isinstance(reuse.get('shapes'),list) or not reuse.get('shapes'):
                issues.append({'rule':'PLAN-07','page':i,'message':'reuse must record candidate shape IDs'})
            if not isinstance(reuse.get('fit'),dict):
                issues.append({'rule':'PLAN-07','page':i,'message':'reuse must record its fit transform'})
        if s.get('role') not in ROLES:
            issues.append({'rule': 'PLAN-05', 'page': i, 'message': 'role must be one of ' + '/'.join(ROLES)})
        if s.get('role') == 'cover' and i != 1:
            issues.append({'rule': 'PLAN-06', 'page': i, 'message': 'cover role only on page 1'})
    return issues

def plan_check(run):
    plan = load(run/'plan.json'); issues = inspect_plan(plan)
    try:
        r = visual_retrieval.collect(run)
        retrieved = {'files': len(r['files']), 'searched_decks': r.get('searched_decks'),
                     'folder': str(run/'retrieved-visuals')}
    except Exception as exc:
        issues.append({'rule': 'RETRIEVE-01', 'message': str(exc)})
        retrieved = {'error': str(exc)}
    result = {'status': 'BLOCKED' if issues else 'PASS', 'issues': issues, 'retrieved': retrieved}
    save(run/'plan-report.json', result)
    return result


def inspect(run, visual=True):
    issues = []
    def fail(code, message, page=None):
        issues.append({'rule': code, 'message': message, 'page': page})

    req = load(run/'request.json'); plan = load(run/'plan.json'); asset = load(run/'assets.json'); deck = run/'candidate.pptx'
    try:
        issues.extend(visual_retrieval.inspect(run))
    except Exception as exc:
        fail('RETRIEVE-01', str(exc))
    for key in questions(req): fail('INPUT-01', 'Missing answer: ' + key)
    logo = req.get('logo', {})
    mode = logo.get('mode')
    if mode not in ('official', 'user', 'placeholder', 'none'):
        fail('LOGO-01', 'logo.mode must be official, user, placeholder or none')
    elif logo.get('institution') != req.get('answers', {}).get('institution', {}).get('value'):
        fail('LOGO-01', 'Logo institution differs from the request')
    elif mode != 'none':
        try:
            _check_logo(resolve(run, logo.get('file', '')))
        except Exception as exc:
            fail('LOGO-01', str(exc))
        if mode in ('official', 'user') and not logo.get('source'):
            fail('LOGO-01', 'Official/user logo requires source')
    try:
        if float(req['answers']['duration_minutes']['value']) <= 0: raise ValueError()
    except (KeyError, TypeError, ValueError):
        fail('INPUT-02', 'Education duration must be positive minutes')
    slides = plan.get('slides', [])
    for i in inspect_plan(plan): issues.append(i)
    if not deck.exists():
        fail('FILE-01', 'candidate.pptx missing'); return {'status': 'BLOCKED', 'issues': issues}

    try:
        expected = brand(run, req); template = jt.read(run/'template.pptx'); files = jt.read(deck)
        if set(expected) != set(template) or any(jt.signature(v, n) != jt.signature(template[n], n) for n, v in expected.items()):
            fail('BASE-01', 'Working template is not the permitted course/logo adaptation')
        ordered=inspect_package_graph(files,fail)
        for err in jt.verify(files, expected): fail('MASTER-01', err)
        inspect_master_graph(files,fail)
    except Exception as exc:
        fail('BASE-02', str(exc)); return {'status': 'BLOCKED', 'issues': issues}

    pres = jt.xml(files['ppt/presentation.xml'])
    if len(ordered) != len(slides): fail('PLAN-02', 'Actual slide count differs from plan')
    specified = req.get('fixed_slide_count')
    if specified is not None and len(ordered) != specified: fail('PLAN-03', 'Explicit user slide count not met')
    try:
        live_graph=reuse_slide.reachable_parts(files,ordered)
        unreachable_extra=sorted(name for name in files if name not in expected and name not in live_graph)
        if unreachable_extra:
            fail('PACKAGE-01','Non-baseline package parts must be reachable from a live slide: '+', '.join(unreachable_extra[:8]))
    except Exception as exc:
        fail('PACKAGE-01','Could not validate live package reachability: '+str(exc))

    allowed_colors = {'000000', 'FFFFFF', '333333', '404040', '666666', '808080', 'BFBFBF'}
    for n, data in expected.items():
        if n.endswith('.xml') and ('slideLayouts/' in n or 'theme/' in n):
            allowed_colors.update(c.get('val', '').upper() for c in jt.xml(data).findall('.//a:srgbClr', NS))

    db_path = library_db()
    for i, part in enumerate(ordered, 1):
        item = slides[i-1] if i <= len(slides) else {}
        e = jt.xml(files[part]); rels = list(jt.xml(files[jt.relpart(part)]))
        lr = [r for r in rels if r.get('Type') == R + '/slideLayout']
        style_exempt_ids = inspect_reuse_record(run, item, part, files, db_path, fail, i)
        issues.extend(shape_labels.inspect_slide(e, i, item.get('label_exclusions', [])))
        issues.extend(teaching_visuals.inspect_slide(e, i, item))
        match = re.search(r'slideLayout(\d+)\.xml', lr[0].get('Target')) if lr else None
        actual = int(match[1]) if match else None
        if actual != item.get('layout'): fail('LAYOUT-01', 'Layout differs from planned ID', i)
        text = ''.join(t.text or '' for t in e.findall('.//a:t', NS))
        if not text.strip(): fail('CONTENT-01', 'Empty or fully rasterized slide', i)
        issues.extend(textfit_issues(e, i))
        inherited_text = visible_text(files, part)
        direct_text = '\n'.join(node.text or '' for node in e.findall('.//a:t', NS))
        for value in (COURSE_LABEL, 'AI챔피언 전문인재 보수교육 in Jeju'):
            if value != req.get('title') and value in inherited_text:
                fail('RESIDUAL-01', 'Old course metadata remains: ' + value[:80], i)
        for value in dict.fromkeys(str(row) for row in item.get('reuse', {}).get('band_texts', []) if row):
            if value != req.get('title') and value in direct_text:
                fail('RESIDUAL-01', 'Old source band text remains: ' + value[:80], i)
        for src in item.get('sources', []):
            if src.get('kind') == 'internal' and db_path.exists():
                with sqlite3.connect('file:' + db_path.resolve().as_posix() + '?mode=ro', uri=True) as db:
                    if not db.execute('SELECT 1 FROM slides WHERE id=?', (src.get('slide_id', ''),)).fetchone():
                        fail('SOURCE-03', 'Internal slide ID not in library: ' + str(src.get('slide_id')), i)
        kinds = {'table': bool(e.findall('.//a:tbl', NS)),
                 'chart': any(r.get('Type') == R + '/chart' for r in rels),
                 'diagram': bool(e.findall('.//p:cxnSp', NS) or e.findall('.//p:grpSp', NS))}
        for kind in item.get('required_native', []):
            if not kinds.get(kind, False): fail('EDIT-01', 'Required native object missing: ' + kind, i)
        is_cover = i == 1 and item.get('role') == 'cover'
        tree=e.find('p:cSld/p:spTree',NS)
        for _,b,label,path in shape_boxes(tree):
            x,y,w,height=b
            if min(x,y,w,height)<-1000 or x+w>12193000 or y+height>6859000:
                suffix=' (group '+('/'.join(path))+')' if path else ''
                fail('BOUNDS-01','Object outside canvas: '+str(label)+suffix,i)
        for sp in e.findall('p:cSld/p:spTree/*', NS):
            if sp.tag.split('}')[-1] in ('nvGrpSpPr', 'grpSpPr'): continue
            b = bbox(sp); name = sp.find('.//p:cNvPr', NS); label = name.get('name', '') if name is not None else '?'
            if b:
                x, y, w, h = b
                if not is_cover:
                    if y < 440000 and sp.find('.//p:ph', NS) is None: fail('HEADER-01', 'Body object intrudes into header: ' + label, i)
                    if w*h > 12192000*6858000*0.9 and sp.tag == '{'+P+'}pic': fail('RASTER-01', 'Full-slide image outside cover', i)
            shape_id,_=reuse_slide.direct_shape_identity(sp)
            style_exempt=str(shape_id) in style_exempt_ids
            for f in ([] if style_exempt else sp.findall('.//a:latin', NS) + sp.findall('.//a:ea', NS)):
                face = f.get('typeface', '')
                if face and not (face.startswith('+') or face.startswith('Noto Sans KR') or face.startswith('페이퍼로지') or face.startswith('Paperlogy')):
                    fail('FONT-01', 'Unexpected font: ' + face, i)
            for color in ([] if style_exempt else sp.findall('.//a:rPr/a:solidFill/a:srgbClr', NS)):
                if color.get('val', '').upper() not in allowed_colors: fail('COLOR-01', 'Text color outside template palette', i)
    # Repeated composition remains a review prompt rather than a release blocker.

    # Images are tracked by bytes, so assets.json cannot lie about what is in the deck.
    documented = {}
    for a in asset.get('images', []):
        try:
            p = resolve(run, a['file']); h = sha(p)
            if a.get('sha256') != h: fail('ASSET-01', 'Asset changed: ' + a['file'])
            if not a.get('source'): fail('ASSET-02', 'Asset source missing: ' + a['file'])
            documented[h] = a
        except Exception as exc:
            fail('ASSET-03', str(exc))
    for i, part in enumerate(ordered, 1):
        cover_found = False
        for r in jt.xml(files[jt.relpart(part)]):
            if r.get('Type') == R + '/image' and r.get('TargetMode') != 'External':
                n = jt.target(part, r); h = blobsha(files[n])
                if h not in documented: fail('ASSET-04', 'Unrecorded slide image: ' + n, i)
                elif documented[h].get('role') == 'cover': cover_found = True
        if i <= len(slides) and slides[i-1].get('role') == 'cover' and not cover_found:
            fail('COVER-02', 'Cover slide needs its recorded cover image', i)

    binding = {'deck': sha(deck), 'request': sha(run/'request.json'), 'plan': sha(run/'plan.json'),
               'assets': sha(run/'assets.json'), 'template': sha(run/'template.pptx')}
    if visual:
        rp = run/'visual-review.json'
        if not rp.exists(): fail('VISUAL-01', 'Visual review missing (run review after rendering)')
        else:
            review = load(rp)
            if review.get('binding') != binding: fail('VISUAL-02', 'Review belongs to different inputs or PPTX')
            entries = review.get('pages', [])
            if [p.get('page') for p in entries] != list(range(1, len(ordered)+1)): fail('VISUAL-03', 'Review must cover every page once')
            for entry in entries:
                i = entry.get('page'); png = resolve(run, entry.get('render', ''))
                if not png.is_file() or entry.get('sha256') != sha(png): fail('VISUAL-04', 'Missing or changed page render', i)
                keys = CHECKS + (COVER_CHECKS if i == 1 and slides and slides[0].get('role') == 'cover' else [])
                for c in keys:
                    if entry.get('checks', {}).get(c, {}).get('status') != 'pass': fail('VISUAL-05', 'Visual check not passed: ' + c, i)
            manifest = run/'renders/manifest.json'
            if not manifest.exists() or load(manifest).get('deck_sha256') != binding['deck']: fail('RENDER-01', 'No render for current deck')
            elif [{k: p.get(k) for k in ('page', 'render', 'sha256')} for p in entries] != load(manifest).get('pages'): fail('RENDER-02', 'Review pages differ from render manifest')
    return {'status': 'PASS' if not issues else 'BLOCKED', 'binding': binding, 'issues': issues}


def check(run, visual=True):
    try:
        report = inspect(run, visual)
    except Exception as exc:
        report = {'status': 'BLOCKED', 'issues': [{'rule': 'INVALID-01', 'message': str(exc)}]}
    save(run/'report.json', report)
    return report

def review(run):
    report = check(run, False)
    if report['status'] != 'PASS':
        raise ValueError('Fix structural failures first; see report.json')
    manifest = load(run/'renders/manifest.json'); plan = load(run/'plan.json')
    if manifest['deck_sha256'] != report['binding']['deck']:
        raise ValueError('Stale render: render candidate.pptx again')
    review = {'binding': report['binding'], 'pages': []}
    for p in manifest['pages']:
        keys = CHECKS + (COVER_CHECKS if p['page'] == 1 and plan['slides'] and plan['slides'][0].get('role') == 'cover' else [])
        review['pages'].append(dict(p, checks={k: {'status': 'pending'} for k in keys}))
    save(run/'visual-review.json', review)
    print('visual-review.json created: inspect every render PNG, then set each status to pass')

def release(run):
    report = check(run, True)
    if report['status'] != 'PASS':
        raise ValueError('Delivery blocked; see report.json')
    dest = run/'delivery'; dest.mkdir(exist_ok=False)
    shutil.copy2(run/'candidate.pptx', dest/'lecture.pptx')
    print('delivered', dest/'lecture.pptx')


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('command', choices=['init', 'prepare', 'plan-check', 'check', 'review', 'release'])
    p.add_argument('--run', type=pathlib.Path, required=True)
    p.add_argument('--structural-only', action='store_true')
    a = p.parse_args(); run = a.run.resolve()
    if a.command == 'init': init(run)
    elif a.command == 'prepare': prepare(run)
    elif a.command == 'plan-check':
        result = plan_check(run)
        print(json.dumps(result, ensure_ascii=False, indent=2)); return int(result['status'] != 'PASS')
    elif a.command == 'check':
        result = check(run, not a.structural_only)
        print(json.dumps(result, ensure_ascii=False, indent=2)); return int(result['status'] != 'PASS')
    elif a.command == 'review': review(run)
    else: release(run)
    return 0

if __name__ == '__main__':
    try:
        if hasattr(sys.stdout, 'reconfigure'): sys.stdout.reconfigure(encoding='utf-8')
        sys.exit(main())
    except Exception as exc:
        print(str(exc), file=sys.stderr); sys.exit(1)
