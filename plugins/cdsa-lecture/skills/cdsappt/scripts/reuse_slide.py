"""Copy editable source shapes into one CDSA-master slide.

Use ``--all-body-shapes`` for the source body or repeat ``--include-shape``
for selected top-level objects.  The copier resolves source theme tokens,
remaps relationships and connector endpoints, fits the result into the
destination layout's safe area, and records provenance in ``plan.json``.
"""
import argparse, copy, hashlib, json, math, pathlib, posixpath, re, sqlite3, sys, zipfile
import jeju_template as jt

P, A, R, PKG, CT = jt.P, 'http://schemas.openxmlformats.org/drawingml/2006/main', jt.R, jt.PKG, jt.CT
NS = {'p': P, 'a': A}
REL_ATTRS = ('embed', 'link', 'id', 'dm', 'lo', 'qs', 'cs', 'pict')
SKIP_REL_TYPES = {R + '/notesSlide', R + '/comments', R + '/slide'}  # never drag notes/comments/other slides along


def sha(data):
    return hashlib.sha256(data).hexdigest()

def file_sha(path):
    with pathlib.Path(path).open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()

def load(p):
    return json.loads(p.read_text(encoding='utf-8-sig'))

def save(p, obj):
    p.write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding='utf-8')

def ordered_slides(files):
    pres = jt.xml(files['ppt/presentation.xml'])
    rr = {r.get('Id'): r for r in jt.xml(files['ppt/_rels/presentation.xml.rels'])}
    return [jt.target('ppt/presentation.xml', rr[s.get('{'+R+'}id')]) for s in pres.find('p:sldIdLst', NS)]

def slide_size(files):
    size = jt.xml(files['ppt/presentation.xml']).find('p:sldSz', NS)
    if size is None or size.get('cx') is None or size.get('cy') is None:
        raise ValueError('Presentation is missing its slide size')
    return int(size.get('cx')), int(size.get('cy'))

def export_from_library(sid, run):
    import slide_store
    import visual_retrieval
    requested = sid.removeprefix('ssj-slide://')
    with sqlite3.connect('file:' + slide_store.DB.resolve().as_posix() + '?mode=ro', uri=True) as db:
        if requested.startswith('g-'):
            row = db.execute('SELECT id FROM slides WHERE group_id=? ORDER BY id LIMIT 1', (requested,)).fetchone()
            if not row:
                raise ValueError('Unknown group ID: ' + requested)
            resolved = row[0]
        else:
            row = db.execute('SELECT id FROM slides WHERE id=?', (requested,)).fetchone()
            if not row:
                raise ValueError('Unknown slide ID: ' + requested)
            resolved = row[0]
        out = run / 'retrieved-visuals' / (resolved + '.pptx')
        # Rebuild from the DB even when a file with this slide ID already exists.
        # visual_retrieval writes a deterministic zip, so its hash is stable later.
        visual_retrieval.refresh_native(db, resolved, out)
    return out, resolved

def xfrm_of(sp):
    for path in ('p:spPr/a:xfrm', 'p:xfrm', 'p:grpSpPr/a:xfrm'):
        x = sp.find(path, NS)
        if x is not None and x.find('a:off', NS) is not None:
            return x
    return None

def ph_key(sp):
    ph = sp.find('.//p:nvPr/p:ph', NS)
    return None if ph is None else (ph.get('type', 'body'), ph.get('idx'))

def layout_and_master(files, slide_part):
    rels = jt.xml(files[jt.relpart(slide_part)])
    layout = next((jt.target(slide_part, r) for r in rels if r.get('Type') == R + '/slideLayout'), None)
    master = None
    if layout and jt.relpart(layout) in files:
        lrels = jt.xml(files[jt.relpart(layout)])
        master = next((jt.target(layout, r) for r in lrels if r.get('Type') == R + '/slideMaster'), None)
    return layout, master

def reachable_parts(files, roots):
    """Return the relationship closure used to prune stale imported parts."""
    todo=list(roots);seen=set()
    while todo:
        part=todo.pop()
        if part in seen:continue
        if part not in files:raise ValueError('Missing relationship target: '+part)
        seen.add(part);rp=jt.relpart(part)
        if rp not in files:continue
        seen.add(rp)
        for rel in jt.xml(files[rp]):
            if rel.get('TargetMode')=='External' or rel.get('Type') in SKIP_REL_TYPES:continue
            todo.append(jt.target(part,rel))
    return seen

def direct_shape_identity(shape):
    """Return the non-visual identity of one top-level slide shape."""
    nv=shape.find('./*/p:cNvPr',NS)
    if nv is None:nv=shape.find('.//p:cNvPr',NS)
    return (nv.get('id') if nv is not None else None,
            nv.get('name','') if nv is not None else '')


def shape_hash(shape):
    """Stable copied-shape hash that permits literal copy edits, not style edits."""
    clone=copy.deepcopy(shape)
    for node in clone.findall('.//a:t',NS):node.text=''
    return sha(jt.serialize(clone))


def prune_unreachable_imports(files, baseline):
    """Remove stale imported dependencies after a page is replaced."""
    ordered=ordered_slides(files);ordered_set=set(ordered)
    # A presentation relationship that is not represented by sldIdLst is an orphan
    # page. Remove it before computing reachability so a later reuse cannot preserve
    # hidden slide parts accidentally.
    prels=jt.xml(files['ppt/_rels/presentation.xml.rels'])
    for rel in list(prels):
        if rel.get('Type')==R+'/slide' and jt.target('ppt/presentation.xml',rel) not in ordered_set:
            prels.remove(rel)
    files['ppt/_rels/presentation.xml.rels']=jt.serialize(prels)
    live_graph=reachable_parts(files,ordered)
    # create() removes source sample slides from the working deck. Do not treat an
    # imported slide that happens to reuse one of those part names as baseline-owned.
    baseline_parts={n for n in baseline if not n.startswith('ppt/slides/')}
    needed_masters={n for n in live_graph if re.fullmatch(r'ppt/slideMasters/[^/]+\.xml',n)}
    pres=jt.xml(files['ppt/presentation.xml']);prels=jt.xml(files['ppt/_rels/presentation.xml.rels'])
    removed_rids=set()
    for rel in list(prels):
        if rel.get('Type')!=R+'/slideMaster':continue
        target=jt.target('ppt/presentation.xml',rel)
        if target not in needed_masters:
            removed_rids.add(rel.get('Id'));prels.remove(rel)
    masters=pres.find('p:sldMasterIdLst',NS)
    if masters is not None:
        for entry in list(masters):
            if entry.get('{'+R+'}id') in removed_rids:masters.remove(entry)
    files['ppt/presentation.xml']=jt.serialize(pres)
    files['ppt/_rels/presentation.xml.rels']=jt.serialize(prels)
    removed={n for n in files if n not in baseline_parts and n not in live_graph}
    for name in removed:files.pop(name,None)
    ct=jt.xml(files['[Content_Types].xml'])
    for entry in list(ct):
        if entry.get('PartName','').lstrip('/') in removed:ct.remove(entry)
    files['[Content_Types].xml']=jt.serialize(ct)
    return sorted(removed)

def placeholder_table(files, slide_part):
    table = []
    for part in [p for p in layout_and_master(files, slide_part) if p and p in files]:
        for sp in jt.xml(files[part]).find('p:cSld/p:spTree', NS):
            k = ph_key(sp); x = xfrm_of(sp)
            if k and x is not None:
                table.append((k, copy.deepcopy(x)))
    return table

def resolve_xfrm(key, table):
    typ, idx = key
    if idx is not None:
        for k, x in table:
            if k[1] == idx:
                return x
    for k, x in table:
        if k[0] == typ:
            return x
    return None

def strip_placeholders(shape, table, warnings):
    for sp in [shape] + shape.findall('.//p:sp', NS) + shape.findall('.//p:pic', NS) + shape.findall('.//p:graphicFrame', NS):
        key = ph_key(sp)
        if not key:
            continue
        if xfrm_of(sp) is None:
            x = resolve_xfrm(key, table)
            if x is None:
                warnings.append(f'placeholder {key} has no geometry in source layout/master; shape may be unpositioned')
            elif sp.tag == '{'+P+'}graphicFrame':
                if sp.find('p:xfrm', NS) is None:
                    sp.insert(1, copy.deepcopy(x))
            else:
                sppr = sp.find('p:spPr', NS)
                if sppr is None:
                    sppr = jt.ET.SubElement(sp, '{'+P+'}spPr')
                sppr.insert(0, copy.deepcopy(x))
        nvpr = sp.find('.//p:nvPr', NS)
        ph = nvpr.find('p:ph', NS) if nvpr is not None else None
        if ph is not None:
            nvpr.remove(ph)


def _theme_context(files, slide_part):
    """Resolve source theme fonts, colors and the master color map."""
    layout, master = layout_and_master(files, slide_part)
    theme = None
    if master and jt.relpart(master) in files:
        theme = next((jt.target(master, r) for r in jt.xml(files[jt.relpart(master)])
                      if r.get('Type') == R + '/theme'), None)
    fonts = {}
    colors = {}
    if theme and theme in files:
        root = jt.xml(files[theme])
        for prefix, path in (('mj', './/a:fontScheme/a:majorFont'), ('mn', './/a:fontScheme/a:minorFont')):
            family = root.find(path, NS)
            if family is not None:
                for suffix, tag in (('lt', 'latin'), ('ea', 'ea'), ('cs', 'cs')):
                    node = family.find('a:' + tag, NS)
                    if node is not None and node.get('typeface'):
                        fonts[f'+{prefix}-{suffix}'] = node.get('typeface')
        scheme = root.find('.//a:clrScheme', NS)
        if scheme is not None:
            for entry in scheme:
                child = next(iter(entry), None)
                if child is None:
                    continue
                value = child.get('lastClr') or child.get('val')
                if value and re.fullmatch(r'[0-9A-Fa-f]{6}', value):
                    colors[entry.tag.split('}')[-1]] = value.upper()
    aliases = {}
    if master and master in files:
        cmap = jt.xml(files[master]).find('p:clrMap', NS)
        if cmap is not None:
            aliases.update(cmap.attrib)
    return fonts, colors, aliases, layout, master


def _placeholder_shapes(files, layout, master, key):
    result = []
    for part in (layout, master):
        if not part or part not in files:
            continue
        tree = jt.xml(files[part]).find('p:cSld/p:spTree', NS)
        if tree is None:
            continue
        exact = [shape for shape in tree if ph_key(shape) == key]
        typed = [shape for shape in tree if ph_key(shape) and ph_key(shape)[0] == key[0]]
        result.extend(exact or typed)
    return result


def _level_properties(shape, level):
    if shape is None:
        return None
    return shape.find(f'p:txBody/a:lstStyle/a:lvl{level + 1}pPr', NS)


def _merge_missing(target, source):
    if target is None or source is None:
        return
    for key, value in source.attrib.items():
        target.attrib.setdefault(key, value)
    existing = {child.tag for child in target}
    for child in source:
        if child.tag not in existing:
            target.append(copy.deepcopy(child))
            existing.add(child.tag)


_PPR_ORDER = {
    name: index for index, name in enumerate((
        'lnSpc', 'spcBef', 'spcAft', 'buClrTx', 'buClr', 'buSzTx',
        'buSzPct', 'buSzPts', 'buFontTx', 'buFont', 'buNone', 'buAutoNum',
        'buChar', 'buBlip', 'tabLst', 'defRPr', 'extLst'
    ))
}

_RPR_ORDER = {
    name: index for index, name in enumerate((
        'ln', 'noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill',
        'grpFill', 'effectLst', 'effectDag', 'highlight', 'uLnTx', 'uLn',
        'uFillTx', 'uFill', 'latin', 'ea', 'cs', 'sym', 'hlinkClick',
        'hlinkMouseOver', 'rtl', 'extLst'
    ))
}


def _schema_order(node, order):
    """Keep DrawingML property children in the sequence Office validates."""
    children = list(node)
    ranked = sorted(enumerate(children),
                    key=lambda item: (order.get(item[1].tag.split('}')[-1], 10_000), item[0]))
    node[:] = [child for _, child in ranked]


def materialize_source_style(shape, files, slide_part, context=None):
    """Keep source appearance after removing its placeholder/master binding."""
    fonts, colors, aliases, layout, master = context or _theme_context(files, slide_part)
    key = ph_key(shape)
    inherited = _placeholder_shapes(files, layout, master, key) if key else []
    master_tree = jt.xml(files[master]) if master and master in files else None
    ph_type = key[0] if key else 'body'
    style_name = 'titleStyle' if ph_type in ('title', 'ctrTitle') else 'bodyStyle' if ph_type in ('body', 'obj') else 'otherStyle'
    master_style = master_tree.find('p:txStyles/p:' + style_name, NS) if master_tree is not None else None
    # A group or ordinary authored shape has no placeholder inheritance to
    # materialize.  Adding master paragraph children to such shapes can also
    # violate DrawingML's strict child ordering, so only expand true
    # placeholders.  Theme tokens are still resolved below for every shape.
    if key:
        for paragraph in shape.findall('.//a:p', NS):
            ppr = paragraph.find('a:pPr', NS)
            if ppr is None:
                ppr = jt.ET.Element('{' + A + '}pPr')
                paragraph.insert(0, ppr)
            level = max(0, min(8, int(ppr.get('lvl', '0'))))
            sources = [_level_properties(item, level) for item in inherited]
            if master_style is not None:
                sources.append(master_style.find(f'a:lvl{level + 1}pPr', NS))
            for source in sources:
                _merge_missing(ppr, source)
            defr = ppr.find('a:defRPr', NS)
            if defr is None:
                defr = jt.ET.SubElement(ppr, '{' + A + '}defRPr')
            for source in sources:
                _merge_missing(defr, source.find('a:defRPr', NS) if source is not None else None)
            if defr.get('sz') is None:
                defr.set('sz', '1800')
            _schema_order(defr, _RPR_ORDER)
            _schema_order(ppr, _PPR_ORDER)
    # Resolve theme font tokens before the source master is discarded.
    for node in shape.findall('.//a:latin', NS) + shape.findall('.//a:ea', NS) + shape.findall('.//a:cs', NS):
        token = node.get('typeface', '')
        if token in fonts and fonts[token]:
            node.set('typeface', fonts[token])
    # Resolve source scheme colors to RGB.  Preserve color transforms such as
    # alpha/lumMod by retaining the child nodes on the replacement element.
    for node in list(shape.findall('.//a:schemeClr', NS)):
        name = aliases.get(node.get('val'), node.get('val'))
        value = colors.get(name)
        if not value:
            continue
        parent = next((p for p in shape.iter() if node in list(p)), None)
        if parent is None:
            continue
        replacement = jt.ET.Element('{' + A + '}srgbClr', {'val': value})
        for child in node:
            replacement.append(copy.deepcopy(child))
        parent.insert(list(parent).index(node), replacement)
        parent.remove(node)


def _raw_box(shape):
    xf = xfrm_of(shape)
    if xf is None:
        return None
    off = xf.find('a:off', NS); ext = xf.find('a:ext', NS)
    if off is None or ext is None:
        return None
    x, y = int(off.get('x')), int(off.get('y'))
    w, h = int(ext.get('cx')), int(ext.get('cy'))
    angle = math.radians(int(xf.get('rot', '0')) / 60000)
    if not angle:
        return x, y, w, h
    cx, cy = x + w / 2, y + h / 2
    points = []
    for px, py in ((x, y), (x + w, y), (x, y + h), (x + w, y + h)):
        dx, dy = px - cx, py - cy
        points.append((cx + dx * math.cos(angle) - dy * math.sin(angle),
                       cy + dx * math.sin(angle) + dy * math.cos(angle)))
    left = min(p[0] for p in points); top = min(p[1] for p in points)
    return round(left), round(top), round(max(p[0] for p in points) - left), round(max(p[1] for p in points) - top)


def _scale_text_and_table(shape, scale):
    for node in shape.findall('.//a:rPr', NS) + shape.findall('.//a:defRPr', NS) + shape.findall('.//a:endParaRPr', NS):
        if node.get('sz'):
            value = max(900, int(round((int(node.get('sz')) * scale) / 50) * 50))
            node.set('sz', str(value))
    for node in shape.findall('.//a:gridCol', NS):
        if node.get('w'):
            node.set('w', str(max(1, round(int(node.get('w')) * scale))))
    for node in shape.findall('.//a:tr', NS):
        if node.get('h'):
            node.set('h', str(max(1, round(int(node.get('h')) * scale))))


def _normalize_table_labels(shape):
    """Apply the retained teaching-table rule to short headers and row labels."""
    table = shape.find('.//a:tbl', NS)
    if table is None:
        return
    for row_index, row in enumerate(table.findall('a:tr', NS)):
        for column_index, cell in enumerate(row.findall('a:tc', NS)):
            text = ''.join(node.text or '' for node in cell.findall('.//a:t', NS)).strip()
            if not text or len(text) > 20 or (row_index != 0 and column_index != 0):
                continue
            props = cell.find('a:tcPr', NS)
            if props is None:
                props = jt.ET.SubElement(cell, '{' + A + '}tcPr')
            props.set('anchor', 'ctr')
            for paragraph in cell.findall('a:txBody/a:p', NS):
                ppr = paragraph.find('a:pPr', NS)
                if ppr is None:
                    ppr = jt.ET.Element('{' + A + '}pPr')
                    paragraph.insert(0, ppr)
                ppr.set('algn', 'ctr')


def _fit_shapes(imported, box):
    boxes = [_raw_box(shape) for shape, _ in imported]
    boxes = [value for value in boxes if value]
    if not boxes:
        raise ValueError('Selected source shapes have no usable geometry')
    ux = min(v[0] for v in boxes); uy = min(v[1] for v in boxes)
    ur = max(v[0] + v[2] for v in boxes); ub = max(v[1] + v[3] for v in boxes)
    bx, by, bw, bh = box
    scale = min(1.0, bw / max(1, ur - ux), bh / max(1, ub - uy))
    dx = bx - ux * scale + (bw - (ur - ux) * scale) / 2
    dy = by - uy * scale + (bh - (ub - uy) * scale) / 2
    for shape, _ in imported:
        xf = xfrm_of(shape)
        if xf is None:
            continue
        off = xf.find('a:off', NS); ext = xf.find('a:ext', NS)
        off.set('x', str(round(int(off.get('x')) * scale + dx)))
        off.set('y', str(round(int(off.get('y')) * scale + dy)))
        ext.set('cx', str(max(1, round(int(ext.get('cx')) * scale))))
        ext.set('cy', str(max(1, round(int(ext.get('cy')) * scale))))
        _scale_text_and_table(shape, scale)
    return {'scale': round(scale, 6), 'dx': round(dx), 'dy': round(dy),
            'box': {'x': bx, 'y': by, 'w': bw, 'h': bh}}

def source_shape_catalog(src,src_part):
    result=[]
    tree=jt.xml(src[src_part]).find('p:cSld/p:spTree',NS)
    for shape in tree:
        if shape.tag.split('}')[-1] in ('nvGrpSpPr','grpSpPr'):continue
        nv=shape.find('.//p:cNvPr',NS);box=None;x=xfrm_of(shape)
        if x is not None:
            off=x.find('a:off',NS);ext=x.find('a:ext',NS)
            if off is not None and ext is not None:box=[int(off.get('x')),int(off.get('y')),int(ext.get('cx')),int(ext.get('cy'))]
        result.append({'id':nv.get('id') if nv is not None else None,
                       'name':nv.get('name','') if nv is not None else '',
                       'kind':shape.tag.split('}')[-1],'box':box,
                       'text':''.join(t.text or '' for t in shape.findall('.//a:t',NS))[:160]})
    return result


class Importer:
    """Copy parts from src package into dst, following relationships, renaming on collision."""
    def __init__(self, src, dst):
        self.src, self.dst = src, dst
        self.copied = {}
        self.src_ct = jt.xml(src['[Content_Types].xml'])
        self.dst_ct = jt.xml(dst['[Content_Types].xml'])

    def _unique(self, name):
        if name not in self.dst:
            return name
        # A structural XML part can have identical bytes while its .rels points to a
        # different layout, master or theme in another source deck. Reusing that path
        # would overwrite the first import's relationship graph on a later reuse call.
        # Only self-contained identical binary/XML parts are safe to deduplicate.
        rp = jt.relpart(name)
        # PowerPoint rejects a second imported master that shares even a byte-identical
        # theme part with another master. Master/layout/theme XML is therefore isolated
        # per import; identical leaf media and other self-contained parts may deduplicate.
        design_xml = bool(re.fullmatch(r'ppt/(slideMasters|slideLayouts|theme)/[^/]+\.xml', name))
        if self.dst[name] == self.src[name] and not design_xml and rp not in self.src and rp not in self.dst:
            return name
        stem, ext = posixpath.splitext(name); n = 1
        while f'{stem}-r{n}{ext}' in self.dst:
            n += 1
        return f'{stem}-r{n}{ext}'

    def _content_type(self, src_name, dst_name):
        for el in self.src_ct:
            if el.tag == '{'+CT+'}Override' and el.get('PartName', '').lstrip('/') == src_name:
                if not any(o.get('PartName', '').lstrip('/') == dst_name for o in self.dst_ct):
                    jt.ET.SubElement(self.dst_ct, '{'+CT+'}Override', PartName='/' + dst_name, ContentType=el.get('ContentType'))
                return
        ext = posixpath.splitext(dst_name)[1].lstrip('.').lower()
        if ext and not any(d.tag == '{'+CT+'}Default' and d.get('Extension', '').lower() == ext for d in self.dst_ct):
            ctype = next((d.get('ContentType') for d in self.src_ct if d.tag == '{'+CT+'}Default' and d.get('Extension', '').lower() == ext), None)
            if ctype:
                jt.ET.SubElement(self.dst_ct, '{'+CT+'}Default', Extension=ext, ContentType=ctype)

    def copy_part(self, name):
        if name in self.copied:
            return self.copied[name]
        if name not in self.src:
            raise ValueError('Source package is missing part ' + name)
        new = self._unique(name)
        self.copied[name] = new
        self.dst[new] = self.src[name]
        self._content_type(name, new)
        rp = jt.relpart(name)
        if rp in self.src:
            rels = jt.xml(self.src[rp])
            for r in list(rels):
                if r.get('Type') in SKIP_REL_TYPES:
                    rels.remove(r); continue
                if r.get('TargetMode') == 'External':
                    continue
                child = self.copy_part(jt.target(name, r))
                r.set('Target', posixpath.relpath(child, posixpath.dirname(new)))
            self.dst[jt.relpart(new)] = jt.serialize(rels)
        return new

    def finish(self):
        self.dst['[Content_Types].xml'] = jt.serialize(self.dst_ct)


def mode_reuse(src, dst, src_part, dst_part, box, include_shapes=None, all_body_shapes=False):
    """Copy selected editable source shapes into one CDSA layout and fit them."""
    warnings = []
    src_tree = jt.xml(src[src_part]).find('p:cSld/p:spTree', NS)
    src_rels = {r.get('Id'): r for r in jt.xml(src[jt.relpart(src_part)])}
    dst_xml = jt.xml(dst[dst_part]); dst_tree = dst_xml.find('p:cSld/p:spTree', NS)
    dst_rels = jt.xml(dst[jt.relpart(dst_part)])
    for sp in list(dst_tree):
        if sp.tag.split('}')[-1] not in ('nvGrpSpPr', 'grpSpPr'):
            dst_tree.remove(sp)
    # The emptied destination page keeps only its Jeju layout relationship.
    # Stale image/chart links from an earlier reuse must not retain orphan parts.
    for rel in list(dst_rels):
        if rel.get('Type') != R + '/slideLayout':
            dst_rels.remove(rel)
    table = placeholder_table(src, src_part)
    style_context = _theme_context(src, src_part)
    requested = {str(v) for v in (include_shapes or [])}
    if not requested and not all_body_shapes:
        raise ValueError('reuse requires --include-shape ID/name (repeatable) or --all-body-shapes; available: ' + json.dumps(source_shape_catalog(src, src_part), ensure_ascii=False))
    imported = []
    for sp in src_tree:
        if sp.tag.split('}')[-1] in ('nvGrpSpPr', 'grpSpPr'):
            continue
        sid,name=direct_shape_identity(sp)
        source_box = _raw_box(sp)
        placeholder = ph_key(sp)
        furniture = (placeholder and placeholder[0] in ('title', 'ctrTitle', 'dt', 'ftr', 'sldNum', 'hdr'))
        if source_box and source_box[1] < round(1.45 * jt.EMU) and source_box[3] < round(1.05 * jt.EMU):
            furniture = True
        if all_body_shapes and furniture:
            continue
        if not all_body_shapes and sid not in requested and name not in requested:
            continue
        clone = copy.deepcopy(sp)
        materialize_source_style(clone, src, src_part, style_context)
        strip_placeholders(clone, table, warnings)
        _normalize_table_labels(clone)
        imported.append((clone,{'source_id':sid,'source_name':name,
                                'source_shape_sha256':shape_hash(clone),
                                'kind':sp.tag.split('}')[-1]}))
        requested.discard(sid);requested.discard(name)
    if requested:
        raise ValueError('Unknown --include-shape ID/name: '+', '.join(sorted(requested))+'; available: '+json.dumps(source_shape_catalog(src,src_part),ensure_ascii=False))
    if not imported:
        raise ValueError('Source slide has no selected body shapes to reuse')
    imp = Importer(src, dst)
    next_rid = max([int(m) for r in dst_rels for m in re.findall(r'\d+$', r.get('Id', ''))] + [0]) + 1
    rid_map = {}
    for shape,_ in imported:
        for el in shape.iter():
            for attr in REL_ATTRS:
                key = '{'+R+'}' + attr; old = el.get(key)
                if not old:
                    continue
                if old not in rid_map:
                    rel = src_rels.get(old)
                    if rel is None:
                        raise ValueError('Selected source shape has a missing relationship: '+old)
                    new_id = f'rIdReuse{next_rid}'; next_rid += 1
                    if rel.get('TargetMode') == 'External':
                        jt.ET.SubElement(dst_rels, '{'+PKG+'}Relationship', Id=new_id, Type=rel.get('Type'), Target=rel.get('Target'), TargetMode='External')
                    else:
                        new_part = imp.copy_part(jt.target(src_part, rel))
                        jt.ET.SubElement(dst_rels, '{'+PKG+'}Relationship', Id=new_id, Type=rel.get('Type'), Target=posixpath.relpath(new_part, posixpath.dirname(dst_part)))
                    rid_map[old] = new_id
                if old in rid_map:
                    el.set(key, rid_map[old])
    imp.finish()
    maxid = max([int(c.get('id', '0')) for c in dst_xml.findall('.//p:cNvPr', NS)] + [1])
    id_map = {}
    for shape,_ in imported:
        for c in shape.findall('.//p:cNvPr', NS):
            old = c.get('id')
            maxid += 1
            if old:
                id_map[old] = str(maxid)
            c.set('id', str(maxid))
    # Connector endpoints reference cNvPr IDs.  Remap them with the copied
    # shapes or PowerPoint rejects the package even though the XML is valid.
    for shape,_ in imported:
        for node in shape.iter():
            if node.tag.split('}')[-1] in ('stCxn', 'endCxn') and node.get('id') in id_map:
                node.set('id', id_map[node.get('id')])
    fit = _fit_shapes(imported, box)
    if fit['scale'] < 0.7:
        warnings.append('fit scale is below 0.70; choose another layout or split the slide')
    if len(imported) > 30:
        warnings.append(f'density risk: {len(imported)} top-level shapes were copied; select fewer shapes or split the slide')
    selected=[]
    for shape,identity in imported:
        dst_tree.append(shape)
        candidate_id,candidate_name=direct_shape_identity(shape)
        selected.append(dict(identity,candidate_id=candidate_id,candidate_name=candidate_name,
                             kind=shape.tag.split('}')[-1]))
    dst[dst_part] = jt.serialize(dst_xml)
    dst[jt.relpart(dst_part)] = jt.serialize(dst_rels)
    # Report only source/candidate identities actually selected.  Full XML stays in
    # the automatic binding below rather than becoming user-authored plan data.
    return dst_part, imp, warnings, len(imported), fit, selected


def _text_luminances(slide):
    values=[]
    for shape in slide.findall('.//p:sp',NS):
        if not ''.join(node.text or '' for node in shape.findall('.//a:t',NS)).strip():continue
        # Filled labels and cards carry their own contrast surface.
        if shape.find('p:spPr/a:solidFill',NS) is not None:continue
        colors=[]
        for node in shape.findall('.//a:rPr/a:solidFill/a:srgbClr',NS)+shape.findall('.//a:defRPr/a:solidFill/a:srgbClr',NS):
            value=node.get('val','')
            if re.fullmatch(r'[0-9A-Fa-f]{6}',value):colors.append(value)
        for value in colors:
            r,g,b=(int(value[i:i+2],16)/255 for i in (0,2,4))
            values.append(0.2126*r+0.7152*g+0.0722*b)
    return values


def contrast_warnings(files,slide_part,layout):
    values=_text_luminances(jt.xml(files[slide_part]))
    if not values:return []
    values.sort();median=values[len(values)//2]
    dark=layout in {6,7,10,12,13}
    if dark and median < 0.46:
        return [f'contrast risk: dark layout {layout} contains predominantly dark unfilled text; use a light layout or recolor after visual review']
    return []


def _layout_id(files, slide_part):
    rels = jt.xml(files[jt.relpart(slide_part)])
    layout = next((jt.target(slide_part, r) for r in rels if r.get('Type') == R + '/slideLayout'), '')
    match = re.search(r'slideLayout(\d+)\.xml$', layout)
    if not match:
        raise ValueError('Destination slide has no CDSA layout')
    return int(match.group(1))


def _band_texts(files, slide_part):
    layout, master = layout_and_master(files, slide_part)
    values = []
    for part in (slide_part, layout, master):
        if not part or part not in files:
            continue
        tree = jt.xml(files[part]).find('p:cSld/p:spTree', NS)
        if tree is None:
            continue
        for shape in tree:
            box = _raw_box(shape)
            if box and box[1] > round(1.5 * jt.EMU):
                continue
            value = ' '.join(''.join(t.text or '' for t in p.findall('.//a:t', NS)).strip()
                             for p in shape.findall('.//a:p', NS)).strip()
            if len(value) >= 3 and re.search(r'[A-Za-z가-힣]', value) and value not in values:
                values.append(value)
    return values


def reuse(run, page, source, include_shapes=None, all_body_shapes=False, box=None):
    run = run.resolve()
    cand_path = run / 'candidate.pptx'
    if not cand_path.exists():
        raise ValueError('candidate.pptx not found in ' + str(run))
    src_label = source;resolved_sid=None
    requested=source.removeprefix('ssj-slide://')
    if re.fullmatch(r'[sg]-[0-9a-f]{20}',requested):
        src_path, resolved_sid = export_from_library(requested, run)
        # Group IDs are aliases. Record the concrete representative slide so source
        # validation and later re-runs refer to the exact object that was imported.
        src_label = 'ssj-slide://' + resolved_sid
    else:
        src_path = pathlib.Path(source)
        if not src_path.exists():raise ValueError('Source must be a library slide ID or an existing PPTX')
    src_path=src_path.resolve()
    source_sha256=file_sha(src_path)
    try:source_file=src_path.relative_to(run).as_posix()
    except ValueError:source_file=str(src_path)
    source_fingerprint=None
    if resolved_sid:
        import slide_store,visual_retrieval
        with sqlite3.connect('file:'+slide_store.DB.resolve().as_posix()+'?mode=ro',uri=True) as db:
            source_fingerprint=visual_retrieval.source_fingerprint(db,resolved_sid)
    src = jt.read(src_path); dst = jt.read(cand_path)
    src_size, dst_size = slide_size(src), slide_size(dst)
    if src_size != dst_size:
        raise ValueError(
            'Source and candidate slide sizes differ: '
            f'{src_size[0]}x{src_size[1]} vs {dst_size[0]}x{dst_size[1]} EMU'
        )
    src_part = ordered_slides(src)[0]
    dst_parts = ordered_slides(dst)
    if not 1 <= page <= len(dst_parts):
        raise ValueError(f'page must be 1..{len(dst_parts)}')
    dst_part = dst_parts[page-1]

    template = jt.read(run / 'template.pptx') if (run / 'template.pptx').exists() else jt.read(jt.DEFAULT)
    layout = _layout_id(dst, dst_part)
    target_box = tuple(box) if box is not None else jt.content_box(layout)
    if len(target_box) != 4 or min(target_box[2:]) <= 0:
        raise ValueError('box must contain positive x,y,w,h values')
    band_texts = _band_texts(src, src_part)
    slide_part, imp, warnings, shapes, fit, selected_shapes = mode_reuse(
        src, dst, src_part, dst_part, target_box, include_shapes, all_body_shapes)
    warnings.extend(contrast_warnings(dst,slide_part,layout))

    removed_parts=prune_unreachable_imports(dst,template)
    errors = jt.verify(dst, template)
    if errors:
        raise ValueError('Reuse would break package integrity:\n' + '\n'.join(errors))
    with zipfile.ZipFile(cand_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for n, data in dst.items():
            z.writestr(n, data)

    # record the images this slide itself references
    assets_path = run / 'assets.json'
    assets = load(assets_path) if assets_path.exists() else {'images': []}
    known = {a.get('sha256') for a in assets['images']}
    recorded = 0
    for r in jt.xml(dst[jt.relpart(slide_part)]):
        if r.get('Type') != R + '/image' or r.get('TargetMode') == 'External':
            continue
        part = jt.target(slide_part, r); data = dst[part]; h = sha(data)
        if h in known:
            continue
        (run / 'reused').mkdir(exist_ok=True)
        f = run / 'reused' / (h[:16] + posixpath.splitext(part)[1])
        f.write_bytes(data)
        assets['images'].append({'file': f.relative_to(run).as_posix(), 'sha256': h, 'role': 'reused', 'source': f'{src_label} page {page}'})
        known.add(h); recorded += 1
    save(assets_path, assets)

    plan_path = run / 'plan.json'
    if plan_path.exists():
        plan = load(plan_path)
        if len(plan.get('slides', [])) >= page:
            item = plan['slides'][page-1]
            item['reuse'] = {'source': src_label, 'mode': 'reuse',
                             'shapes': [str(row['candidate_id']) for row in selected_shapes],
                             'source_shapes': [str(row['source_id']) for row in selected_shapes],
                             'fit': fit, 'band_texts': band_texts,
                             'source_file': source_file, 'source_sha256': source_sha256,
                             'warnings': warnings}
            if src_label.startswith('ssj-slide://'):
                sid = src_label.removeprefix('ssj-slide://')
                srcs = item.setdefault('sources', [])
                if not any(s.get('kind')=='internal' and s.get('slide_id') == sid for s in srcs):
                    srcs.append({'kind': 'internal', 'slide_id': sid})
            save(plan_path, plan)
    return {'page': page, 'source': src_label, 'mode': 'reuse', 'slide_part': slide_part, 'shapes': shapes,
            'parts_copied': len(imp.copied), 'parts_pruned': len(removed_parts),
            'images_recorded': recorded, 'fit': fit, 'selected_shapes': selected_shapes,
            'warnings': warnings,
            'next': 'harness check --structural-only, render, then review this page'}


def main():
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--run', type=pathlib.Path, required=True)
    ap.add_argument('--page', type=int, required=True)
    ap.add_argument('--source', required=True, help='library slide id (s-…/g-…) or a .pptx path')
    ap.add_argument('--include-shape',action='append',default=[],metavar='ID_OR_NAME',help='copy only this top-level source shape; repeat as needed')
    ap.add_argument('--all-body-shapes',action='store_true',help='copy all non-header top-level source shapes')
    ap.add_argument('--box', help='optional x,y,w,h in inches; defaults to the selected layout safe area')
    a = ap.parse_args()
    custom_box = None
    if a.box:
        values = [float(value.strip()) for value in a.box.split(',')]
        if len(values) != 4:
            raise ValueError('--box requires x,y,w,h in inches')
        custom_box = tuple(round(value * jt.EMU) for value in values)
    print(json.dumps(reuse(a.run,a.page,a.source,a.include_shape,a.all_body_shapes,custom_box),ensure_ascii=False,indent=2))

if __name__ == '__main__':
    try:
        main()
    except Exception as exc:
        print(str(exc), file=sys.stderr); sys.exit(1)
