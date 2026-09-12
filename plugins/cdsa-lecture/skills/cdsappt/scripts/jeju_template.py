"""Create a deck from the preserved Jeju master and verify its structural fidelity."""
import argparse
import json
import pathlib
import posixpath
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
import ooxml

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
DEFAULT = pathlib.Path(__file__).resolve().parents[1] / 'assets' / 'jeju-master.pptx'
EMU = 914400

# Safe body areas measured from the 13 rendered layouts.  These are explicit
# because geometry-only heuristics mistake large background panels for headers.
# Values are inches: x, y, width, height.
CONTENT_BOXES_IN = {
    1: (0.45, 0.65, 12.43, 6.25),
    2: (0.45, 1.48, 12.43, 5.42),
    3: (0.45, 1.48, 12.43, 5.42),
    4: (0.45, 1.52, 12.43, 5.38),
    5: (0.45, 1.52, 12.43, 5.38),
    6: (0.55, 0.72, 12.23, 6.18),
    7: (0.55, 0.72, 12.23, 6.18),
    8: (0.55, 1.62, 12.23, 5.28),
    9: (0.78, 1.38, 11.77, 5.48),
    10: (0.55, 0.80, 12.23, 5.98),
    11: (0.70, 0.78, 11.93, 5.95),
    12: (0.45, 0.45, 12.43, 6.60),
    13: (0.70, 0.75, 11.93, 5.95),
}

def content_box(layout):
    """Return the validated body area for a Jeju layout as EMU."""
    if layout not in CONTENT_BOXES_IN:
        raise ValueError('layout must be 1..13')
    return tuple(round(v * EMU) for v in CONTENT_BOXES_IN[layout])

def read(path):
    with zipfile.ZipFile(path) as z:
        return {n: z.read(n) for n in z.namelist() if not n.endswith('/')}

def xml(data):
    return ooxml.parse(data)

def serialize(e):
    return ooxml.serialize(e)

def relpart(part):
    return posixpath.join(posixpath.dirname(part), '_rels', posixpath.basename(part) + '.rels')

def target(part, rel):
    return posixpath.normpath(posixpath.join(posixpath.dirname(part), rel.get('Target'))).lstrip('/')

def layout_slides(files):
    pres = xml(files['ppt/presentation.xml'])
    rr = {r.get('Id'):r for r in xml(files['ppt/_rels/presentation.xml.rels'])}
    result = {}
    for sid in pres.find('{' + P + '}sldIdLst'):
        part = target('ppt/presentation.xml', rr[sid.get('{' + R + '}id')])
        relations = xml(files[relpart(part)])
        lr = next(r for r in relations if r.get('Type') == R + '/slideLayout')
        lp = target(part, lr)
        index = int(re.search(r'slideLayout(\d+)\.xml$', lp)[1])
        result[index] = (part, lp)
    return result

def protected(files):
    # Follow only registered presentation masters.  A chart may legitimately
    # bring a private theme part; it is not a second slide master.
    pres = xml(files['ppt/presentation.xml'])
    prels = {r.get('Id'): r for r in xml(files['ppt/_rels/presentation.xml.rels'])}
    todo = []
    for node in pres.findall('.//{' + P + '}sldMasterId'):
        rel = prels.get(node.get('{' + R + '}id'))
        if rel is not None and rel.get('Type') == R + '/slideMaster':
            todo.append(target('ppt/presentation.xml', rel))
    seen = set()
    while todo:
        part = todo.pop()
        if part in seen:
            continue
        if part not in files:
            raise ValueError('Missing template dependency: ' + part)
        seen.add(part)
        rp = relpart(part)
        if rp in files:
            seen.add(rp)
            for r in xml(files[rp]):
                if r.get('TargetMode') != 'External':
                    todo.append(target(part, r))
    return seen

def signature(data, name):
    if name.endswith(('.xml', '.rels')):
        def tree(e):
            return (e.tag, tuple(sorted(e.attrib.items())), (e.text or '').strip(), tuple(tree(c) for c in e))
        return tree(xml(data))
    return data

def verify(files, baseline):
    """Require the intact single CDSA master and its 13 layouts."""
    errors = []
    expected = protected(baseline)
    for n in sorted(expected):
        if n not in files:
            errors.append('Missing protected part: ' + n)
        elif signature(files[n], n) != signature(baseline[n], n):
            errors.append('Changed protected part: ' + n)
    try:
        extra = protected(files) - expected
    except ValueError as exc:
        errors.append(str(exc))
        extra = set()
    errors.extend('Unexpected master/layout/theme dependency: ' + n for n in sorted(extra))
    srcsize = xml(baseline['ppt/presentation.xml']).find('{' + P + '}sldSz').attrib
    if xml(files['ppt/presentation.xml']).find('{' + P + '}sldSz').attrib != srcsize:
        errors.append('Slide size differs from template')
    for n, data in files.items():
        if n.endswith('.rels'):
            part = '' if n == '_rels/.rels' else posixpath.join(posixpath.dirname(posixpath.dirname(n)), posixpath.basename(n)[:-5])
            for r in xml(data):
                if r.get('TargetMode') != 'External' and target(part,r) not in files:
                    errors.append('Broken relationship: ' + n + ' -> ' + r.get('Target'))
    for ov in xml(files['[Content_Types].xml']):
        if ov.tag == '{' + CT + '}Override' and ov.get('PartName').lstrip('/') not in files:
            errors.append('Missing content type part: ' + ov.get('PartName'))
    baseline_layouts = {lp for _, lp in layout_slides(baseline).values()}
    for n in files:
        if re.fullmatch(r'ppt/slides/slide[^/]*\.xml', n):
            e = xml(files[n])
            rr = xml(files[relpart(n)])
            lp = next((target(n,r) for r in rr if r.get('Type') == R + '/slideLayout'), None)
            if lp not in baseline_layouts:
                errors.append('Slide uses a layout outside the template: ' + n)
                continue
            if e.get('showMasterSp') in ('0','false'):
                errors.append('Master graphics hidden: ' + n)
            if e.find('{' + P + '}cSld/{' + P + '}bg') is not None:
                # Template 13 may contain a native background. Only a new override is invalid.
                originals = layout_slides(baseline)
                idx = int(re.search(r'slideLayout(\d+)\.xml$',lp)[1])
                original = xml(baseline[originals[idx][0]]).find('{' + P + '}cSld/{' + P + '}bg')
                if original is None or signature(serialize(original),'x.xml') != signature(serialize(e.find('{' + P + '}cSld/{' + P + '}bg')),'x.xml'):
                    errors.append('Slide background overrides template: ' + n)
    return errors

def imported_slides(files, baseline):
    """Slide parts whose layout is not one of the baseline (Jeju) layouts."""
    baseline_layouts = {lp for _, lp in layout_slides(baseline).values()}
    out = set()
    for n in files:
        if re.fullmatch(r'ppt/slides/slide[^/]*\.xml', n) and relpart(n) in files:
            rr = xml(files[relpart(n)])
            lp = next((target(n,r) for r in rr if r.get('Type') == R + '/slideLayout'), None)
            if lp not in baseline_layouts:
                out.add(n)
    return out

def create(baseline, layouts, output):
    if output.exists():
        raise ValueError('Output already exists; choose a new path: ' + str(output))
    originals = layout_slides(baseline)
    if not layouts or any(i not in originals for i in layouts):
        raise ValueError('Select one or more layout IDs from 1 through 13')
    files = dict(baseline)
    pres = xml(files['ppt/presentation.xml'])
    ids = pres.find('{' + P + '}sldIdLst')
    ids.clear()
    # Remove slide-specific extension records containing stale slide IDs.
    ext = pres.find('{' + P + '}extLst')
    if ext is not None:
        pres.remove(ext)
    rr = xml(files['ppt/_rels/presentation.xml.rels'])
    for r in list(rr):
        if r.get('Type') == R + '/slide':
            rr.remove(r)
    removed = [n for n in files if n.startswith('ppt/slides/')]
    for n in removed:
        del files[n]
    ct = xml(files['[Content_Types].xml'])
    for el in list(ct):
        if el.get('PartName','').lstrip('/') in removed:
            ct.remove(el)
    for page, idx in enumerate(layouts,1):
        src, lp = originals[idx]
        part = f'ppt/slides/slide{page}.xml'
        files[part] = baseline[src]
        sr = ET.Element('{' + PKG + '}Relationships')
        ET.SubElement(sr,'{' + PKG + '}Relationship',Id='rId1',Type=R+'/slideLayout',Target=posixpath.relpath(lp,'ppt/slides'))
        # Newly generated template slides must have only their layout dependency.
        source_rels = list(xml(baseline[relpart(src)]))
        if len(source_rels) != 1 or source_rels[0].get('Type') != R + '/slideLayout':
            raise ValueError('Template sample slide has unexpected dependencies: ' + src)
        files[relpart(part)] = serialize(sr)
        rid = f'rIdJeju{page}'
        ET.SubElement(rr,'{' + PKG + '}Relationship',Id=rid,Type=R+'/slide',Target=f'slides/slide{page}.xml')
        ET.SubElement(ids,'{' + P + '}sldId',{'id':str(255+page),'{'+R+'}id':rid})
        ET.SubElement(ct,'{' + CT + '}Override',PartName='/'+part,ContentType='application/vnd.openxmlformats-officedocument.presentationml.slide+xml')
    files['ppt/presentation.xml'] = serialize(pres)
    files['ppt/_rels/presentation.xml.rels'] = serialize(rr)
    files['[Content_Types].xml'] = serialize(ct)
    if 'docProps/app.xml' in files:
        app = xml(files['docProps/app.xml'])
        for el in app.iter():
            if el.tag.endswith('}Slides'):
                el.text = str(len(layouts))
        files['docProps/app.xml'] = serialize(app)
    # Remove obsolete preview rather than advertising the source cover.
    if '_rels/.rels' in files:
        rootrels=xml(files['_rels/.rels'])
        for r in list(rootrels):
            if r.get('Type','').endswith('/thumbnail'):
                rootrels.remove(r)
        files['_rels/.rels']=serialize(rootrels)
    issues=verify(files,baseline)
    if issues:
        raise ValueError('\n'.join(issues))
    output.parent.mkdir(parents=True,exist_ok=True)
    with zipfile.ZipFile(output,'w',zipfile.ZIP_DEFLATED) as z:
        for n,data in files.items():
            z.writestr(n,data)
    print(json.dumps({'output':str(output),'layouts':layouts,'master_integrity':'PASS','content':'empty template slides; content still required'},ensure_ascii=False))

def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--template',type=pathlib.Path,default=DEFAULT)
    sub=parser.add_subparsers(dest='command',required=True)
    sub.add_parser('list')
    init=sub.add_parser('init');init.add_argument('--layouts',required=True);init.add_argument('--output',required=True,type=pathlib.Path)
    check=sub.add_parser('verify');check.add_argument('deck',type=pathlib.Path)
    args=parser.parse_args();baseline=read(args.template)
    if args.command=='list':
        for i,(s,lp) in sorted(layout_slides(baseline).items()):
            print(i,xml(baseline[lp]).find('{'+P+'}cSld').get('name'))
    elif args.command=='init':
        create(baseline,[int(x.strip()) for x in args.layouts.split(',')],args.output)
    else:
        errors=verify(read(args.deck),baseline)
        print(json.dumps({'master_integrity':'FAIL' if errors else 'PASS','errors':errors,
                          'single_master_required':True,'visual_review_required':True},ensure_ascii=False,indent=2))
        return bool(errors)
    return 0

if __name__=='__main__':
    sys.exit(main())
