"""Insert a full-slide cover image on page 1 and register its provenance."""
import argparse
import hashlib
import json
import pathlib
import posixpath
import shutil
import sys
import zipfile
import xml.etree.ElementTree as E

from PIL import Image
import jeju_template as jt
import reuse_slide

P, A, R, PKG, CT = jt.P, reuse_slide.A, jt.R, jt.PKG, jt.CT
NS = {'p': P, 'a': A}


def sha(path):
    with pathlib.Path(path).open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()


def apply(run, image_path, source):
    run = pathlib.Path(run).resolve()
    deck_path = run / 'candidate.pptx'
    if not deck_path.is_file():
        raise ValueError('candidate.pptx is missing')
    image_path = pathlib.Path(image_path).resolve()
    with Image.open(image_path) as image:
        width, height = image.size
        fmt = image.format
    suffix = image_path.suffix.lower()
    if fmt not in ('PNG', 'JPEG') or suffix not in ('.png', '.jpg', '.jpeg'):
        raise ValueError('Cover image must be PNG or JPEG')
    digest = sha(image_path)
    asset_dir = run / 'cover-assets'; asset_dir.mkdir(exist_ok=True)
    local_image = asset_dir / (digest[:16] + suffix)
    if not local_image.exists():
        shutil.copy2(image_path, local_image)

    files = jt.read(deck_path)
    slide_part = reuse_slide.ordered_slides(files)[0]
    slide = jt.xml(files[slide_part]); tree = slide.find('p:cSld/p:spTree', NS)
    rels = jt.xml(files[jt.relpart(slide_part)])
    for shape in list(tree):
        nv = shape.find('.//p:cNvPr', NS)
        if nv is not None and nv.get('name') == 'CDSA cover background':
            tree.remove(shape)
    for rel in list(rels):
        if rel.get('Id') == 'rIdCdsapptCover':
            rels.remove(rel)

    media_part = 'ppt/media/cdsappt-cover-' + digest[:16] + suffix
    files[media_part] = local_image.read_bytes()
    E.SubElement(rels, '{'+PKG+'}Relationship', Id='rIdCdsapptCover', Type=R+'/image',
                 Target=posixpath.relpath(media_part, posixpath.dirname(slide_part)))
    canvas_w, canvas_h = reuse_slide.slide_size(files)
    pic = E.Element('{'+P+'}pic')
    nv = E.SubElement(pic, '{'+P+'}nvPicPr')
    max_id = max([int(node.get('id', '0')) for node in slide.findall('.//p:cNvPr', NS)] + [1])
    E.SubElement(nv, '{'+P+'}cNvPr', id=str(max_id+1), name='CDSA cover background')
    E.SubElement(nv, '{'+P+'}cNvPicPr'); E.SubElement(nv, '{'+P+'}nvPr')
    fill = E.SubElement(pic, '{'+P+'}blipFill')
    E.SubElement(fill, '{'+A+'}blip', {'{'+R+'}embed':'rIdCdsapptCover'})
    image_ratio = width / height; canvas_ratio = canvas_w / canvas_h
    crop = {'l':'0','r':'0','t':'0','b':'0'}
    if image_ratio > canvas_ratio:
        kept = canvas_ratio / image_ratio; cut = round((1-kept)*50000)
        crop['l'] = crop['r'] = str(cut)
    elif image_ratio < canvas_ratio:
        kept = image_ratio / canvas_ratio; cut = round((1-kept)*50000)
        crop['t'] = crop['b'] = str(cut)
    E.SubElement(fill, '{'+A+'}srcRect', crop)
    stretch = E.SubElement(fill, '{'+A+'}stretch'); E.SubElement(stretch, '{'+A+'}fillRect')
    props = E.SubElement(pic, '{'+P+'}spPr'); xfrm = E.SubElement(props, '{'+A+'}xfrm')
    E.SubElement(xfrm, '{'+A+'}off', x='0', y='0')
    E.SubElement(xfrm, '{'+A+'}ext', cx=str(canvas_w), cy=str(canvas_h))
    E.SubElement(props, '{'+A+'}prstGeom', prst='rect').append(E.Element('{'+A+'}avLst'))
    tree.insert(2, pic)
    files[slide_part] = jt.serialize(slide); files[jt.relpart(slide_part)] = jt.serialize(rels)
    types = jt.xml(files['[Content_Types].xml'])
    extension = suffix.lstrip('.')
    if not any(node.tag == '{'+CT+'}Default' and node.get('Extension','').lower() == extension for node in types):
        default = E.Element('{'+CT+'}Default', Extension=suffix.lstrip('.'),
                            ContentType='image/png' if suffix == '.png' else 'image/jpeg')
        first_override = next((index for index, node in enumerate(types)
                               if node.tag == '{'+CT+'}Override'), len(types))
        types.insert(first_override, default)
    files['[Content_Types].xml'] = jt.serialize(types)
    with zipfile.ZipFile(deck_path, 'w', zipfile.ZIP_DEFLATED) as archive:
        for name, data in files.items(): archive.writestr(name, data)

    assets_path = run / 'assets.json'
    assets = json.loads(assets_path.read_text(encoding='utf-8-sig')) if assets_path.exists() else {'images':[]}
    assets['images'] = [row for row in assets.get('images', []) if row.get('role') != 'cover']
    assets['images'].append({'file':local_image.relative_to(run).as_posix(),'sha256':digest,
                             'role':'cover','source':source})
    assets_path.write_text(json.dumps(assets,ensure_ascii=False,indent=2),encoding='utf-8')
    return {'page':1,'file':str(local_image),'sha256':digest,'crop':crop}


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--run',type=pathlib.Path,required=True)
    parser.add_argument('--image',type=pathlib.Path,required=True)
    parser.add_argument('--source',default='Codex image generation')
    args=parser.parse_args()
    print(json.dumps(apply(args.run,args.image,args.source),ensure_ascii=False,indent=2))


if __name__=='__main__':
    try:
        if hasattr(sys.stdout,'reconfigure'):sys.stdout.reconfigure(encoding='utf-8')
        main()
    except Exception as exc:
        print(str(exc),file=sys.stderr);sys.exit(1)
