"""Embedded slide objects: no original PPTX or user-machine paths needed at runtime."""
import argparse,hashlib,json,os,pathlib,posixpath,sqlite3,zipfile,zlib
import ooxml
BASE=pathlib.Path(__file__).resolve().parents[1]
DB=pathlib.Path(os.environ.get('CDSA_LECTURE_DB') or BASE/'data/lecture-library.sqlite')
P='http://schemas.openxmlformats.org/presentationml/2006/main'
R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG='http://schemas.openxmlformats.org/package/2006/relationships'

def relpart(p):return posixpath.join(posixpath.dirname(p),'_rels',posixpath.basename(p)+'.rels')
def target(p,r):return posixpath.normpath(posixpath.join(posixpath.dirname(p),r.get('Target'))).lstrip('/')
def xml(e):return ooxml.serialize(e)
E=type('E',(),{'fromstring':staticmethod(ooxml.parse)})  # keep call sites unchanged

def ingest(db):
    db.executescript('CREATE TABLE IF NOT EXISTS blobs(sha TEXT PRIMARY KEY,data BLOB); CREATE TABLE IF NOT EXISTS packages(file_id INTEGER PRIMARY KEY,parts TEXT);')
    for fid,path in db.execute('SELECT id,path FROM files ORDER BY id').fetchall():
        if db.execute('SELECT 1 FROM packages WHERE file_id=?',(fid,)).fetchone():continue
        src=pathlib.Path(path)
        if not src.is_absolute():src=(BASE/src).resolve()
        parts={}
        with zipfile.ZipFile(src) as z:
            for name in z.namelist():
                if name.endswith('/'):continue
                data=z.read(name);sha=hashlib.sha256(data).hexdigest();parts[name]=sha
                if not db.execute('SELECT 1 FROM blobs WHERE sha=?',(sha,)).fetchone():
                    db.execute('INSERT INTO blobs VALUES(?,?)',(sha,zlib.compress(data,1)))
        db.execute('INSERT INTO packages VALUES(?,?)',(fid,json.dumps(parts)));db.commit();print('Embedded source',fid,flush=True)
    # Human provenance remains a name and page number, never an opening path.
    for fid,path in db.execute('SELECT id,path FROM files').fetchall():
        title=pathlib.PurePosixPath(path.replace('\\','/')).name
        db.execute('UPDATE files SET path=? WHERE id=?',('ssj-source://'+str(fid)+'/'+title,fid))
    info=json.loads(db.execute("SELECT value FROM meta WHERE key='summary'").fetchone()[0]);info.pop('source_root',None)
    info['storage']='embedded slide objects and deduplicated binary assets';info['original_files_required']=False
    db.execute("UPDATE meta SET value=? WHERE key='summary'",(json.dumps(info,ensure_ascii=False),));db.commit()
    DB.with_suffix('.summary.json').write_text(json.dumps(info,ensure_ascii=False,indent=2),encoding='utf-8')
    print('Embedded store complete',flush=True)

def export(db,sid,out):
    sid=sid.removeprefix('ssj-slide://')
    if sid.startswith('g-'):
        row=db.execute('SELECT id FROM slides WHERE group_id=? ORDER BY id LIMIT 1',(sid,)).fetchone()
        if not row:raise ValueError('Unknown group')
        sid=row[0]
    row=db.execute('SELECT file_id,detail FROM slides WHERE id=?',(sid,)).fetchone()
    if not row:raise ValueError('Unknown slide ID')
    fid,detail=row;part=json.loads(detail)['part']
    pm=db.execute('SELECT parts FROM packages WHERE file_id=?',(fid,)).fetchone()
    if not pm:raise ValueError('Embedded package missing')
    manifest=json.loads(pm[0]);overrides={}
    def get(n):
        if n in overrides:return overrides[n]
        blob=db.execute('SELECT data FROM blobs WHERE sha=?',(manifest[n],)).fetchone()
        if not blob:raise ValueError('Missing asset '+manifest[n])
        value=zlib.decompress(blob[0])
        if hashlib.sha256(value).hexdigest()!=manifest[n]:raise ValueError('Corrupt asset '+manifest[n])
        return value
    pr=E.fromstring(get('ppt/_rels/presentation.xml.rels'));keepid=None
    for r in list(pr):
        if r.get('Type')==R+'/slide':
            if target('ppt/presentation.xml',r)==part:keepid=r.get('Id')
            else:pr.remove(r)
    if keepid is None:raise ValueError('Source slide relationship missing')
    pres=E.fromstring(get('ppt/presentation.xml'));ids=pres.find('{'+P+'}sldIdLst')
    for s in list(ids):
        if s.get('{'+R+'}id')!=keepid:ids.remove(s)
    for tag in ('extLst','custShowLst'):
        node=pres.find('{'+P+'}'+tag)
        if node is not None:pres.remove(node)
    overrides['ppt/presentation.xml']=xml(pres);overrides['ppt/_rels/presentation.xml.rels']=xml(pr)
    rootrels=E.fromstring(get('_rels/.rels'))
    for r in list(rootrels):
        if r.get('Type','').endswith('/thumbnail'):rootrels.remove(r)
    overrides['_rels/.rels']=xml(rootrels)
    todo=[target('',r) for r in rootrels if r.get('TargetMode')!='External'];keep={'_rels/.rels'}
    while todo:
        n=todo.pop()
        if n in keep:continue
        if n not in manifest:raise ValueError('Broken dependency '+n)
        keep.add(n);rp=relpart(n)
        if rp in manifest:
            keep.add(rp)
            todo.extend(target(n,r) for r in E.fromstring(get(rp)) if r.get('TargetMode')!='External')
    ct=E.fromstring(get('[Content_Types].xml'))
    for el in list(ct):
        if el.tag.endswith('}Override') and el.get('PartName','').lstrip('/') not in keep:ct.remove(el)
    overrides['[Content_Types].xml']=xml(ct);keep.add('[Content_Types].xml')
    if 'docProps/app.xml' in keep:
        app=E.fromstring(get('docProps/app.xml'))
        for el in app.iter():
            if el.tag.endswith('}Slides'):el.text='1'
        overrides['docProps/app.xml']=xml(app)
    if out.exists():raise ValueError('Output exists')
    out.parent.mkdir(parents=True,exist_ok=True)
    with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as z:
        for n in sorted(keep):z.writestr(n,get(n))
    return {'slide_id':sid,'output':str(out),'source_files_used':False,'parts':len(keep)}

def main():
    global DB
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('--db',type=pathlib.Path,default=DB)
    sub=p.add_subparsers(dest='command',required=True);sub.add_parser('ingest')
    ex=sub.add_parser('export');ex.add_argument('id');ex.add_argument('--output',type=pathlib.Path,required=True)
    a=p.parse_args();DB=a.db;d=sqlite3.connect(a.db if a.command=='ingest' else 'file:'+a.db.as_posix()+'?mode=ro',uri=a.command!='ingest')
    if a.command=='ingest':ingest(d)
    else:print(json.dumps(export(d,a.id,a.output),ensure_ascii=False))
if __name__=='__main__':main()
