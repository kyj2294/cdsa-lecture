"""Execute corpus visual retrieval for a harness run; metadata is automatic and internal."""
import hashlib
import json
import pathlib
import posixpath
import sqlite3
import uuid
import zipfile
import zlib
import xml.etree.ElementTree as ET
import visual_library as library
PIPELINE_VERSION = 5
EXPECTED_DECKS = 78
EXPECTED_SLIDES = 7961
REL_NS='http://schemas.openxmlformats.org/package/2006/relationships'
OFFICE_REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
SKIP_SOURCE_RELS={OFFICE_REL+'/notesSlide',OFFICE_REL+'/comments',OFFICE_REL+'/slide'}


def _relpart(part):
    return posixpath.join(posixpath.dirname(part),'_rels',posixpath.basename(part)+'.rels')


def _target(part,rel):
    return posixpath.normpath(posixpath.join(posixpath.dirname(part),rel.get('Target'))).lstrip('/')


def source_fingerprint(db, sid):
    """Cheap DB identity for one embedded slide without exporting a PPTX.

    The whole source-package manifest is included so a slide ID cannot silently
    move to a different package.  The selected slide's visual hash and original
    part/page identify the exact page.  Harness checks this value on every run;
    the expensive native export happens only when reuse is first applied.
    """
    row=db.execute('SELECT file_id,page,visual_hash,detail FROM slides WHERE id=?',(sid,)).fetchone()
    if not row:raise ValueError('Unknown slide ID: '+str(sid))
    file_id,page,visual_hash,detail_text=row
    package=db.execute('SELECT parts FROM packages WHERE file_id=?',(file_id,)).fetchone()
    if not package or not package[0]:raise ValueError('Embedded package missing for '+sid)
    detail=json.loads(detail_text);part=detail.get('part')
    if not isinstance(part,str) or not part:raise ValueError('Embedded slide part missing for '+sid)
    manifest=json.loads(package[0])
    if part not in manifest:raise ValueError('Embedded slide part is absent from package for '+sid)
    package_digest=hashlib.sha256(package[0].encode('utf-8')).hexdigest()
    graph={};todo=[part]
    while todo:
        name=todo.pop()
        if name in graph:continue
        digest=manifest.get(name)
        if not digest:raise ValueError('Embedded graph part is missing for '+sid+': '+name)
        graph[name]=digest;relationship_part=_relpart(name)
        if relationship_part not in manifest:continue
        relationship_digest=manifest[relationship_part];graph[relationship_part]=relationship_digest
        blob=db.execute('SELECT data FROM blobs WHERE sha=?',(relationship_digest,)).fetchone()
        if not blob:raise ValueError('Embedded relationship blob is missing for '+sid)
        raw=zlib.decompress(blob[0])
        if hashlib.sha256(raw).hexdigest()!=relationship_digest:
            raise ValueError('Embedded relationship blob is corrupt for '+sid)
        for rel in ET.fromstring(raw):
            if rel.get('TargetMode')=='External' or rel.get('Type') in SKIP_SOURCE_RELS:continue
            todo.append(_target(name,rel))
    payload={'version':1,'source_slide_id':sid,'file_id':file_id,'page':page,
             'slide_part':part,'slide_part_sha256':manifest[part],
             'visual_hash':visual_hash,'package_manifest_sha256':package_digest,
             'graph_parts':{name:graph[name] for name in sorted(graph)}}
    payload['sha256']=hashlib.sha256(json.dumps(payload,ensure_ascii=False,separators=(',',':'),sort_keys=True).encode('utf-8')).hexdigest()
    return payload

def sha(path):
    with path.open('rb') as f:
        return hashlib.file_digest(f, 'sha256').hexdigest()

def load(path):
    return json.loads(path.read_text(encoding='utf-8-sig'))

def queries_for(run):
    """Return only values that can change corpus search results.

    Layout and reuse metadata are deliberately excluded: reuse_slide.py updates those
    after retrieval, and that edit must not invalidate otherwise identical results.
    """
    request=load(run/'request.json');plan=load(run/'plan.json')
    specs=[];seen=set()
    def add(axis,kind,text):
        if not isinstance(text,str) or not text.strip():return
        key=(axis,kind,text.strip())
        if key not in seen:
            seen.add(key);specs.append({'axis':axis,'kind':kind,'query':text.strip()})
    add('content','slide',request.get('answers',{}).get('topic',{}).get('value',''))
    add('content','slide',request.get('title',''))
    for slide in plan.get('slides',[]):
        if slide.get('role')=='cover':continue
        # Whole-slide reuse is searched from both the teaching purpose and explicit title.
        add('content','slide',slide.get('purpose',''))
        add('content','slide',slide.get('title',''))
        # Visual lookup is a separate axis; it must not replace purpose/title retrieval.
        add('visual','any',slide.get('visual_query',''))
    return specs


def search_many(db,specs,limit=8):
    """Run all content/visual queries in one corpus scan and one JSON parse per slide."""
    prepared=[]
    for spec in specs:
        prepared.append((spec,[term for term in spec['query'].casefold().split() if term],[]))
    for sid,page,detail_text,title in db.execute(
            'SELECT s.id,s.page,s.detail,f.path FROM slides s JOIN files f ON f.id=s.file_id'):
        detail=json.loads(detail_text);stats=detail.get('stats',{});images=detail.get('images',[])
        available={'image':bool(images),'table':bool(stats.get('tables')),
                   'diagram':bool(stats.get('groups') or stats.get('connectors'))}
        text=detail.get('title','')+' '+detail.get('text','')+' '+detail.get('notes','')
        for spec,terms,found in prepared:
            kind=spec['kind']
            if kind not in ('any','slide') and not available[kind]:continue
            if kind!='slide' and not any(available.values()):continue
            score=sum(min(library.term_count(term,text),5)+3*bool(library.term_count(term,detail.get('title',''))) for term in terms)
            if not score:continue
            found.append({'slide_id':sid,'source_title':title.replace('\\','/').split('/')[-1],
                          'page':page,'score':score,'title':detail.get('title',''),
                          'excerpt':detail.get('text','')[:350],
                          'images':[{'index':i,'sha256':im.get('sha256'),'part':im.get('part')} for i,im in enumerate(images)],
                          'tables':stats.get('tables',0),'groups':stats.get('groups',0),
                          'connectors':stats.get('connectors',0)})
    result=[]
    for spec,_,found in prepared:
        found.sort(key=lambda row:(-row['score'],row['slide_id']))
        result.append({'axis':spec['axis'],'kind':spec['kind'],'query':spec['query'],
                       'matching_slides':len(found),'candidates':found[:max(1,limit)]})
    return result

def semantic_input_signature(run):
    payload=json.dumps({'search_axes':queries_for(run)},ensure_ascii=False,separators=(',',':'),sort_keys=True)
    return hashlib.sha256(payload.encode('utf-8')).hexdigest()

def corpus_state(db):
    decks=[];packaged=0;metadata_complete=0
    for row in db.execute('''SELECT f.id,f.path,f.size,f.mtime_ns,f.slide_count,f.status,f.error,p.parts
                             FROM files f LEFT JOIN packages p ON p.file_id=f.id ORDER BY f.id'''):
        fid,path,size,mtime,slides,status,error,parts=row
        try:manifest=json.loads(parts) if parts else None
        except (TypeError,ValueError):manifest=None
        manifest_valid=isinstance(manifest,dict) and bool(manifest) and all(
            isinstance(name,str) and isinstance(digest,str) and len(digest)==64 for name,digest in manifest.items())
        packaged+=manifest_valid
        complete=(isinstance(path,str) and bool(path) and isinstance(size,int) and size>0
                  and isinstance(mtime,int) and mtime>0 and isinstance(slides,int) and slides>=0
                  and status=='indexed' and not error)
        metadata_complete+=complete
        decks.append({'id':fid,'path':path,'size':size,'mtime_ns':mtime,'slide_count':slides,
                      'status':status,'error':error,
                      'package_manifest_sha256':hashlib.sha256((parts or '').encode('utf-8')).hexdigest()})
    slide_count=db.execute('SELECT COUNT(*) FROM slides').fetchone()[0]
    payload=json.dumps({'files':decks,'slide_count':slide_count},ensure_ascii=False,separators=(',',':'),sort_keys=True)
    return {'deck_count':len(decks),'slide_count':slide_count,
            'declared_slide_count':sum(d.get('slide_count') for d in decks if isinstance(d.get('slide_count'),int)),
            'indexed_decks':sum(d.get('status')=='indexed' for d in decks),
            'packaged_decks':packaged,'metadata_complete_decks':metadata_complete,
            'signature':hashlib.sha256(payload.encode('utf-8')).hexdigest()}

def require_complete_corpus(state):
    if (state.get('deck_count')!=EXPECTED_DECKS or state.get('indexed_decks')!=EXPECTED_DECKS
            or state.get('packaged_decks')!=EXPECTED_DECKS
            or state.get('metadata_complete_decks')!=EXPECTED_DECKS):
        raise ValueError(f'Embedded corpus must contain {EXPECTED_DECKS} indexed and packaged PPT files')
    if state.get('slide_count')!=EXPECTED_SLIDES or state.get('declared_slide_count')!=EXPECTED_SLIDES:
        raise ValueError(f'Embedded corpus must contain {EXPECTED_SLIDES} slides')

def _canonical_zip(source,dest):
    """Write a byte-stable PPTX so a DB export has one reproducible file hash."""
    with zipfile.ZipFile(source) as zin, zipfile.ZipFile(dest,'w',zipfile.ZIP_DEFLATED) as zout:
        for name in sorted(n for n in zin.namelist() if not n.endswith('/')):
            info=zipfile.ZipInfo(name,date_time=(1980,1,1,0,0,0))
            info.compress_type=zipfile.ZIP_DEFLATED;info.create_system=3;info.external_attr=0o600<<16
            zout.writestr(info,zin.read(name))

def refresh_native(db,sid,dest):
    """Re-export one source slide from the embedded DB and atomically replace any cache."""
    dest.parent.mkdir(parents=True,exist_ok=True)
    token=uuid.uuid4().hex
    raw=dest.with_name('.'+dest.name+'.'+token+'.raw')
    stable=dest.with_name('.'+dest.name+'.'+token+'.stable')
    try:
        library.export_slide(db,sid,raw)
        _canonical_zip(raw,stable)
        stable.replace(dest)
    finally:
        for path in (raw,stable):
            if path.exists():path.unlink()
    return sha(dest)

def refresh_image(db,sid,index,digest,dest):
    """Restore an image from its DB blob when the local cache is absent or changed."""
    if dest.exists() and sha(dest)==digest:return digest
    temp=dest.with_name('.'+dest.name+'.'+uuid.uuid4().hex+'.tmp'+dest.suffix)
    try:
        library.image_export(db,sid,index,temp)
        if sha(temp)!=digest:raise ValueError('Exported image bytes differ from corpus')
        temp.replace(dest)
    finally:
        if temp.exists():temp.unlink()
    return digest

def inspect(run):
    path=run/'retrieved-visuals/index.json'
    try:
        data=load(path)
        if (data.get('pipeline_version')!=PIPELINE_VERSION
                or data.get('semantic_input_signature')!=semantic_input_signature(run)
                or data.get('status')!='complete'):
            raise ValueError('Missing or stale full-corpus visual retrieval')
        if data.get('scope')!='all-ppt' or not data.get('queries'):
            raise ValueError('Full-corpus search was not completed')
        specs=queries_for(run)
        recorded_specs=[{k:q.get(k) for k in ('axis','kind','query')} for q in data.get('queries',[])]
        if recorded_specs!=specs:raise ValueError('Recorded retrieval queries differ from the current plan')
        with sqlite3.connect('file:'+library.DB.resolve().as_posix()+'?mode=ro',uri=True) as db:
            corpus=corpus_state(db)
            require_complete_corpus(corpus)
        if data.get('searched_decks')!=corpus['deck_count']:
            raise ValueError('Retrieved deck count does not match the embedded corpus')
        if data.get('corpus_signature')!=corpus['signature']:
            raise ValueError('Embedded corpus changed after retrieval')
        if data.get('searched_slides')!=corpus['slide_count']:
            raise ValueError('Retrieved slide count does not match the embedded corpus')
        for query in data['queries']:
            candidates=query.get('candidates')
            if not isinstance(candidates,list) or not isinstance(query.get('matching_slides'),int):
                raise ValueError('Retrieved candidate metadata is invalid')
            if query['matching_slides']>0 and not candidates:
                raise ValueError('A matching corpus query has no recorded candidates')
            if query['matching_slides']<len(candidates):
                raise ValueError('Retrieved candidate count exceeds corpus matches')
        for item in data.get('files',[]):
            p=(run/item['file']).resolve()
            if not p.is_relative_to(run.resolve()) or not p.is_file() or sha(p)!=item['sha256']:
                raise ValueError('Retrieved visual asset missing or changed')
    except (OSError, ValueError, KeyError, TypeError, sqlite3.Error) as exc:
        return [{'rule':'RETRIEVE-01','message':str(exc)+'; run plan-check to retrieve actual corpus assets'}]
    return []

def collect(run):
    if not inspect(run):
        return load(run/'retrieved-visuals/index.json')
    query_specs=queries_for(run)
    if not query_specs:
        raise ValueError('Provide a lecture topic before retrieving visual assets')
    output=run/'retrieved-visuals';output.mkdir(exist_ok=True)
    result={'pipeline_version':PIPELINE_VERSION,'status':'collecting','scope':'all-ppt',
            'semantic_input_signature':semantic_input_signature(run),'queries':[], 'files':[]}
    seen_images=set()
    with sqlite3.connect('file:'+library.DB.resolve().as_posix()+'?mode=ro',uri=True) as db:
        corpus=corpus_state(db)
        require_complete_corpus(corpus)
        result['searched_decks']=corpus['deck_count']
        result['searched_slides']=corpus['slide_count']
        result['corpus_signature']=corpus['signature']
        result['queries']=search_many(db,query_specs,8)
        for query_result in result['queries']:
            for candidate in query_result['candidates']:
                sid=candidate['slide_id']
                for image in candidate['images']:
                    digest=image['sha256']
                    if digest in seen_images:continue
                    seen_images.add(digest)
                    suffix=pathlib.PurePosixPath(image['part']).suffix
                    dest=output/(digest+suffix)
                    refresh_image(db,sid,image['index'],digest,dest)
                    result['files'].append({'file':dest.relative_to(run).as_posix(),'sha256':digest,'slide_id':sid,'source_title':candidate['source_title'],'page':candidate['page'],'kind':'image'})
    result['status']='complete'
    (output/'index.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    return result
