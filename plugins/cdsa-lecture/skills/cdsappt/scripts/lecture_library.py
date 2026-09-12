"""Read-only PPTX corpus indexing, conservative deduplication, and local retrieval."""
import argparse
import collections
import datetime
import difflib
import hashlib
import json
import os
import pathlib
import posixpath
import re
import sqlite3
import sys
import unicodedata
import zipfile
import xml.etree.ElementTree as E
import math

BASE = pathlib.Path(__file__).resolve().parents[1]
# Keep user-built corpus data outside the plugin folder.
DEFAULT_DB = pathlib.Path(os.environ.get('CDSA_LECTURE_DB') or BASE / 'data' / 'lecture-library.sqlite')
NS = {'a':'http://schemas.openxmlformats.org/drawingml/2006/main','p':'http://schemas.openxmlformats.org/presentationml/2006/main'}
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PATTERNS = ['table','chart','diagram','image-led','image-text','comparison','text-led','sparse']

def norm(s):
    return re.sub(r'\s+','',unicodedata.normalize('NFKC',s)).casefold()

def digest(s):
    return hashlib.sha256(s.encode('utf-8')).hexdigest()

def term_count(term,text):
    if re.fullmatch(r'[a-z0-9_]+',term):
        return len(re.findall(r'(?<![a-z0-9_])'+re.escape(term)+r'(?![a-z0-9_])',text.casefold()))
    return norm(text).count(term)

SEARCH_TOKEN_RE = re.compile(r'[가-힣]+|[a-z0-9_]+')


def _iter_search_tokens(value):
    """Yield NFKC Korean 2-grams and lower-cased Latin/number words."""
    value = unicodedata.normalize('NFKC', value or '').casefold()
    for match in SEARCH_TOKEN_RE.finditer(value):
        chunk = match.group(0)
        if '가' <= chunk[0] <= '힣':
            for index in range(max(1, len(chunk) - 1)):
                yield chunk[index:index + 2]
        else:
            yield chunk


def search_tokens(value):
    return list(_iter_search_tokens(value))


def _relevant_token_counts(value, wanted):
    counts = collections.Counter()
    length = 0
    for token in _iter_search_tokens(value):
        length += 1
        if token in wanted:
            counts[token] += 1
    return counts, length

def bm25_rank(records, query, limit=None):
    """Rank dict records with id/title/text using BM25 without mutating the DB."""
    query_counts = collections.Counter(search_tokens(query))
    if not query_counts:
        return []
    wanted = set(query_counts)
    prepared = []
    document_frequency = collections.Counter()
    for record in records:
        counts, text_length = _relevant_token_counts(record.get('text', ''), wanted)
        # Title occurrences carry two additional term occurrences.
        title_counts, title_length = _relevant_token_counts(record.get('title', ''), wanted)
        counts.update({term: count * 2 for term, count in title_counts.items()})
        length = text_length + title_length * 2 or 1
        prepared.append((record, counts, length))
        document_frequency.update(counts.keys())
    total = len(prepared); average = sum(row[2] for row in prepared) / max(1, total)
    results = []
    for record, counts, length in prepared:
        score = 0.0
        for term, query_frequency in query_counts.items():
            frequency = counts.get(term, 0)
            if not frequency:
                continue
            inverse = math.log(1 + (total - document_frequency[term] + 0.5) / (document_frequency[term] + 0.5))
            score += query_frequency * inverse * frequency * 2.2 / (frequency + 1.2 * (0.25 + 0.75 * length / average))
        if score:
            results.append((score, record))
    results.sort(key=lambda row: (-row[0], str(row[1].get('id', ''))))
    return results if limit is None else results[:limit]

def relations(z, part):
    rp=posixpath.join(posixpath.dirname(part),'_rels',posixpath.basename(part)+'.rels')
    try: root=E.fromstring(z.read(rp))
    except KeyError:return {}
    return {r.get('Id'):{'kind':r.get('Type').split('/')[-1], 'external':r.get('TargetMode')=='External', 'target':r.get('Target') if r.get('TargetMode')=='External' else posixpath.normpath(posixpath.join(posixpath.dirname(part),r.get('Target'))).lstrip('/')} for r in root}

def paragraphs(el):
    return [''.join(t.text or '' for t in p.findall('.//a:t',NS)) for p in el.findall('.//a:p',NS)]

def extract(path):
    slides=[]
    with zipfile.ZipFile(path) as z:
        pres=E.fromstring(z.read('ppt/presentation.xml'));pr=relations(z,'ppt/presentation.xml')
        size=pres.find('p:sldSz',NS);w,h=int(size.get('cx')),int(size.get('cy'))
        media_cache={}
        for page,sid in enumerate(pres.findall('p:sldIdLst/p:sldId',NS),1):
            part=pr[sid.get('{'+REL+'}id')]['target'];e=E.fromstring(z.read(part));rr=relations(z,part)
            shapes=[];content=[];edge=[];images=[];evidence=[];notes=[]
            for s in e.findall('p:cSld/p:spTree/*',NS):
                kind=s.tag.split('}')[-1]
                if kind in ('nvGrpSpPr','grpSpPr'):continue
                texts=paragraphs(s);text='\n'.join(t for t in texts if t.strip())
                nv=s.find('.//p:cNvPr',NS);ph=s.find('.//p:ph',NS)
                xf=s.find('.//a:xfrm',NS)
                if xf is None:xf=s.find('p:xfrm',NS)
                box=None
                if xf is not None:
                    off=xf.find('a:off',NS);ext=xf.find('a:ext',NS)
                    if off is not None and ext is not None:
                        box=[round(int(off.get('x'))/w,5),round(int(off.get('y'))/h,5),round(int(ext.get('cx'))/w,5),round(int(ext.get('cy'))/h,5)]
                isedge=bool(box and (box[1]+box[3]<=0.095 or box[1]>=0.93))
                isfooter=ph is not None and ph.get('type') in ('dt','ftr','sldNum','hdr')
                if text and not isfooter:
                    (edge if isedge else content).append(text)
                shapes.append({'kind':kind,'name':nv.get('name') if nv is not None else '', 'box_normalized':box,'text':text,'placeholder':ph.attrib if ph is not None else None})
            for rel in rr.values():
                if rel['external']:continue
                t=rel['target']
                if rel['kind'] in ('image','chart','diagramData','oleObject','video','audio','media'):
                    if t not in media_cache:media_cache[t]=hashlib.sha256(z.read(t)).hexdigest()
                    item={'kind':rel['kind'],'part':t,'sha256':media_cache[t]}
                    (images if rel['kind']=='image' else evidence).append(item)
                    if rel['kind'] in ('chart','diagramData'):
                        try:
                            er=E.fromstring(z.read(t));embedded=[el.text for el in er.iter() if el.tag.split('}')[-1] in ('t','v') and el.text]
                            if embedded:content.append('[Embedded data]\n'+'\n'.join(embedded))
                        except E.ParseError:pass
                elif rel['kind']=='notesSlide':
                    nr=E.fromstring(z.read(t))
                    for sp in nr.findall('p:cSld/p:spTree/p:sp',NS):
                        ph=sp.find('.//p:ph',NS)
                        if ph is None or ph.get('type')=='body':notes.extend(paragraphs(sp))
            tables=len(e.findall('.//a:tbl',NS));connectors=len(e.findall('.//p:cxnSp',NS));groups=len(e.findall('.//p:grpSp',NS))
            img_area=sum(s['box_normalized'][2]*s['box_normalized'][3] for s in shapes if s['kind']=='pic' and s['box_normalized'])
            visible_text='\n'.join(content);ntext=len(norm(visible_text))
            nonflat_connectors=sum(s['kind']=='cxnSp' and bool(s['box_normalized']) and s['box_normalized'][3]>0.015 for s in shapes)
            pattern='table' if tables else 'chart' if any(v['kind']=='chart' for v in evidence) else 'diagram' if nonflat_connectors>=2 or any(v['kind']=='diagramData' for v in evidence) else 'image-led' if img_area>=0.5 else 'image-text' if images and ntext>=60 else 'comparison' if len([s for s in shapes if s['text'] and s['box_normalized'] and 0.25<s['box_normalized'][2]<0.55])>=2 else 'text-led' if ntext>=120 else 'sparse'
            slides.append({'page':page,'part':part,'content_blocks':content,'edge_blocks':edge,'notes':'\n'.join(notes),'images':images,'evidence':evidence,'shapes':shapes,'pattern':pattern,'hidden':e.get('show') in ('0','false'),'layout':next((r['target'] for r in rr.values() if r['kind']=='slideLayout'),None),'stats':{'tables':tables,'connectors':connectors,'groups':groups,'image_area':round(img_area,3)}})
        # Strip only repeated top/bottom furniture. Unique edge text remains searchable.
        count=collections.Counter(norm(t) for s in slides for t in set(s['edge_blocks']))
        for s in slides:
            keep=[t for t in s['edge_blocks'] if count[norm(t)]<max(3,len(slides)*0.25)]
            s['text']='\n'.join(keep+s.pop('content_blocks'))
            s['edge_blocks_removed']=[t for t in s.pop('edge_blocks') if t not in keep]
            s['title']=next((t.strip()[:160] for t in s['text'].splitlines() if t.strip()),'[visual slide]')
    return slides

def schema(db):
    db.executescript('''
    CREATE TABLE files(id INTEGER PRIMARY KEY,path TEXT UNIQUE,size INTEGER,mtime_ns INTEGER,slide_count INTEGER,status TEXT,error TEXT);
    CREATE TABLE groups(id TEXT PRIMARY KEY,title TEXT,text TEXT,pattern TEXT,needs_visual INTEGER);
    CREATE TABLE slides(id TEXT PRIMARY KEY,file_id INTEGER,page INTEGER,group_id TEXT,visual_hash TEXT,notes TEXT,detail TEXT);
    CREATE INDEX sg ON slides(group_id);
    CREATE TABLE related(a TEXT,b TEXT,score REAL,reason TEXT,PRIMARY KEY(a,b));
    CREATE TABLE meta(key TEXT PRIMARY KEY,value TEXT);
    ''')

def build(root, output):
    output.parent.mkdir(parents=True,exist_ok=True)
    staging=output.with_suffix('.building.sqlite')
    if staging.exists():raise ValueError('An unfinished build exists: '+str(staging))
    db=sqlite3.connect(staging);schema(db)
    paths=sorted(p for p in root.rglob('*') if p.is_file() and p.suffix.lower() in ('.pptx','.ppt','.pptm') and not p.name.startswith('~$'))
    errors=[];total=0
    for index,path in enumerate(paths,1):
        st=path.stat()
        try:
            if path.suffix.lower()=='.ppt':raise ValueError('Legacy PPT requires read-only conversion to PPTX before indexing')
            slides=extract(path)
            db.execute('INSERT INTO files VALUES(?,?,?,?,?,?,?)',(index,str(path),st.st_size,st.st_mtime_ns,len(slides),'indexed',None))
            for s in slides:
                normalized=norm(s['text'])
                visual=digest(json.dumps({'media':sorted(v['sha256'] for v in s['images']+s['evidence']),'geometry':[(v['kind'],v['box_normalized']) for v in s['shapes']]},sort_keys=True))
                # Long identical content forms one group with all visual/source variants retained.
                # Short/generic headings cannot collapse unrelated diagrams or screenshots.
                key=normalized if len(normalized)>=80 else normalized+'|'+visual
                if not normalized and not s['images'] and not s['evidence']:
                    key+='|'+str(path)+'|'+str(s['page'])
                gid='g-'+digest(key)[:20];sid='s-'+digest(str(path)+'#'+str(s['page']))[:20]
                db.execute('INSERT OR IGNORE INTO groups VALUES(?,?,?,?,?)',(gid,s['title'],s['text'],s['pattern'],int(len(normalized)<80 and bool(s['images'] or s['evidence']))))
                db.execute('INSERT INTO slides VALUES(?,?,?,?,?,?,?)',(sid,index,s['page'],gid,visual,s['notes'],json.dumps(s,ensure_ascii=False)))
            total+=len(slides);db.commit()
            print(f'[{index}/{len(paths)}] {path.name}: {len(slides)} slides',flush=True)
        except Exception as exc:
            db.rollback()
            db.execute('INSERT OR REPLACE INTO files VALUES(?,?,?,?,?,?,?)',(index,str(path),st.st_size,st.st_mtime_ns,0,'error',str(exc)));db.commit()
            errors.append({'path':str(path),'error':str(exc)});print(f'ERROR {path.name}: {exc}',flush=True)
    # Near-duplicates are review candidates, never silently deleted or merged.
    rows=db.execute('SELECT id,title,text FROM groups WHERE length(text)>=80').fetchall()
    buckets=collections.defaultdict(list);pairs=set();related=0
    for gid,title,text in rows:
        txt=norm(text)
        keys={norm(title)[:32]} if len(norm(title))>=8 else set()
        words=set(re.findall(r'[\w가-힣]{3,}',text.casefold()))
        # Five rare-ish deterministic token hashes increase recall across title edits.
        keys.update('h:'+str(v) for v in sorted(int(hashlib.blake2s(w.encode(),digest_size=4).hexdigest(),16) for w in words)[:5])
        candidates={other for k in keys for other in buckets[k][-120:]}
        for oid,otxt in candidates:
            if min(len(txt),len(otxt))/max(len(txt),len(otxt))<0.78:continue
            pair=tuple(sorted((gid,oid)))
            if pair in pairs:continue
            pairs.add(pair)
            score=difflib.SequenceMatcher(None,txt,otxt,autojunk=False).ratio()
            if score>=0.88:
                db.execute('INSERT OR IGNORE INTO related VALUES(?,?,?,?)',(*pair,round(score,4),'near-text; inspect differences in numbers, claims, examples and images'));related+=1
        for k in keys:buckets[k].append((gid,txt))
    summary={'built_at':datetime.datetime.now().astimezone().isoformat(),'source_root':str(root),'files':len(paths),'indexed_files':len(paths)-len(errors),'slides':total,'content_groups':db.execute('SELECT count(*) FROM groups').fetchone()[0],'near_duplicate_pairs':related,'visual_review_groups':db.execute('SELECT count(*) FROM groups WHERE needs_visual=1').fetchone()[0],'errors':errors,'patterns':dict(db.execute('SELECT pattern,count(*) FROM groups GROUP BY pattern')),'dedup':'normalized identical text >=80 chars grouped with source/visual variants retained; shorter text requires matching media/geometry; near-text candidates are not merged'}
    summary['repeated_occurrences']=total-summary['content_groups']
    db.execute('INSERT INTO meta VALUES(?,?)',('summary',json.dumps(summary,ensure_ascii=False)));db.commit();db.close()
    staging.replace(output)
    output.with_suffix('.summary.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(summary,ensure_ascii=False,indent=2))

def source_refs(db,gid):
    refs=[dict(zip(('slide_id','path','page','visual_hash'),r)) for r in db.execute('SELECT s.id,f.path,s.page,s.visual_hash FROM slides s JOIN files f ON f.id=s.file_id WHERE s.group_id=? ORDER BY f.mtime_ns DESC,s.page',(gid,))]
    for ref in refs:
        ref['source_title']=ref.pop('path').rsplit('/',1)[-1]
        ref['resource']='ssj-slide://'+ref['slide_id']
    return refs

def emit_json(value, output=None):
    payload=json.dumps(value,ensure_ascii=False,indent=2)
    if output:
        output.parent.mkdir(parents=True,exist_ok=True)
        output.write_text(payload+'\n',encoding='utf-8')
    else:
        print(payload)

def main():
    ap=argparse.ArgumentParser(description=__doc__);ap.add_argument('--db',type=pathlib.Path,default=DEFAULT_DB)
    sub=ap.add_subparsers(dest='command',required=True)
    b=sub.add_parser('build');b.add_argument('--source',required=True,type=pathlib.Path)
    q=sub.add_parser('search');q.add_argument('query');q.add_argument('--limit',type=int,default=8);q.add_argument('--pattern',choices=PATTERNS);q.add_argument('--output',type=pathlib.Path)
    g=sub.add_parser('show');g.add_argument('id');g.add_argument('--shapes',action='store_true');g.add_argument('--limit',type=int,default=3);g.add_argument('--output',type=pathlib.Path)
    sub.add_parser('stats')
    args=ap.parse_args()
    if args.command=='build':return build(args.source,args.db)
    if not args.db.is_file():
        raise ValueError('Corpus database not found. Run <plugin>/scripts/setup-corpus.ps1 or set CDSA_LECTURE_DB: '+str(args.db))
    db=sqlite3.connect('file:'+args.db.as_posix()+'?mode=ro',uri=True)
    if args.command=='stats':print(db.execute("SELECT value FROM meta WHERE key='summary'").fetchone()[0]);return
    if args.command=='search':
        records=[]
        for gid,title,text,pattern,visual in db.execute('SELECT * FROM groups'):
            if args.pattern and pattern!=args.pattern:continue
            records.append({'id':gid,'title':title,'text':text,'pattern':pattern,'visual':visual})
        result=[]
        ranked=bm25_rank(records,args.query,args.limit)
        # Short literal queries sometimes work better with exact substring matching.
        if not ranked:
            terms=[norm(t) for t in args.query.split() if norm(t)]
            fallback=[]
            for record in records:
                counts=[term_count(term,record['text']) for term in terms]
                if any(counts):fallback.append((sum(counts),record))
            ranked=sorted(fallback,key=lambda row:(-row[0],row[1]['id']))[:args.limit]
        for score,record in ranked:
            gid,title,text,pattern,visual=(record[key] for key in ('id','title','text','pattern','visual'))
            refs=source_refs(db,gid)
            result.append({'id':gid,'score':round(score,2),'title':title,'excerpt':text[:1000],'pattern_candidate':pattern,'needs_visual_review':True,'text_insufficient':bool(visual),'occurrences':len(refs),'visual_variants':len({r['visual_hash'] for r in refs}),'sources':refs[:5],'related':[r[0] for r in db.execute('SELECT CASE WHEN a=? THEN b ELSE a END FROM related WHERE a=? OR b=?',(gid,gid,gid))]})
        emit_json(result,args.output)
    else:
        gid=args.id
        if gid.startswith('s-'):
            row=db.execute('SELECT group_id FROM slides WHERE id=?',(gid,)).fetchone()
            if not row:raise ValueError('Unknown slide ID')
            gid=row[0]
        row=db.execute('SELECT * FROM groups WHERE id=?',(gid,)).fetchone()
        if not row:raise ValueError('Unknown group ID')
        result=dict(zip(('id','title','text','pattern_candidate','text_insufficient'),row));result['needs_visual_review']=True;refs=source_refs(db,gid);result['source_count']=len(refs);result['sources']=refs[:args.limit]
        result['variants']=[]
        selected=[args.id] if args.id.startswith('s-') else [r['slide_id'] for r in refs[:args.limit]]
        for sid in selected:
            detail,notes=db.execute('SELECT detail,notes FROM slides WHERE id=?',(sid,)).fetchone()
            s=json.loads(detail)
            if not args.shapes:s.pop('shapes',None)
            result['variants'].append({'slide_id':sid,'detail':s,'notes':notes})
        emit_json(result,args.output)

if __name__=='__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    main()
