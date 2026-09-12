"""Search visual assets across every embedded lecture deck; export original image bytes."""
import argparse,hashlib,json,os,pathlib,sqlite3,sys,zlib
from lecture_library import bm25_rank, term_count
from slide_store import export as export_slide
BASE=pathlib.Path(__file__).resolve().parents[1]
# Keep user-built corpus data outside the plugin folder.
DB=pathlib.Path(os.environ.get('CDSA_LECTURE_DB') or BASE/'data/lecture-library.sqlite')
def emit_json(value,output=None):
    payload=json.dumps(value,ensure_ascii=False,indent=2)
    if output:
        output.parent.mkdir(parents=True,exist_ok=True);output.write_text(payload+'\n',encoding='utf-8')
    else:print(payload)
def search(db,query,kind,limit):
    candidates=[]
    for sid,page,detail,title in db.execute('SELECT s.id,s.page,s.detail,f.path FROM slides s JOIN files f ON f.id=s.file_id'):
        d=json.loads(detail);stats=d.get('stats',{});images=d.get('images',[])
        available={'image':bool(images),'table':bool(stats.get('tables')),'diagram':bool(stats.get('groups') or stats.get('connectors'))}
        if kind not in ('any','slide') and not available[kind]:continue
        if kind!='slide' and not any(available.values()):continue
        text=d.get('title','')+' '+d.get('text','')+' '+d.get('notes','')
        candidates.append({'id':sid,'slide_id':sid,'source_title':title.replace('\\','/').split('/')[-1],
                           'page':page,'title':d.get('title',''),'text':text,
                           'excerpt':d.get('text','')[:350],
                           'images':[{'index':i,'sha256':im.get('sha256'),'part':im.get('part')} for i,im in enumerate(images)],
                           'tables':stats.get('tables',0),'groups':stats.get('groups',0),'connectors':stats.get('connectors',0)})
    ranked=bm25_rank(candidates,query)
    if not ranked:
        terms=query.casefold().split();ranked=[]
        for row in candidates:
            score=sum(min(term_count(term,row['text']),5)+3*bool(term_count(term,row['title'])) for term in terms)
            if score:ranked.append((score,row))
        ranked.sort(key=lambda pair:(-pair[0],pair[1]['slide_id']))
    results=[]
    for score,row in ranked[:limit]:
        item={key:value for key,value in row.items() if key not in ('id','text')};item['score']=round(score,4);results.append(item)
    return {'scope':'all embedded PPT files, no Jeju filter','matching_slides':len(ranked),'results':results,'search_basis':'BM25 over slide title, text and notes; image pixels are not OCRed','selection_warning':'검색 점수는 후보 순위일 뿐이다. 계획한 장표 목적과 제목·본문·시각 구조를 직접 대조한 뒤 선택한다.'}
def image_export(db,sid,index,out):
    row=db.execute('SELECT file_id,detail FROM slides WHERE id=?',(sid,)).fetchone()
    if not row:raise ValueError('Unknown slide ID')
    images=json.loads(row[1]).get('images',[])
    if index<0 or index>=len(images):raise ValueError('Image index out of range')
    im=images[index];manifest=json.loads(db.execute('SELECT parts FROM packages WHERE file_id=?',(row[0],)).fetchone()[0]);digest=manifest[im['part']]
    raw=zlib.decompress(db.execute('SELECT data FROM blobs WHERE sha=?',(digest,)).fetchone()[0])
    if hashlib.sha256(raw).hexdigest()!=digest or im.get('sha256') not in (None,digest):raise ValueError('Image integrity mismatch')
    if out.exists():raise ValueError('Output exists')
    suffix=pathlib.PurePosixPath(im['part']).suffix
    if out.suffix.lower()!=suffix.lower():raise ValueError('Use original image extension: '+suffix)
    out.parent.mkdir(parents=True,exist_ok=True);out.write_bytes(raw)
    return {'slide_id':sid,'image_index':index,'output':str(out),'sha256':digest,'original_files_required':False}
def main():
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('--db',type=pathlib.Path,default=DB);sub=p.add_subparsers(dest='command',required=True)
    s=sub.add_parser('search');s.add_argument('query');s.add_argument('--kind',choices=['any','slide','image','table','diagram'],default='any');s.add_argument('--limit',type=int,default=8);s.add_argument('--output',type=pathlib.Path)
    x=sub.add_parser('image');x.add_argument('slide_id');x.add_argument('--index',type=int,required=True);x.add_argument('--output',type=pathlib.Path,required=True)
    x=sub.add_parser('slide');x.add_argument('slide_id');x.add_argument('--output',type=pathlib.Path,required=True)
    a=p.parse_args()
    if not a.db.is_file():raise ValueError('Corpus database not found. Run <plugin>/scripts/setup-corpus.ps1 or set CDSA_LECTURE_DB: '+str(a.db))
    with sqlite3.connect('file:'+a.db.resolve().as_posix()+'?mode=ro',uri=True) as db:
        result=search(db,a.query,a.kind,max(1,a.limit)) if a.command=='search' else image_export(db,a.slide_id,a.index,a.output) if a.command=='image' else export_slide(db,a.slide_id,a.output)
    # Search writes JSON to --output. Binary export commands must preserve the
    # image/PPTX at --output and print their manifest to stdout instead.
    emit_json(result,a.output if a.command=='search' else None)
if __name__=='__main__':
    if hasattr(sys.stdout,'reconfigure'):sys.stdout.reconfigure(encoding='utf-8')
    main()
