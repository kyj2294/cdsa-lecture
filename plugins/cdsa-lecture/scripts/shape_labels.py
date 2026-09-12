"""Deterministic checks for compact rounded label shapes, including nested groups."""
import argparse,json,pathlib,zipfile,xml.etree.ElementTree as E,posixpath
P='http://schemas.openxmlformats.org/presentationml/2006/main';A='http://schemas.openxmlformats.org/drawingml/2006/main';R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS={'p':P,'a':A};PT=12700

def flatten(container,transform=(1,1,0,0),path=()):
    sx,sy,tx,ty=transform
    for s in container:
        if s.tag=='{'+P+'}grpSp':
            xf=s.find('p:grpSpPr/a:xfrm',NS)
            if xf is None:continue
            off,ext,co,ce=[xf.find('a:'+n,NS) for n in ('off','ext','chOff','chExt')]
            if any(v is None for v in (off,ext,co,ce)):continue
            if int(ce.get('cx','0'))==0 or int(ce.get('cy','0'))==0:continue
            gx=int(ext.get('cx'))/int(ce.get('cx'));gy=int(ext.get('cy'))/int(ce.get('cy'))
            nested=(sx*gx,sy*gy,tx+sx*(int(off.get('x'))-gx*int(co.get('x'))),ty+sy*(int(off.get('y'))-gy*int(co.get('y'))))
            nv=s.find('p:nvGrpSpPr/p:cNvPr',NS)
            yield from flatten(s,nested,path+(nv.get('id','?') if nv is not None else '?',))
        elif s.tag=='{'+P+'}sp':
            xf=s.find('p:spPr/a:xfrm',NS);nv=s.find('p:nvSpPr/p:cNvPr',NS)
            if xf is None or nv is None:continue
            o=xf.find('a:off',NS);ex=xf.find('a:ext',NS)
            if o is None or ex is None:continue
            x,y=tx+sx*int(o.get('x')),ty+sy*int(o.get('y'))
            w,h=sx*int(ex.get('cx')),sy*int(ex.get('cy'))
            yield {'el':s,'id':nv.get('id'),'name':nv.get('name',''),'box':(x/PT,y/PT,w/PT,h/PT),'path':path,'text':''.join(t.text or '' for t in s.findall('p:txBody//a:t',NS))}

def compact_round(s):
    e=s['el'];g=e.find('p:spPr/a:prstGeom',NS);x,y,w,h=s['box']
    return g is not None and g.get('prst')=='roundRect' and 40<=w<=360 and 18<=h<=65 and 1.2<=w/h<=9 and e.find('p:spPr/a:noFill',NS) is None

def inspect_slide(e,page,exclusions=()):
    """Exclusions must be specific IDs with documented intentional non-label purpose."""
    issues=[];tree=e.find('p:cSld/p:spTree',NS)
    if tree is None:return issues
    items=list(flatten(tree));excluded={str(v['shape_id']) for v in exclusions if isinstance(v,dict) and v.get('shape_id') and str(v.get('reason','')).strip()}
    def issue(code,s,msg,other=None):
        v={'rule':code,'page':page,'shape_id':s['id'],'shape_name':s['name'],'group_path':list(s['path']),'message':msg}
        if other is not None:v['label_shape_id']=other['id']
        issues.append(v)
    for s in items:
        if s['id'] in excluded or not compact_round(s):continue
        text=s['text'].strip();x,y,w,h=s['box']
        if not text:
            for label in items:
                if label['id']==s['id'] or not label['text'].strip() or len(label['text'])>45:continue
                lx,ly,lw,lh=label['box'];cx,cy=lx+lw/2,ly+lh/2
                # A separate label positioned in a compact rounded background.
                overlap=max(0,min(x+w,lx+lw)-max(x,lx))*max(0,min(y+h,ly+lh)-max(y,ly))
                if x<=cx<=x+w and y<=cy<=y+h and lw*lh>0 and overlap/(lw*lh)>=0.75:
                    issue('LABEL-01',s,'Separate text box over label shape; move the text into the shape',label)
            continue
        if len(text)>45:continue
        bp=s['el'].find('p:txBody/a:bodyPr',NS)
        if bp is None or bp.get('anchor','t')!='ctr':issue('LABEL-02',s,'Label must have vertical center anchoring')
        paras=s['el'].findall('p:txBody/a:p',NS)
        for p in paras:
            if not ''.join(t.text or '' for t in p.findall('.//a:t',NS)).strip():continue
            pp=p.find('a:pPr',NS)
            if pp is None or pp.get('algn','l')!='ctr' or any(int(pp.get(k,'0'))!=0 for k in ('marL','marR','indent')):
                issue('LABEL-03',s,'Label paragraphs must be explicitly centered with zero indent');break
        if bp is not None:
            left,right=int(bp.get('lIns','91440')),int(bp.get('rIns','91440'))
            top,bottom=int(bp.get('tIns','45720')),int(bp.get('bIns','45720'))
            if abs(left-right)>PT/4 or abs(top-bottom)>PT/4:issue('LABEL-04',s,'Label insets are asymmetric by more than 0.25 pt')
            if bp.find('a:spAutoFit',NS) is not None:issue('LABEL-05',s,'Shape auto-resize can move the label; use a fixed label shape')
    return issues

def inspect_deck(path,pages=None):
    with zipfile.ZipFile(path) as z:
        rels={r.get('Id'):r.get('Target') for r in E.fromstring(z.read('ppt/_rels/presentation.xml.rels'))}
        e=E.fromstring(z.read('ppt/presentation.xml'));issues=[]
        for i,s in enumerate(e.find('p:sldIdLst',NS),1):
            if pages and i not in pages:continue
            target=rels[s.get('{'+R+'}id')];part=target.lstrip('/') if target.startswith('/') else posixpath.normpath('ppt/'+target)
            issues.extend(inspect_slide(E.fromstring(z.read(part)),i))
    return {'status':'BLOCKED' if issues else 'PASS','scope':'compact rounded label geometry, not overall slide quality','issues':issues}

if __name__=='__main__':
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('deck',type=pathlib.Path);p.add_argument('--pages',help='Comma-separated slide numbers');p.add_argument('--output',type=pathlib.Path);a=p.parse_args()
    r=inspect_deck(a.deck,{int(v) for v in a.pages.split(',')} if a.pages else None);text=json.dumps(r,ensure_ascii=False,indent=2)
    if a.output:a.output.write_text(text,encoding='utf-8')
    print(text);raise SystemExit(r['status']!='PASS')
