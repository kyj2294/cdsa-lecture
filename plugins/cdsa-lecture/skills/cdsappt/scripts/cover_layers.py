"""Reuse original editable cover capsule, rules, title and subtitle over a generated scene."""
import argparse,copy,zipfile
from pathlib import Path
import xml.etree.ElementTree as E
import jeju_template as jt
P,A=jt.P,'http://schemas.openxmlformats.org/drawingml/2006/main'
NS={'p':P,'a':A}
def apply(files,title,course,subtitle):
    part='ppt/slides/slide1.xml';e=jt.xml(files[part]);tree=e.find('p:cSld/p:spTree',NS)
    # Caller provides a cover with its background picture only. Prevent stacked old titles.
    if any(s.tag in ('{'+P+'}sp','{'+P+'}grpSp','{'+P+'}cxnSp') for s in tree):
        raise ValueError('Use a cover containing the background picture only; remove previous body titles explicitly')
    original=E.parse(Path(__file__).resolve().parents[1]/'assets/jeju-body-patterns/cover-original.xml').getroot()
    wanted={'14':None,'19':None,'3':title,'5':course,'30':subtitle}
    nextid=max([int(n.get('id','0')) for n in e.findall('.//p:cNvPr',NS)]+[100])+1
    for sp in original.find('p:cSld/p:spTree',NS):
        nv=sp.find('.//p:cNvPr',NS)
        if nv is None or nv.get('id') not in wanted:continue
        old=nv.get('id');s=copy.deepcopy(sp)
        for n in s.findall('.//p:cNvPr',NS):n.set('id',str(nextid));nextid+=1
        if wanted[old] is not None:
            tx=s.find('p:txBody',NS);first=tx.find('a:p/a:r',NS);rpr=copy.deepcopy(first.find('a:rPr',NS)) if first is not None else E.Element('{'+A+'}rPr')
            for child in list(tx):
                if child.tag=='{'+A+'}p':tx.remove(child)
            pa=E.SubElement(tx,'{'+A+'}p');pp=E.SubElement(pa,'{'+A+'}pPr',algn='ctr' if old=='5' else 'l',marL='0',indent='0')
            r=E.SubElement(pa,'{'+A+'}r');r.append(rpr);E.SubElement(r,'{'+A+'}t').text=wanted[old]
            body=tx.find('a:bodyPr',NS)
            for ch in list(body):
                if ch.tag.endswith('spAutoFit'):body.remove(ch)
            if old=='5':body.attrib.update(anchor='ctr',lIns='60000',rIns='60000',tIns='0',bIns='0')
            if old=='30':
                body.attrib.update(tIns='0',bIns='0')
                xfrm=s.find('p:spPr/a:xfrm',NS);ext=xfrm.find('a:ext',NS) if xfrm is not None else None
                if ext is not None:ext.set('cy',str(max(int(ext.get('cy','0')),640080)))
        tree.append(s)
    files[part]=jt.serialize(e);return files
if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--input',type=Path,required=True);p.add_argument('--output',type=Path,required=True);p.add_argument('--title',required=True);p.add_argument('--course',required=True);p.add_argument('--subtitle',default='');a=p.parse_args()
    if a.output.exists():raise ValueError('Output exists')
    files=apply(jt.read(a.input),a.title,a.course,a.subtitle)
    with zipfile.ZipFile(a.output,'w',zipfile.ZIP_DEFLATED) as z:
        for n,b in files.items():z.writestr(n,b)
