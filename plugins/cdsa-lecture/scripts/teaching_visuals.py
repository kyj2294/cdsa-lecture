"""Deterministic table layout checks: vertical centering, header alignment, oversized tables.

Whether an exercise slide has the *right* picture is a visual_fit review item, not a code rule.
"""
import xml.etree.ElementTree as E
NS = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
PT = 12700

def inspect_slide(e, page, item=None):
    item = item or {}
    issues = []
    def fail(code, msg, shape=None):
        issues.append({'rule': code, 'page': page, 'shape_id': shape, 'message': msg})
    for f in e.findall('.//p:graphicFrame', NS):
        tbl = f.find('.//a:tbl', NS)
        if tbl is None:
            continue
        sid = f.find('p:nvGraphicFramePr/p:cNvPr', NS).get('id')
        rows = tbl.findall('a:tr', NS); roomy = 0; measured = 0
        for ri, row in enumerate(rows):
            height = int(row.get('h', '0')) / PT; short = True; largest = 0
            for ci, cell in enumerate(row.findall('a:tc', NS)):
                text = ''.join(t.text or '' for t in cell.findall('.//a:t', NS)).strip()
                if not text:
                    continue
                pr = cell.find('a:tcPr', NS)
                if pr is None or pr.get('anchor', 't') != 'ctr':
                    fail('TABLE-ALIGN-01', f'Cell {ri+1},{ci+1} must use vertical center', sid)
                paras = cell.findall('a:txBody/a:p', NS)
                if (ri == 0 or ci == 0) and len(text) <= 20:
                    if any(p.find('a:pPr', NS) is None or p.find('a:pPr', NS).get('algn', 'l') != 'ctr'
                           for p in paras if p.find('.//a:t', NS) is not None):
                        fail('TABLE-ALIGN-02', f'Short header/label cell {ri+1},{ci+1} must be horizontally centered', sid)
                sizes = [int(p.get('sz'))/100 for p in cell.findall('.//a:rPr', NS) + cell.findall('.//a:defRPr', NS) if p.get('sz')]
                if not sizes:
                    short = False
                else:
                    largest = max(largest, max(sizes))
                if len(text) > 35 or len([p for p in paras if ''.join(t.text or '' for t in p.findall('.//a:t', NS)).strip()]) > 1:
                    short = False
            if ri > 0 and largest:
                measured += 1
                if short and height > max(32, 2.2*largest):
                    roomy += 1
        ext = f.find('p:xfrm/a:ext', NS)
        area = int(ext.get('cx'))*int(ext.get('cy'))/(12192000*6858000) if ext is not None else 0
        if measured and roomy >= max(2, measured/2) and area > 0.22:
            reason = item.get('table_purpose', {}).get(sid, {}).get('size_exception', '')
            if not str(reason).strip():
                fail('TABLE-SIZE-01', 'Large table with mostly short text and oversized rows; shrink it or record a size_exception in plan.json', sid)
    return issues
