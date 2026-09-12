"""Render candidate.pptx to renders/page-NNN.png + manifest.json without PowerPoint.

Uses LibreOffice (soffice) -> PDF, then PyMuPDF or pdftoppm -> PNG.
Same manifest format as render.ps1 so the harness accepts either engine.
Fonts may be substituted by LibreOffice; prefer render.ps1 on Windows for the final check.
"""
import argparse, hashlib, json, pathlib, shutil, subprocess, sys, tempfile

def sha(p):
    with open(p, 'rb') as f:
        return hashlib.file_digest(f, 'sha256').hexdigest()

def find_soffice():
    for name in ('soffice', 'libreoffice'):
        p = shutil.which(name)
        if p: return p
    for c in ('/Applications/LibreOffice.app/Contents/MacOS/soffice',
              r'C:\Program Files\LibreOffice\program\soffice.exe',
              r'C:\Program Files (x86)\LibreOffice\program\soffice.exe'):
        if pathlib.Path(c).exists(): return c
    raise SystemExit('LibreOffice (soffice) not found. Install it or use render.ps1 on Windows.')

def pdf_to_png(pdf, out, dpi):
    try:
        import fitz  # PyMuPDF
        doc = fitz.open(pdf); paths = []
        for i, page in enumerate(doc, 1):
            dest = out / f'page-{i:03d}.png'
            page.get_pixmap(dpi=dpi).save(dest); paths.append(dest)
        return paths
    except ImportError:
        pass
    if shutil.which('pdftoppm'):
        subprocess.run(['pdftoppm', '-png', '-r', str(dpi), str(pdf), str(out / 'page')], check=True)
        paths = sorted(out.glob('page-*.png'))
        # pdftoppm pads digits by page count; normalise to page-NNN.png
        fixed = []
        for i, p in enumerate(paths, 1):
            dest = out / f'page-{i:03d}.png'
            if p != dest: p.rename(dest)
            fixed.append(dest)
        return fixed
    raise SystemExit('Need PyMuPDF (pip install pymupdf) or pdftoppm (poppler) to rasterise the PDF.')

def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--run', type=pathlib.Path, required=True)
    ap.add_argument('--dpi', type=int, default=120, help='120 dpi -> 1600x900 for a 13.33x7.5in slide')
    a = ap.parse_args(); run = a.run.resolve(); deck = run / 'candidate.pptx'
    if not deck.exists(): raise SystemExit('candidate.pptx missing in ' + str(run))
    before = sha(deck)
    out = run / 'renders'
    if out.exists(): shutil.rmtree(out)
    out.mkdir()
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run([find_soffice(), '--headless', '--convert-to', 'pdf', '--outdir', tmp, str(deck)],
                       check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        pdf = pathlib.Path(tmp) / 'candidate.pdf'
        if not pdf.exists(): raise SystemExit('LibreOffice did not produce a PDF')
        pngs = pdf_to_png(pdf, out, a.dpi)
    if sha(deck) != before: raise SystemExit('PPTX changed while rendering; render again.')
    pages = [{'page': i, 'render': f'renders/{p.name}', 'sha256': sha(p)} for i, p in enumerate(pngs, 1)]
    (out / 'manifest.json').write_text(json.dumps({'engine': 'LibreOffice', 'deck_sha256': before, 'pages': pages}, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f'rendered {len(pages)} pages to {out}')

if __name__ == '__main__':
    main()
