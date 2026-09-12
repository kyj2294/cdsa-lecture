import unittest, xml.etree.ElementTree as E
import teaching_visuals as t
P = t.NS['p']; A = t.NS['a']

def table_fixture(height=55, align=True):
    e = E.fromstring(f'<p:sld xmlns:p="{P}" xmlns:a="{A}"><p:cSld><p:spTree/></p:cSld></p:sld>')
    tree = e.find('p:cSld/p:spTree', t.NS)
    f = E.SubElement(tree, '{'+P+'}graphicFrame'); nv = E.SubElement(f, '{'+P+'}nvGraphicFramePr')
    E.SubElement(nv, '{'+P+'}cNvPr', id='200', name='test table')
    xf = E.SubElement(f, '{'+P+'}xfrm')
    E.SubElement(xf, '{'+A+'}off', x=str(50*12700), y=str(170*12700))
    E.SubElement(xf, '{'+A+'}ext', cx=str(800*12700), cy=str(height*3*12700))
    tbl = E.SubElement(f, '{'+A+'}tbl')
    for r in range(3):
        row = E.SubElement(tbl, '{'+A+'}tr', h=str(height*12700))
        for col in range(2):
            c = E.SubElement(row, '{'+A+'}tc'); tx = E.SubElement(c, '{'+A+'}txBody'); p = E.SubElement(tx, '{'+A+'}p')
            E.SubElement(p, '{'+A+'}pPr', algn='ctr' if align else 'l')
            run = E.SubElement(p, '{'+A+'}r'); E.SubElement(run, '{'+A+'}rPr', sz='1600'); E.SubElement(run, '{'+A+'}t').text = '짧은 내용'
            E.SubElement(c, '{'+A+'}tcPr', **({'anchor': 'ctr'} if align else {}))
    return e

class TableGate(unittest.TestCase):
    def codes(self, e, item=None):
        return {x['rule'] for x in t.inspect_slide(e, 1, item or {})}
    def test_roomy(self):
        self.assertIn('TABLE-SIZE-01', self.codes(table_fixture()))
    def test_compact(self):
        self.assertEqual(self.codes(table_fixture(28)), set())
    def test_alignment_defaults(self):
        self.assertTrue({'TABLE-ALIGN-01', 'TABLE-ALIGN-02'} <= self.codes(table_fixture(28, False)))
    def test_size_exception(self):
        self.assertEqual(self.codes(table_fixture(), {'table_purpose': {'200': {'size_exception': 'annotated cells need room'}}}), set())

if __name__ == '__main__':
    unittest.main()
