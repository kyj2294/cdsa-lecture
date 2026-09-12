"""Regression checks for layered labels and nested group alignment. No user files needed."""
import copy,unittest,xml.etree.ElementTree as E
import shape_labels as s

def fixture(layered=False):
    # Two nested coordinate systems, both with scaling/translation.
    xml=f'''<p:sld xmlns:p="{s.P}" xmlns:a="{s.A}"><p:cSld><p:spTree>
    <p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="outer"/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="1270000" y="1270000"/><a:ext cx="5080000" cy="2540000"/><a:chOff x="0" y="0"/><a:chExt cx="2540000" cy="1270000"/></a:xfrm></p:grpSpPr>
    <p:grpSp><p:nvGrpSpPr><p:cNvPr id="11" name="inner"/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2540000" cy="1270000"/><a:chOff x="0" y="0"/><a:chExt cx="2540000" cy="1270000"/></a:xfrm></p:grpSpPr>
    <p:sp><p:nvSpPr><p:cNvPr id="12" name="rounded node"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1143000" cy="254000"/></a:xfrm><a:prstGeom prst="roundRect"/><a:solidFill><a:srgbClr val="333333"/></a:solidFill></p:spPr><p:txBody><a:bodyPr anchor="ctr" lIns="63500" rIns="63500" tIns="25400" bIns="25400"><a:noAutofit/></a:bodyPr><a:p><a:pPr algn="ctr"/><a:r><a:t>단계 이름</a:t></a:r></a:p></p:txBody></p:sp>
    </p:grpSp></p:grpSp></p:spTree></p:cSld></p:sld>'''
    e=E.fromstring(xml)
    if layered:
        node=e.find('.//p:sp',s.NS);label=copy.deepcopy(node);label.find('p:nvSpPr/p:cNvPr',s.NS).set('id','13')
        label.find('p:spPr/a:prstGeom',s.NS).set('prst','rect');pr=label.find('p:spPr',s.NS);pr.remove(pr.find('a:solidFill',s.NS));E.SubElement(pr,'{'+s.A+'}noFill')
        xf=label.find('p:spPr/a:xfrm',s.NS);xf.find('a:off',s.NS).set('x','127000');xf.find('a:ext',s.NS).set('cx','889000')
        node.find('.//a:t',s.NS).text='';e.findall('.//p:grpSp',s.NS)[1].append(label)
    return e

class Labels(unittest.TestCase):
    def codes(self,e,exclusions=()):return {v['rule'] for v in s.inspect_slide(e,7,exclusions)}
    def test_valid_center(self):self.assertEqual(self.codes(fixture()),set())
    def test_nested_layered(self):
        r=s.inspect_slide(fixture(True),7);self.assertEqual([v['rule'] for v in r],['LABEL-01']);self.assertEqual(r[0]['group_path'],['10','11']);self.assertEqual(r[0]['label_shape_id'],'13')
    def test_vertical(self):
        e=fixture();e.find('.//a:bodyPr',s.NS).set('anchor','t');self.assertIn('LABEL-02',self.codes(e))
    def test_horizontal(self):
        e=fixture();e.find('.//a:pPr',s.NS).set('algn','r');self.assertIn('LABEL-03',self.codes(e))
    def test_indent(self):
        e=fixture();e.find('.//a:pPr',s.NS).set('indent','12700');self.assertIn('LABEL-03',self.codes(e))
    def test_insets(self):
        e=fixture();e.find('.//a:bodyPr',s.NS).set('lIns','127000');self.assertIn('LABEL-04',self.codes(e))
    def test_resize(self):
        e=fixture();E.SubElement(e.find('.//a:bodyPr',s.NS),'{'+s.A+'}spAutoFit');self.assertIn('LABEL-05',self.codes(e))
    def test_textbox_is_not_node(self):
        e=fixture();pr=e.find('.//p:spPr',s.NS);E.SubElement(pr,'{'+s.A+'}noFill');e.find('.//a:bodyPr',s.NS).set('anchor','t');self.assertEqual(self.codes(e),set())
    def test_missing_exclusion_reason_does_not_bypass(self):self.assertIn('LABEL-01',self.codes(fixture(True),[{'shape_id':'12'}]))
    def test_documented_nonlabel_exclusion(self):self.assertEqual(self.codes(fixture(True),[{'shape_id':'12','reason':'Deliberate annotation background, not a step label'}]),set())

if __name__=='__main__':unittest.main()
