import json
import pathlib
import sys
import tempfile
import unittest

SKILL_SCRIPTS = pathlib.Path(__file__).resolve().parents[1] / 'skills' / 'cdsappt' / 'scripts'
sys.path.insert(0, str(SKILL_SCRIPTS))

import jeju_template as jt
import lecture_library
import reuse_slide


class UsabilityGuardTests(unittest.TestCase):
    def test_emit_json_writes_utf8_without_shell_pipe(self):
        with tempfile.TemporaryDirectory() as folder:
            output = pathlib.Path(folder) / '검색결과.json'
            lecture_library.emit_json({'title': '관리자 AI 리더십'}, output)
            self.assertEqual(json.loads(output.read_text(encoding='utf-8'))['title'], '관리자 AI 리더십')

    def test_dark_layout_blocks_dark_unfilled_text(self):
        xml = f'''<p:sld xmlns:p="{jt.P}" xmlns:a="{reuse_slide.A}"><p:cSld><p:spTree>
        <p:sp><p:nvSpPr><p:cNvPr id="2" name="body"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr><a:solidFill><a:srgbClr val="222222"/></a:solidFill></a:rPr><a:t>어두운 글자</a:t></a:r></a:p></p:txBody></p:sp>
        </p:spTree></p:cSld></p:sld>'''.encode('utf-8')
        warnings = reuse_slide.contrast_warnings({'ppt/slides/slide1.xml': xml}, 'ppt/slides/slide1.xml', 6)
        self.assertTrue(warnings and warnings[0].startswith('contrast risk:'))

    def test_light_layout_does_not_false_positive_on_light_text(self):
        xml = f'''<p:sld xmlns:p="{jt.P}" xmlns:a="{reuse_slide.A}"><p:cSld><p:spTree>
        <p:sp><p:nvSpPr><p:cNvPr id="2" name="body"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr><a:t>이미지 위 밝은 글자</a:t></a:r></a:p></p:txBody></p:sp>
        </p:spTree></p:cSld></p:sld>'''.encode('utf-8')
        self.assertEqual(reuse_slide.contrast_warnings({'ppt/slides/slide1.xml': xml}, 'ppt/slides/slide1.xml', 3), [])


if __name__ == '__main__':
    unittest.main()
