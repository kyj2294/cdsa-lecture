import pathlib
import sys
import unittest

SKILL_SCRIPTS = pathlib.Path(__file__).resolve().parents[1] / 'skills' / 'cdsappt' / 'scripts'
sys.path.insert(0, str(SKILL_SCRIPTS))
from lecture_library import bm25_rank, search_tokens


class SearchTests(unittest.TestCase):
    def test_korean_tokens_preserve_word_boundaries(self):
        tokens = search_tokens('경영진 AI 리터러시')
        self.assertIn('경영', tokens)
        self.assertIn('ai', tokens)
        self.assertNotIn('진a', tokens)

    def test_title_match_outranks_body_only_match(self):
        records = [
            {'id': 'body', 'title': '일반 교육', 'text': '경영진 AI 리터러시 과정'},
            {'id': 'title', 'title': '경영진 AI 리터러시', 'text': '과정 개요'},
        ]
        ranked = bm25_rank(records, '경영진 AI 리터러시')
        self.assertEqual(ranked[0][1]['id'], 'title')

    def test_unrelated_record_is_excluded(self):
        records = [
            {'id': 'hit', 'title': '데이터 분석 표', 'text': '비교'},
            {'id': 'miss', 'title': '인사 규정', 'text': '복무'},
        ]
        self.assertEqual([item[1]['id'] for item in bm25_rank(records, '데이터 표')], ['hit'])


if __name__ == '__main__':
    unittest.main()
