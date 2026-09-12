"""Real-corpus regression for the cdsappt single-master reuse workflow.

Run with a fresh work directory:
    python scripts/test_harness.py --run <empty-or-new-directory>
"""
import argparse
import copy
import pathlib
import sys
import zipfile

ROOT = pathlib.Path(__file__).resolve().parents[1]
SKILL_SCRIPTS = ROOT / 'skills' / 'cdsappt' / 'scripts'
sys.path.insert(0, str(SKILL_SCRIPTS))
sys.path.insert(0, str(ROOT / 'scripts'))

import harness as h
import jeju_template as jt
import reuse_slide


SLIDES = (
    ('s-08846a51f4a191232e31', 3, 'AI 리터러시와 업무 활용', None),
    ('s-7c29bb02319836f63e96', 9, '딥리서치 도구 비교', 'table'),
    ('s-2566dee9c2582601177b', 3, 'AI 업무 분석 프로세스', 'diagram'),
)


def write_deck(path, files):
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as archive:
        for name, data in files.items():
            archive.writestr(name, data)


def require_pass(run, label):
    result = h.check(run, visual=False)
    if result['status'] != 'PASS':
        raise AssertionError(f'{label}: {result}')
    return result


def setup_run(run):
    h.init(run)
    request = h.load(run / 'request.json')
    request.update(title='경영진 AI 리터러시', subtitle='', presenter='CDSA', date='2026-09-11')
    request['answers'] = {
        'topic': {'value': '경영진 AI 리터러시'},
        'audience': {'value': '공공기관 관리자'},
        'duration_minutes': {'value': '90'},
        'institution': {'value': 'CDSA'},
        'materials': {'value': '내장 강의 라이브러리'},
        'cover_mood': {'value': '표지 없음'},
    }
    request['logo'] = {'mode': 'none', 'file': '', 'source': 'none', 'institution': 'CDSA'}
    h.save(run / 'request.json', request)
    h.prepare(run)

    plan = {'slides': []}
    for slide_id, layout, title, required in SLIDES:
        item = {
            'layout': layout,
            'role': 'content',
            'title': title,
            'purpose': title + '을 설명한다',
            'visual_query': title,
            'sources': [{'kind': 'internal', 'slide_id': slide_id}],
        }
        if required:
            item['required_native'] = [required]
        plan['slides'].append(item)
    h.save(run / 'plan.json', plan)
    plan_issues = h.inspect_plan(plan)
    if plan_issues:
        raise AssertionError(plan_issues)
    retrieval = h.plan_check(run)
    if retrieval['status'] != 'PASS':
        raise AssertionError(retrieval)

    jt.create(jt.read(run / 'template.pptx'), [item[1] for item in SLIDES], run / 'candidate.pptx')
    for page, (slide_id, _, _, _) in enumerate(SLIDES, 1):
        reuse_slide.reuse(run, page, slide_id, all_body_shapes=True)


def assert_single_master_and_connections(run):
    files = jt.read(run / 'candidate.pptx')
    presentation = jt.xml(files['ppt/presentation.xml'])
    masters = presentation.findall('p:sldMasterIdLst/p:sldMasterId', h.NS)
    if len(masters) != 1:
        raise AssertionError(f'expected one master, found {len(masters)}')
    for part in reuse_slide.ordered_slides(files):
        root = jt.xml(files[part])
        ids = {node.get('id') for node in root.findall('.//p:cNvPr', h.NS)}
        for node in root.findall('.//a:stCxn', h.NS) + root.findall('.//a:endCxn', h.NS):
            if node.get('id') not in ids:
                raise AssertionError(f'{part}: dangling connector target {node.get("id")}')
    errors = jt.verify(files, jt.read(run / 'template.pptx'))
    if errors:
        raise AssertionError(errors)


def assert_reused_shapes_remain_editable(run):
    files = jt.read(run / 'candidate.pptx')
    part = reuse_slide.ordered_slides(files)[0]
    root = jt.xml(files[part])
    shape = next(node for node in root.find('p:cSld/p:spTree', h.NS)
                 if reuse_slide.xfrm_of(node) is not None)
    offset = reuse_slide.xfrm_of(shape).find('a:off', h.NS)
    offset.set('x', str(int(offset.get('x')) + 1000))
    files[part] = jt.serialize(root)
    write_deck(run / 'candidate.pptx', files)
    require_pass(run, 'ordinary edit after reuse')


def assert_missing_recorded_shape_is_rejected(run):
    plan = h.load(run / 'plan.json')
    altered = copy.deepcopy(plan)
    altered['slides'][0]['reuse']['shapes'].append('999999')
    h.save(run / 'plan.json', altered)
    result = h.check(run, visual=False)
    h.save(run / 'plan.json', plan)
    if result['status'] == 'PASS' or not any(issue['rule'] == 'REUSE-01' for issue in result['issues']):
        raise AssertionError(f'missing recorded shape was not rejected: {result}')


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--run', type=pathlib.Path, required=True)
    args = parser.parse_args()
    run = args.run.resolve()
    setup_run(run)
    require_pass(run, 'real corpus reuse')
    assert_single_master_and_connections(run)
    assert_reused_shapes_remain_editable(run)
    assert_missing_recorded_shape_is_rejected(run)
    require_pass(run, 'restored plan')
    print('PASS: single master, real text/table/diagram reuse, connector integrity, and post-reuse editing')


if __name__ == '__main__':
    main()
