---
name: cdsappt
description: CDSA 강의용 편집 가능한 PowerPoint를 만든다. 사용자가 색인한 PPT 코퍼스의 텍스트·표·도식·이미지를 검색해 CDSA 단일 마스터의 13개 레이아웃 안에 재사용하며, 표지·기관 로고·구조 검사·실제 PowerPoint 렌더·시각 검토까지 수행한다. CDSA 강의안, CDSA-lecture, cdsappt, PPTX 제작·수정 요청에 사용한다. HTML 슬라이드에는 사용하지 않는다.
---

# CDSA PowerPoint

## 1. 요청과 작업 폴더

사용자가 이미 말한 내용은 다시 묻지 않는다. 주제, 대상, 시간, 기관, 제공 자료, 표지 분위기 중 결과를 바꾸는 정보만 확인한다. 장수는 시간과 밀도로 계산하며 고정 장수를 임의로 강제하지 않는다.

```powershell
python <plugin>/scripts/harness.py init --run <run>
# request.json을 사용자가 말한 내용으로 채운다.
python <plugin>/scripts/harness.py prepare --run <run>
```

기관 로고는 `official`, `user`, `placeholder`, `none` 중 하나를 명시한다. 공식·사용자 로고는 투명 PNG와 출처가 필요하다. 로고를 찾지 못했다고 강의를 중단하지 말고 placeholder 또는 none으로 진행한다. 자세한 기준은 [logo.md](references/logo.md)를 따른다.

## 2. 검색과 계획

새로 그리기 전에 사용자가 색인한 코퍼스를 검색한다. DB가 없으면 `<plugin>/scripts/setup-corpus.ps1`로 사용 권한이 있는 PPT 폴더를 먼저 색인한다. 같은 뜻을 두세 표현으로 검색하고, 내용과 시각 구조를 함께 확인한다.

```powershell
python <skill>/scripts/lecture_library.py search "<query>" --limit 8 --output <run>/search/<name>-text.json
python <skill>/scripts/visual_library.py search "<query>" --kind any --limit 8 --output <run>/search/<name>-visual.json
python <skill>/scripts/visual_library.py slide <slide_id> --output <run>/retrieved-visuals/<slide_id>.pptx
```

`plan.json`의 각 장표에는 `layout` 1~13, `role`, `title`, `purpose`, `visual_query`, `sources`를 기록한다. 표·도식이 필요한 장표는 `required_native`에 `table` 또는 `diagram`을 쓴다. 코퍼스 자료는 `slide_id`로 출처를 남긴다.

검색 점수만으로 장표를 확정하지 않는다. 후보의 `title`, `excerpt`, `page`, 시각 구조를 `purpose`와 대조한다. 핵심 명사가 겹치지 않거나 결론이 다르면 더 구체적인 동의어로 다시 검색한다. 30개를 넘는 top-level 개체, 0.70 미만 축소, 대상 레이아웃과 글자 명암 불일치는 재사용 경고로 취급하고 장표를 나누거나 레이아웃을 바꾼다.

```powershell
python <plugin>/scripts/harness.py plan-check --run <run>
```

검색과 출처 규칙은 [search.md](references/search.md), 레이아웃 선택은 [layouts.md](references/layouts.md)를 본다.

## 3. 단일 마스터로 제작

모든 장표는 준비된 `template.pptx`의 CDSA 마스터 하나만 사용한다.

```powershell
python <skill>/scripts/jeju_template.py --template <run>/template.pptx init --layouts 3,9,3 --output <run>/candidate.pptx
```

원본 장표를 통째로 가져오지 않는다. 필요한 top-level 개체를 원본 스타일이 보이는 상태로 CDSA 레이아웃의 안전 영역에 복사한다. 텍스트·표·도식은 PowerPoint에서 계속 편집 가능해야 한다.

```powershell
python <skill>/scripts/reuse_slide.py --run <run> --page N --source <slide_id> --all-body-shapes
python <skill>/scripts/reuse_slide.py --run <run> --page N --source <slide_id> --include-shape <ID_OR_NAME> [--include-shape <ID_OR_NAME> ...]
```

`--all-body-shapes`는 원본 제목·기관명·푸터를 자동 제외한 본문 전체에 쓴다. 남는 옛 과정명이나 기관명은 렌더에서 확인하고 삭제·수정한다. 밝은 글자·어두운 도형은 밝은 레이아웃, 흰 글자 중심 구성은 어두운 레이아웃처럼 실제 대비가 맞는 레이아웃을 고른다. 재사용 뒤 개체 위치·문구·크기를 정상 편집할 수 있으며 검사는 원본과 픽셀 일치를 요구하지 않는다.

맞는 자료가 없을 때만 새 개체를 만든다. 구성은 [compositions.md](references/compositions.md)를 참고한다.

## 4. 표지와 로고

표지가 필요하면 16:9 배경 이미지를 만들거나 사용자가 준 PNG/JPEG를 넣고, 제목·부제·발표자 정보는 편집 가능한 PowerPoint 텍스트로 둔다.

```powershell
python <skill>/scripts/cover_background.py --run <run> --image <png-or-jpeg> --source "<source>"
python <skill>/scripts/cover_layers.py --input <run>/candidate.pptx --output <run>/candidate-with-cover.pptx --course "<course>" --title "<title>" --subtitle "<subtitle>"
```

배경의 인물·공간·텍스트 안전 영역은 [cover.md](references/cover.md)를 따른다. 이미지 생성이 필요하면 `<!-- omd:gen-image -->` 같은 사양 블록을 만들지 말고 현재 환경의 이미지 생성 도구를 직접 사용한다.

## 5. 구조 검사와 시각 검토

```powershell
python <plugin>/scripts/harness.py check --run <run> --structural-only
powershell -ExecutionPolicy Bypass -File <plugin>/scripts/render.ps1 -Run <run>
python <plugin>/scripts/harness.py review --run <run>
python <plugin>/scripts/harness.py check --run <run>
```

렌더된 모든 페이지를 실제로 본다. 다음을 고친 뒤 다시 렌더한다.

- 잘림, 겹침, 읽기 어려운 대비, 너무 작은 글자
- CDSA 상단 띠·강의명·로고와 본문 충돌
- 표 머리글·행 라벨 정렬과 열 폭
- 연결선이 끊긴 도식, 원본 과정명·기관명·날짜 잔존
- 슬라이드 목적과 맞지 않는 장식 또는 근거 없는 문장

자동 규칙의 의미는 [quality.md](references/quality.md), 실행 상태 파일은 [harness.md](references/harness.md)를 본다.

## 6. 납품

구조 검사, 최신 렌더, 전체 시각 검토가 모두 통과한 뒤 릴리스한다.

```powershell
python <plugin>/scripts/harness.py release --run <run>
```

최종 파일은 `<run>/release/lecture.pptx`이다. 답변에는 파일 링크, 장표 수, 사용한 코퍼스 출처, 로고·표지 출처, 실행한 검사와 실제 PowerPoint 렌더 결과를 적는다.

질문 UI는 [question-ui.md](references/question-ui.md)를 참고한다.
