# 작업 폴더와 상태 파일

`harness.py`는 편집을 대신하지 않는다. 요청·출처·렌더·시각 검토가 같은 결과물에 묶였는지 확인한다.

## 순서

```powershell
python <plugin>/scripts/harness.py init --run <run>
python <plugin>/scripts/harness.py prepare --run <run>
python <plugin>/scripts/harness.py plan-check --run <run>
python <plugin>/scripts/harness.py check --run <run> --structural-only
powershell -ExecutionPolicy Bypass -File <plugin>/scripts/render.ps1 -Run <run>
python <plugin>/scripts/harness.py review --run <run>
python <plugin>/scripts/harness.py check --run <run>
python <plugin>/scripts/harness.py release --run <run>
```

## 주요 파일

|파일|역할|
|---|---|
|`request.json`|사용자 요구, 기관, 표지·로고 선택|
|`plan.json`|장표별 목적, 레이아웃, 검색어, 출처, 재사용 개체|
|`assets.json`|표지·로고·재사용 이미지의 파일·SHA-256·출처|
|`template.pptx`|기관명과 로고가 적용된 CDSA 단일 마스터 정본|
|`candidate.pptx`|편집 중인 결과물|
|`retrieved-visuals/index.json`|실제 코퍼스에서 회수한 장표 목록|
|`renders/manifest.json`|PowerPoint 렌더와 candidate 해시|
|`visual-review.json`|페이지별 시각 검토 결과와 해시|
|`release/lecture.pptx`|최종 납품 파일|

`init`은 기존 작업 파일이 있는 폴더를 덮어쓰지 않는다. `plan-check`는 코퍼스 출처를 실제 DB에서 다시 회수한다. `release`는 구조 검사, 최신 렌더, 모든 페이지의 시각 검토가 통과해야 실행된다.

재사용 기록은 원본 장표 ID, 원본 파일 해시, 옮긴 개체 ID, 맞춤 배율과 위치를 남긴다. 사용자는 이후 개체를 편집할 수 있다. 검사에서는 기록된 개체가 사라졌는지와 출처가 유효한지만 확인한다.
