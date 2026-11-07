# Paseo Excel 뷰어 플러그인 계획

목표: Excel 파일(.xlsx/.xlsm/.xls)을 Paseo 파일 탭에서 서식 포함 최대한 원본에 가깝게, 빠르게, 데스크톱과 모바일 모두에서 보여준다.

## 1. 플랫폼 제약

| 제약 | 결과 |
|---|---|
| 클라이언트 모듈은 `react`, `react-native`, `zod`, `@tanstack/react-query`, `@getpaseo/plugin/*`만. WebView·SVG·Canvas·DOM 없음 | RN `View`/`Text`로 그리드를 직접 렌더링 |
| `addFilePanel` + 파일 핸들(`readBytes`/`url`/`watch`), Paseo >= 0.13.0 | 파일 탭 안의 뷰어, 파일 변경 시 자동 새로고침 |
| 서버(daemon 하위 프로세스)는 Node API와 npm 의존성 사용 가능. 큰 데이터는 수백 KiB 단위로 나눠 보냄 | 파싱은 서버, 렌더링은 클라이언트 |

## 2. 아키텍처

```
[server]                                         [client (RN)]
.xlsx/.xlsm → unzip(fflate) → XML SAX 파서 ─┐     FilePanel
                                            ├→ WorkbookModel ─RPC→  ├ 시트 탭
.xls → CFB(cfb) → BIFF8 레코드 파서 ────────┘     (공통 중간 모델)     ├ 가상화 그리드 + 틀 고정
                                                                    ├ 병합 셀 = absolute View 하나
캐시: (path, mtime, size) → WorkbookModel (LRU)                      ├ 페이지 나누기 겹침층
                                                                    └ 셀 상세 바(값/수식)
RPC  workbook.open → 메타(시트, 치수, 너비·높이, 병합, 스타일 테이블, 인쇄 설정, 페이지 나누기)
     sheet.tiles   → 셀 타일(예: 200행 × 50열), 화면 근처만 요청
```

- **공통 중간 모델:** 두 파서는 같은 `WorkbookModel`을 만든다. 렌더러, 페이지 계산, 캐시는 파일 형식을 몰라도 된다.
- **수식:** 다시 계산하지 않고 캐시 값을 보여준다(xlsx `<v>`, xls FORMULA 결과 + STRING 레코드). 수식 원문은 상세 바에 표시한다(xls는 P1에서 수식 원문 디코딩).
- **스타일:** 셀에는 스타일 ID만 붙인다. 클라이언트는 ID마다 style 객체를 한 번만 만든다.
- **좌표:** 숨김 행·열은 크기 0으로 둔 누적 합(prefix sum)과 이진 탐색으로 처리한다. 병합 영역은 따로 색인한다.
- **단위:** 열 너비 = 문자 수 × 7px + 5px, 행 높이 = pt × 4/3, 여기에 배율을 곱한다.
- **표시 형식:** `ssf` 패키지 사용.

## 3. 파서

### xlsx / xlsm
workbook(시트, 숨김 상태, 정의된 이름), sharedStrings, styles(numFmt, font, fill, border, cellXfs, dxf), theme(색과 tint), sheetN(cols, rows, mergeCells, sheetViews, pageSetup, pageMargins, rowBreaks/colBreaks, conditionalFormatting)

### xls (BIFF8, Excel 97-2003)
CFB 컨테이너 → Workbook 스트림 레코드:

| 영역 | 레코드 |
|---|---|
| 시트·숨김 | BOUNDSHEET (visible / hidden / veryHidden) |
| 문자열 | SST + CONTINUE (레코드 경계에서 끊긴 문자열 이어 붙이기) |
| 값 | LABELSST, NUMBER, RK, MULRK, BOOLERR, BLANK, MULBLANK, FORMULA + STRING |
| 서식 | FONT, FORMAT, XF, PALETTE(사용자 색상표), STYLE |
| 크기·숨김 | COLINFO(너비, 숨김), ROW(높이, 숨김), DEFCOLWIDTH, DEFAULTROWHEIGHT |
| 병합 | MERGEDCELLS |
| 보기 | WINDOW2(눈금선, 틀 고정, 페이지 나누기 미리 보기), PANE |
| 인쇄 | SETUP, LEFT/RIGHT/TOP/BOTTOMMARGIN, HORIZONTAL/VERTICALPAGEBREAKS, NAME(Print_Area/Print_Titles, ptgArea3d 해석) |
| P1 | CONDFMT/CF, HLINK, 수식 원문 디코딩 |

xls 범위 밖: 암호화 파일(안내 메시지 표시), BIFF5 이하(Excel 95 이전, 값만 시도하는 방안은 P2에서 검토)

## 4. 지원 범위

### P0
- 시트별 보기, 셀 병합, 수식 캐시 값
- **숨김:** 행·열·시트(hidden, veryHidden). 기본은 Excel처럼 감추고, 토글하면 흐리게 표시
- **서식:** 표시 형식, 글꼴(굵게·기울임·밑줄·취소선·크기·색), 단색 채우기, 테두리(두께, 점선·파선 근사), 가로·세로 정렬, 줄바꿈, 열 너비·행 높이, 틀 고정, 시트별 눈금선
- **색:** xlsx 테마 색과 tint, xls 색상표
- **페이지 나누기:** 수동 나누기(파란 실선), 인쇄 영역(밖은 회색), 기본 보기와 미리 보기 전환(처음에는 파일에 저장된 보기 모드를 따름)

### P1
- 자동 페이지 나누기(파란 점선) 계산: 용지 크기, 방향, 여백, 배율·페이지에 맞추기, 인쇄 제목을 반영한 근사값. 경계에서 한 행이나 한 열 정도 어긋날 수 있다는 툴팁 표시
- 조건부 서식(cellIs, colorScale, dataBar), 서식 있는 텍스트, 들여쓰기, 하이퍼링크, 셀 선택과 수식 표시 바, 확대·축소

### 타협하는 부분
- 차트, 도형, 피벗 UI, 스파크라인은 제외하고 "차트 n개 생략"으로 표시
- 패턴·그라디언트 채우기는 대표색 하나로 근사, 회전 텍스트는 90°만, 글꼴은 시스템 글꼴로 대체
- 이미지와 메모는 P2
- 수식 재계산 없음. 캐시 값이 없으면 수식 원문을 회색으로 표시

## 5. 모바일 (`layout.compact`)
- 시트 탭은 하단에 가로 스크롤
- 직접 배율 버튼(50~200%)
- 틀 고정 영역의 최대 크기 제한
- 셀을 탭하면 하단 시트에 값·수식·서식 표시
- UI 색은 `theme.colors`에서 가져오고, 셀 영역은 Excel 원래 색 유지

## 6. 작업 단계

| 단계 | 내용 | 완료 기준 |
|---|---|---|
| 0. 스파이크 | ① 서버에서 `params.path`를 절대 경로로 바꿀 수 있는지(안 되면 `readBytes`를 나눠 보냄) ② RPC 한 번의 실제 크기 한도 ③ RN 2방향 스크롤과 가상화 성능(Android 중첩 스크롤) ④ 클라이언트 번들에 순수 JS 패키지를 넣을 수 있는지 | 테스트 파일이 3개 플랫폼에서 열림 |
| 1. xlsx 파서 + 모델 | WorkbookModel 정의, xlsx 스트리밍 파서, 캐시, 타일 API | 10만 행 파일 메타데이터 1초 이내, 기준 파일 테스트 통과 |
| 2. 그리드 렌더러 | 가상화, 병합, 숨김(+토글), 틀 고정, 시트 탭 | 스크롤 중 빈 화면이 거의 없음 |
| 3. xls 파서 | BIFF8 → 같은 WorkbookModel | 같은 내용의 xlsx와 xls 렌더 결과가 일치 |
| 4. 서식 충실도 | 표시 형식, 색, 테두리, 정렬, 줄바꿈 | Excel 스크린샷과 나란히 비교 |
| 5. 페이지 나누기 | 수동 나누기, 인쇄 영역, 보기 전환 → 자동 나누기 계산 | Excel PDF 내보내기 페이지 수와 비교 |
| 6. 모바일 다듬기 | compact 레이아웃, 확대·축소, 셀 상세 | 실제 기기에서 확인 |
| 7. P1 기능 | 조건부 서식, 서식 있는 텍스트, 링크, xls 수식 원문 | |
| 8. 배포 | `OVERVIEW.md`, `requirements.paseo >=0.13.0`, lockfile | |

## 7. 성능 목표
- 5MB·5만 행 파일: 첫 화면 1.5초 이내
- 스크롤 중 타일 요청: 100ms 이내
- 서버 캐시: 최근 파일 N개 LRU, 메모리 상한

## 8. 기준 파일 세트
서식 많은 보고서, 대용량 데이터, 병합·숨김이 많은 양식, 다중 시트(숨김·veryHidden 포함), 인쇄 설정이 된 양식(수동·자동 나누기). 각 파일은 **xlsx와 xls 두 형식으로** 준비한다.

## 9. 진행 상황 (2026-11-07)

- 완료: 0~3단계, 5단계 일부. xlsx/xls 파서, 공통 모델, 캐시와 청크 RPC, RN 그리드 렌더러(가상화, 병합, 숨김 토글, 틀 고정, 머리글), 서식(표시 형식, 테마 색, 테두리, 정렬, 줄바꿈, 회전, 넘침), 이미지, 메모, 하이퍼링크, 페이지 나누기 미리 보기(수동/자동/인쇄 영역/페이지 번호)
- 검증: Excel COM 정답 비교 테스트(`npm test`, 74개 통과). 샘플과 기능 파일 모두 xlsx/xls에서 표시값, 열 너비, 숨김, 수동 페이지 나누기 일치. 자동 페이지 나누기는 1행 이내 차이
- 성능: 5만 행 × 12열(60만 셀) 첫 화면 준비 xlsx 0.9초, xls 0.7초 (daemon 기준)
- 남은 일: 실제 Paseo 화면 확인(데스크톱, 모바일), 조건부 서식, xls 수식 원문, 핀치 줌, 큰 파일 파싱 추가 최적화
