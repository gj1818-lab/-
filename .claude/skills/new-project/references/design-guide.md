# PPTX Slide Deck Design Guide

원본 PPTX의 슬라이드 문구와 수치는 제외하고, 재사용 가능한 디자인 요소만 추출한 가이드입니다. 새 슬라이드 데크 제작 시 아래 토큰과 레이아웃 규칙을 우선 적용합니다.

## 1\. Design Direction

- 스타일: 데이터 중심의 경영/영업 전략 보고서. 장식보다 정보 밀도, 구조, 수치 비교를 우선한다.  
- 인상: 차분한 녹색 기반, 흰 배경, 얇은 컬러 스트립, 작고 촘촘한 표/카드 구성.  
- 구성 원칙: 한 장에 핵심 메시지 1개, 그 아래 KPI 또는 근거 데이터, 마지막에 표/차트/실행 관점 패널을 배치한다.  
- 레이아웃 성격: 대부분 `Blank` 슬라이드에 도형을 직접 배치한 구조다. 마스터 플레이스홀더보다 수동 그리드와 반복 컴포넌트가 중요하다.

## 2\. Canvas

- 슬라이드 비율: 16:9  
- 실제 크기: 13.319 in x 7.5 in  
- 기준 좌우 여백: 0.5 in  
- 콘텐츠 기준 폭: 12.33 in  
- 주요 x 좌표: 0.5, 0.72, 3.05, 3.65, 5.6, 6.8, 8.15, 9.95, 10.7, 12.2  
- 주요 y 좌표: 0, 0.07, 0.62, 1.25, 1.70, 3.05, 4.85, 5.05, 6.75

## 3\. Color Tokens

| Token | Hex | Usage |
| :---- | ----: | :---- |
| `brand-green` | `#2C5F2D` | 상단 헤더, 제목, 핵심 긍정/기준 컬러, 표 헤더 |
| `positive-green` | `#27AE60` | 상승/성장/긍정 KPI 강조 |
| `light-green` | `#97BC62` | 표 합계/마무리 밴드/보조 포인트 |
| `deep-navy` | `#1B3A4B` | 중립 핵심 KPI, 짙은 비교축 |
| `info-blue` | `#2E86AB` | 정보성 KPI, 추정/기회 영역 |
| `alert-red` | `#E8474C` | 감소, 리스크, 주의 항목 |
| `opportunity-orange` | `#F18F01` | 기회, 경고, 액션 강조 |
| `paper-gray` | `#F5F5F5` | 콜아웃 박스, 카드 테두리, 표 교차 행 |
| `table-total-green` | `#E8F5E9` | 표 합계/총계 행 |
| `text-primary` | `#212121` | 일반 본문/표 본문 |
| `text-secondary` | `#666666` | 보조 설명, 단위, 캡션 |
| `white` | `#FFFFFF` | 배경, 카드, 헤더 텍스트 |

컬러 사용 비율은 `white`와 `brand-green`을 압도적으로 많이 쓰고, 빨강/주황/파랑은 KPI별 식별색으로 제한한다.

## 4\. Typography

지정 폰트는 전 영역`Pretendard`이다. 제작 환경에 해당 폰트가 없으면 `맑은 고딕`, `Apple SD Gothic Neo`, `Noto Sans KR`,  순으로 대체한다.

| Role | Size | Weight | Color |
| :---- | ----: | :---- | :---- |
| Cover title | 56 pt | Bold | `white` |
| Slide title | 26 pt | Bold | `brand-green` |
| Subtitle | 15 pt | Regular | `text-secondary` |
| Main message in gray callout | 22 pt | Bold | `brand-green` |
| KPI label | 15 pt | Regular | `text-secondary` |
| KPI number | 32 pt | Bold | KPI accent color |
| KPI note | 13 pt | Regular | `text-secondary` |
| Section/table caption | 14 pt | Bold or Regular | `brand-green` |
| Body bullet | 13-14 pt | Regular | `text-primary` |
| Table cell | 12 pt | Regular/Bold for key values | `text-primary` or `brand-green` |
| Header deck label/page number | 12 pt | Regular | `white` |

텍스트는 대부분 좌측 정렬이고, KPI 숫자와 카드 내부 일부 텍스트는 세로 중앙 정렬을 사용한다. 제목은 크지만 과도한 장식 없이 얇은 여백으로 정보 영역과 연결한다.

## 5\. Core Components

### Top Header Bar

- 위치/크기: x 0, y 0, w 13.33 in, h 0.42 in  
- 배경: `brand-green`  
- 좌측: **대분류 제목**(섹션/챕터 대분류), x 0.45, y 0.07, 12 pt Bold, `white`  
- 우측: `CONFIDENTIAL` 표기, x 12.2, y 0.07, 11 pt Bold, 대문자, 우측 정렬, `alert-red`  
  - 읽기 쉽도록 권장: `CONFIDENTIAL` 텍스트 뒤에 작은 `white` 배경의 라운드 배지(좌우 패딩 약 0.08 in)를 깔고 그 위에 `alert-red` 텍스트를 올린다. 녹색 헤더 위에 빨강 텍스트만 두면 대비가 약하므로 흰 배지 처리를 우선 적용한다.  
- 페이지 번호는 더 이상 헤더에 두지 않고 하단 `Page Footer Bar` 중앙으로 이동한다.  
- 헤더는 모든 본문 슬라이드에 반복하며, 좌측 대분류 제목은 섹션이 바뀔 때만 변경한다.

### Slide Title Block

- 제목: x 0.5, y 0.62, w 12.33 in, h 약 0.46 in  
- 부제: x 0.5, y 1.25, w 12.33 in, h 약 0.27 in  
- 제목은 `brand-green` 26 pt Bold, 부제는 `text-secondary` 15 pt.

### Gray Message Callout

- 박스: x 0.5, y 1.70, w 12.33 in, h 0.95 in, fill `paper-gray`  
- 좌측 스트라이프: x 0.5, y 1.70, w 0.16 in, h 0.95 in, fill `brand-green`  
- 텍스트: x 1.0, y 약 1.98, w 11.7 in, 22 pt Bold, `brand-green`  
- 용도: 슬라이드의 핵심 판단, 전환 신호, 전략 방향을 한 줄로 강조한다.

### KPI Cards

기본 카드는 흰색 라운드 사각형과 얇은 상단 컬러 스트립으로 구성한다.

- 배경: `white`  
- 테두리: 0.5 pt, `paper-gray`  
- 모서리: 작게 둥근 라운드 사각형  
- 상단 스트립: h 0.12 in, 카드 폭 전체, KPI accent color  
- 내부 좌우 패딩: 약 0.22 in  
- 라벨: 15 pt, `text-secondary`  
- 숫자: 32 pt Bold, 상단 스트립과 같은 accent color  
- 보조 설명: 13 pt, `text-secondary`

권장 크기:

- 5열 KPI: card w 2.4 in, h 1.85 in, x 0.5 / 3.05 / 5.6 / 8.15 / 10.7, gap 0.15 in  
- 4열 KPI: card w 3.0 in, h 1.55 in, x 0.5 / 3.65 / 6.8 / 9.95, gap 0.15 in

### Tables

- 표 제목/캡션: 표 위 x 0.5, 14 pt, `brand-green`  
- 헤더 행: fill `brand-green`, text `white`, 12 pt Bold, 중앙 정렬  
- 본문 행: `white`와 `paper-gray` 교차  
- 합계/총계 행: fill `table-total-green`  
- 핵심 수치: `brand-green` Bold  
- 셀 구분선: 흰색 또는 매우 연한 회색, 두껍게 쓰지 않는다.  
- 표는 슬라이드 하단 전체 폭 또는 좌측 6.0 in 내외 영역에 배치한다.

### Charts

- 차트 유형: 세로 막대/누적 막대 중심  
- 컬러 순서: `deep-navy`, `brand-green`, `opportunity-orange`, `alert-red`, `info-blue`  
- 축/그리드: 밝은 회색, 최소화  
- 차트 제목: 13-15 pt, `text-secondary` 또는 `brand-green`  
- 한 슬라이드에 여러 차트를 넣을 때는 3개 이하, 동일 크기와 동일 기준선으로 정렬한다.

### Insight Panel

- 배경: `white`, 테두리 0.5 pt `paper-gray`, 작은 라운드  
- 상단 스트립: `brand-green` 또는 KPI accent color, h 0.10-0.12 in  
- 제목: 15 pt Bold, `brand-green`  
- 본문: 13-14 pt, `text-primary`, bullet 간격 좁게  
- 배치: 하단 우측 5.7-6.1 in 폭, 또는 2열 패널로 분할

### Footer Summary Band

- 위치: y 약 6.75, w 12.33 in  
- fill: `light-green`  
- 텍스트: 18-22 pt Bold, 중앙 정렬, `brand-green` 또는 `white`  
- 용도: 결론/운영 체크포인트/다음 액션을 요약한다.  
- 주의: 아래 `Page Footer Bar`와 다른 요소다. 이 밴드는 "내용 결론"이고, `Page Footer Bar`는 모든 슬라이드 맨 아래의 "고정 식별 바닥글"이다. 두 요소가 한 슬라이드에 함께 있을 경우, summary band를 위(y 약 6.75)에, page footer를 맨 아래(y 약 7.15)에 둔다.

### Page Footer Bar

모든 본문 슬라이드 맨 아래에 반복되는 고정 식별 바닥글이다.

- 위치/크기: x 0.5, y 7.15, w 12.33 in, h 약 0.3 in  
- 배경: 없음(흰 배경 유지). 상단에 0.5 pt `paper-gray` 가는 구분선 1개만 둔다.  
- 좌측: `KS Supply Chain Management`, x 0.5, 좌측 정렬, 10 pt, `text-secondary`  
- 중앙: 페이지 번호, 수평 중앙(x 약 6.66 기준), 10 pt, `text-secondary`  
- 우측: `KOLON SPORT`, x 12.2, 우측 정렬, 10 pt Bold, `brand-green`  
- 좌·중앙·우 세 항목은 같은 y(baseline)에 정렬한다.  
- 바닥글은 표지(Cover)에는 넣지 않고 본문 슬라이드에만 반복한다.  
- 하단 표/패널이 y 7.15에 닿지 않도록, 콘텐츠 영역은 y 6.85 이전에서 끝낸다(바닥글과 최소 0.3 in 간격 확보).

## 6\. Layout Recipes

### A. Cover Slide

- 전체 배경: `brand-green`  
- 중앙 제목: 56 pt Bold, `white`, 수평 중앙  
- 얇은 가로 라인: `light-green`, 제목 위쪽에 짧게 배치  
- 하단 중앙: 작은 보조 텍스트 14 pt, `white`  
- 표/카드/헤더바 없이 매우 단순하게 유지한다.

### B. Executive Summary

1. 상단 헤더바  
2. 제목 \+ 부제  
3. 회색 메시지 콜아웃  
4. 5열 KPI 카드  
5. 하단 전체 폭 표

이 레이아웃은 “핵심 판단 1개 \+ 수치 근거 여러 개 \+ 상세 테이블” 구조에 적합하다.

### C. Chart Analysis

1. 상단 헤더바  
2. 제목 \+ 짧은 부제  
3. 3열 차트 또는 3개 분석 블록  
4. 각 차트 상단에 컬러 스트립 또는 작은 제목  
5. 하단에는 범례/주석을 작게 배치

차트가 많아도 흰 배경과 균등 간격을 유지한다.

### D. Phase / Timeline Slide

1. 상단 헤더바  
2. 제목 \+ 부제  
3. 회색 메시지 콜아웃  
4. 4열 KPI 카드  
5. 하단 좌측 표, 하단 우측 insight panel

단계별 전략, 시즌 전환, 실행 시점 설명에 적합하다.

### E. Checkpoint / Closing Slide

1. 상단 헤더바  
2. 제목 \+ 부제  
3. 회색 메시지 콜아웃  
4. 2x2 checkpoint 카드 또는 2열 패널  
5. 하단 summary band

컬러 스트립을 카드별로 바꾸면 각 체크포인트의 성격을 빠르게 구분할 수 있다.

## 7\. Spacing Rules

- 좌우 기본 여백은 0.5 in으로 고정한다.  
- 카드 간 간격은 0.15 in을 기준으로 한다.  
- 제목 블록과 메시지 콜아웃 사이에는 약 0.18-0.25 in 여백을 둔다.  
- 메시지 콜아웃과 KPI 카드 사이에는 약 0.40 in 이상 여백을 둔다.  
- 하단 표/패널은 y 5.05 전후에서 시작하고, 슬라이드 하단 0.25 in 이상을 남긴다.  
- 라인과 테두리는 얇게 사용한다. 카드 테두리 0.5 pt, 일반 도형 라인 1 pt.

## 8\. Production Notes

- 모든 슬라이드는 흰 배경 위에 도형을 정렬하는 방식으로 제작한다.  
- 텍스트 박스에는 불필요한 채우기를 넣지 않고, 도형 배경과 텍스트를 분리한다.  
- KPI 색상은 숫자, 카드 상단 스트립, 관련 차트 계열에 반복해서 연결한다.  
- 표와 카드의 모서리는 작게만 둥글게 처리한다. 과한 둥근 카드나 그림자는 쓰지 않는다.  
- 이미지는 사용하지 않는 방향이 원본 스타일에 가깝다. 데이터, 표, 차트, 컬러 스트립만으로 정보 구조를 만든다.  
- 한 슬라이드에 색상을 많이 쓰더라도 각 색은 의미를 가져야 한다. 장식용 색상 추가는 피한다.

