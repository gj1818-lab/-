# 27FW 생진테 추출 시트 생성기 (CSV -> Google Sheet 자동 변환용)
# 원본: (자동화)_KS_2027FW_GREEN Product Line Plan (생진테/리퀘스트 통합)
#   탭 'Total(GR,YL통합)' — Q:StyleNo(품번+컬러13자) ~ AA:생산처명, AG:입고원가~AH:예상소매가,
#   AM:납기1차~AU:라벨(AT=출고일 주차), BE:MD~BF:DS
# 방식: IMPORTRANGE 4개로 필요한 컬럼 블록만 스필 (27SS 백데이터 시트 방식 재사용)
# 주의(27SS 삽질 기록): CSV 변환 시 그리드 열 수가 CSV 헤더 필드 수로 고정된다.
#   스필 최대 열(X=24)보다 넓게 30필드로 패딩. 행은 스필 시 자동 확장이라 패딩 불필요.
import csv

SRC_ID = "1VN6T7hVPzzurd9VDuCgYQTHDUgvRqhDke04BDQCMkdw"
TAB = "Total(GR,YL통합)"
MAXROW = 10000
NCOL = 30

def imp(rng):
    return f'=IMPORTRANGE("https://docs.google.com/spreadsheets/d/{SRC_ID}","{TAB}!{rng}")'

rows = []
# 1행: 블록 라벨 + 검증 카운터
r1 = [""] * NCOL
r1[0] = "[A~K] 원본 Q~AA (StyleNo/브랜드/품종/시즌/년도/컬러/품명/기획수량/수C완/원산지명/생산처명)"
r1[11] = "[L~M] 원본 AG~AH (입고원가/예상소매가)"
r1[13] = "[N~V] 원본 AM~AU (납기1차/MIX/종결여부/입고수량/총수불원가/지수/입고예정일/출고일/라벨)"
r1[22] = "[W~X] 원본 BE~BF (MD/DS)"
r1[25] = "검증_품번행수"
r1[26] = '=COUNTIF($A$3:$A$10001,"?*")'  # 열린 범위(A3:A)는 CSV 변환 시 #NAME? — 닫힌 범위 필수
r1[27] = "검증_출고일수"
r1[28] = '=COUNTIF($U$3:$U$10001,"?*")'
rows.append(r1)
# 2행: IMPORTRANGE 앵커 (스필: A2:K / L2:M / N2:V / W2:X)
r2 = [""] * NCOL
r2[0] = imp(f"Q1:AA{MAXROW}")
r2[11] = imp(f"AG1:AH{MAXROW}")
r2[13] = imp(f"AM1:AU{MAXROW}")
r2[22] = imp(f"BE1:BF{MAXROW}")
rows.append(r2)

with open("output/27FW_생진테_추출.csv", "w", newline="", encoding="utf-8") as f:
    csv.writer(f).writerows(rows)
print("written: output/27FW_생진테_추출.csv", f"({NCOL} cols x {len(rows)} rows)")
