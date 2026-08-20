#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""27SS 판정 백데이터 시트 생성기 (CSV → 구글시트 변환 업로드용)

배경: Drive 커넥터는 셀 쓰기가 불가 → 운영 시트를 직접 못 고침.
우회: 판정값(작지 일매/사양확정/LAB DIP CFM)을 실시간 수식으로 계산하는
      소형 시트를 CSV로 생성 → create_file(text/csv) 업로드 시 구글시트로
      자동 변환되며 '=' 셀은 수식으로 해석됨. 운영 시트는 이 시트를
      IMPORTRANGE+VLOOKUP 한 줄로 참조.

구조(단일 탭):
  A 품번(424) | B 작지발행(일매) | C 사양확정 | D LAB DIP CFM(텍스트)
  E labdip_serial | F 컬러수EDW | G 일매완료 | H 사양완료   ← 모두 배열수식(MAP/ARRAYFORMULA)
  J2  =IMPORTRANGE(EDW 진척현황, EDW_RAW!A2:U1500)   → J:AD 스필 (A→J, R일매→AA, U사양→AD)
  AF3 =IMPORTRANGE(소재DB, Merged Data!A3:AN2000)     → AF:BS 스필 (G품번→AL, C진행구분→AH, AK LABDIP→BP)
  428~437행: 검증 카운터

핵심 수식 결정 사항:
  - 작지/사양 완료수: COUNTIFS(..., "?*") — IMPORTRANGE 빈 문자열 오집계(COUNTIFS "<>") 버그 회피
  - LABDIP: MAXIFS + 품번 와일드카드(x&"*") — 소재DB 품번 뒤 공백(TVJJS27542 ) 매칭 실패 대응,
    진행구분 "<>*drop*"으로 Drop 제외, 텍스트 표기('X')는 MAXIFS가 자동 무시
"""
import csv, io, os

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(BASE, 'data', 'source_0808_style_master.csv')
OUT = os.path.join(BASE, 'output', '27SS_판정_백데이터.csv')

EDW = "https://docs.google.com/spreadsheets/d/1iSDzZFUv3855k9Zh-tNZALH8z3gY7dTw9j8RlHpJETg"
SOJE = "https://docs.google.com/spreadsheets/d/1_LNPz9XGRmRAJ9Y4cqtFF8NtZTLSIld4vtzLJlOcka0"

with open(SRC, newline='', encoding='utf-8') as f:
    codes = [r['StyleCode'].strip() for r in csv.DictReader(f) if r['StyleCode'].strip()]
assert len(codes) == 424, len(codes)
LAST = len(codes) + 1  # 425

A = f"$A$2:$A${LAST}"
F_ = f"$F$2:$F${LAST}"; G_ = f"$G$2:$G${LAST}"; H_ = f"$H$2:$H${LAST}"; E_ = f"$E$2:$E${LAST}"
FORM = {
 'B': f'=ARRAYFORMULA(IF({A}="","",IF({F_}=0,"",IF({G_}={F_},"O",IF({G_}>0,"부분","X")))))',
 'C': f'=ARRAYFORMULA(IF({A}="","",IF({F_}=0,"",IF({H_}={F_},"O",IF({H_}>0,"부분","X")))))',
 'D': f'=ARRAYFORMULA(IF(({A}="")+({E_}=0),"",TEXT({E_},"yyyy-mm-dd")))',
 'E': f'=MAP({A},LAMBDA(x,IF(x="",,MAXIFS($BP$3:$BP$2000,$AL$3:$AL$2000,x&"*",$AH$3:$AH$2000,"<>*drop*"))))',
 'F': f'=MAP({A},LAMBDA(x,IF(x="",,COUNTIF($J$2:$J$1500,x))))',
 'G': f'=MAP({A},LAMBDA(x,IF(x="",,COUNTIFS($J$2:$J$1500,x,$AA$2:$AA$1500,"?*"))))',
 'H': f'=MAP({A},LAMBDA(x,IF(x="",,COUNTIFS($J$2:$J$1500,x,$AD$2:$AD$1500,"?*"))))',
}
IMP_EDW = f'=IMPORTRANGE("{EDW}","EDW_RAW!A2:U1500")'
IMP_SOJE = f'=IMPORTRANGE("{SOJE}","Merged Data!A3:AN2000")'

buf = io.StringIO()
w = csv.writer(buf, lineterminator='\n')
hdr = ["품번","작지발행(일매)","사양확정","LAB DIP CFM","labdip_serial","컬러수EDW","일매완료","사양완료",
       "★사용법: J2와 AF3 셀에서 액세스 허용 클릭(최초1회)","EDW_RAW(수정금지)"] + [""]*21 + ["소재DB_RAW(수정금지)"]
w.writerow(hdr)
w.writerow([codes[0]] + [FORM[c] for c in 'BCDEFGH'] + ["", IMP_EDW])
w.writerow([codes[1]] + [""]*30 + [IMP_SOJE])
for c in codes[2:]:
    w.writerow([c])
w.writerow([]); w.writerow([""])
CNT = [
 ("검증_품번수", f'=COUNTA({A})'),
 ("검증_작지O", f'=COUNTIF($B$2:$B${LAST},"O")'),
 ("검증_작지부분", f'=COUNTIF($B$2:$B${LAST},"부분")'),
 ("검증_작지X", f'=COUNTIF($B$2:$B${LAST},"X")'),
 ("검증_작지공란", f'=COUNTIFS({A},"?*",$B$2:$B${LAST},"")'),
 ("검증_사양O", f'=COUNTIF($C$2:$C${LAST},"O")'),
 ("검증_사양부분", f'=COUNTIF($C$2:$C${LAST},"부분")'),
 ("검증_사양X", f'=COUNTIF($C$2:$C${LAST},"X")'),
 ("검증_사양공란", f'=COUNTIFS({A},"?*",$C$2:$C${LAST},"")'),
 ("검증_LABDIP보유", f'=COUNTIF($D$2:$D${LAST},"?*")'),
]
for label, formula in CNT:
    w.writerow([label, formula])

open(OUT, 'w', encoding='utf-8').write(buf.getvalue())
print("saved:", OUT, len(buf.getvalue()), "chars,", len(codes), "styles")
