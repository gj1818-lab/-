#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""27SS 판정 백데이터 시트 생성용 CSV (Drive create_file text/csv → 구글시트 자동 변환)

배경: Drive 커넥터는 셀 쓰기가 불가하므로, 판정값(작지 일매/사양확정/LAB DIP)을
실시간 수식으로 계산하는 소형 시트를 새 파일로 만들어 운영 시트가 IMPORTRANGE로 당겨 쓰게 한다.

핵심 제약(실측):
 - CSV→구글시트 변환 시 그리드가 CSV 크기로 고정됨. '열'은 자동 확장이 안 되므로
   헤더를 71필드(BS열)까지 늘려 소재DB 스필(AF3:BS2000)이 들어갈 자리를 확보해야 함.
   '행'은 배열 스필 시 자동 추가되므로 패딩 불필요.
 - MAP/LAMBDA·ARRAYFORMULA·COUNTIF(S)·MAXIFS 모두 CSV 수식으로 파싱됨.
 - COUNTIFS "?*" 는 IMPORTRANGE 빈 문자열을 배제(과거 "<>" 빈칸 오집계 버그의 수정판).

배포본: https://docs.google.com/spreadsheets/d/1C3YAhWDo3YPMFYHKwk6gWyGyK56ihSTbyx11Nrw2L4A/
레이아웃: A품번 B작지 C사양 D LABDIP(텍스트) E labdip_serial F컬러수 G일매완료 H사양완료
          J2=IMPORTRANGE(EDW_RAW!A2:U1500 → J:AD), AF3=IMPORTRANGE(Merged Data!A3:AN2000 → AF:BS)
          (EDW: A품번=J, R일매=AA, U사양=AD / 소재DB: G품번=AL, C진행구분=AH, AK LABDIP=BP)
          428~437행 검증 카운터
"""
import csv, io, os

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(BASE, 'data', 'source_0808_style_master.csv')
OUT = os.path.join(BASE, 'output', '27SS_판정_백데이터.csv')
EDW = "https://docs.google.com/spreadsheets/d/1iSDzZFUv3855k9Zh-tNZALH8z3gY7dTw9j8RlHpJETg"
SOJE = "https://docs.google.com/spreadsheets/d/1_LNPz9XGRmRAJ9Y4cqtFF8NtZTLSIld4vtzLJlOcka0"

codes = [r['StyleCode'].strip() for r in csv.DictReader(open(SRC, encoding='utf-8')) if r['StyleCode'].strip()]
assert len(codes) == 424, len(codes)
last = len(codes) + 1  # 425

F = {
 'B': f'=ARRAYFORMULA(IF($A$2:$A${last}="","",IF($F$2:$F${last}=0,"",IF($G$2:$G${last}=$F$2:$F${last},"O",IF($G$2:$G${last}>0,"부분","X")))))',
 'C': f'=ARRAYFORMULA(IF($A$2:$A${last}="","",IF($F$2:$F${last}=0,"",IF($H$2:$H${last}=$F$2:$F${last},"O",IF($H$2:$H${last}>0,"부분","X")))))',
 'D': f'=ARRAYFORMULA(IF(($A$2:$A${last}="")+($E$2:$E${last}=0),"",TEXT($E$2:$E${last},"yyyy-mm-dd")))',
 'E': f'=MAP($A$2:$A${last},LAMBDA(x,IF(x="",,MAXIFS($BP$3:$BP$2000,$AL$3:$AL$2000,x,$AH$3:$AH$2000,"<>*drop*"))))',
 'F': f'=MAP($A$2:$A${last},LAMBDA(x,IF(x="",,COUNTIF($J$2:$J$1500,x))))',
 'G': f'=MAP($A$2:$A${last},LAMBDA(x,IF(x="",,COUNTIFS($J$2:$J$1500,x,$AA$2:$AA$1500,"?*"))))',
 'H': f'=MAP($A$2:$A${last},LAMBDA(x,IF(x="",,COUNTIFS($J$2:$J$1500,x,$AD$2:$AD$1500,"?*"))))',
}
buf = io.StringIO()
w = csv.writer(buf, lineterminator='\n')
hdr = ['품번','작지발행(일매)','사양확정','LAB DIP CFM','labdip_serial','컬러수EDW','일매완료','사양완료',
       "★사용법: J2와 AF3 셀에서 액세스 허용 클릭(최초1회)",'EDW_RAW(수정금지)'] + ['']*21 + ['소재DB_RAW(수정금지)'] + ['']*38 + ['그리드경계(수정금지)']
assert len(hdr) == 71
w.writerow(hdr)
w.writerow([codes[0], F['B'], F['C'], F['D'], F['E'], F['F'], F['G'], F['H'], '',
            f'=IMPORTRANGE("{EDW}","EDW_RAW!A2:U1500")'])
w.writerow([codes[1]] + ['']*30 + [f'=IMPORTRANGE("{SOJE}","Merged Data!A3:AN2000")'])
for c in codes[2:]:
    w.writerow([c])
w.writerow([]); w.writerow(['',''])
w.writerow(['검증_품번수', f'=COUNTA($A$2:$A${last})'])
for lbl, col in [('작지','B'), ('사양','C')]:
    w.writerow([f'검증_{lbl}O',  f'=COUNTIF(${col}$2:${col}${last},"O")'])
    w.writerow([f'검증_{lbl}부분', f'=COUNTIF(${col}$2:${col}${last},"부분")'])
    w.writerow([f'검증_{lbl}X',  f'=COUNTIF(${col}$2:${col}${last},"X")'])
    w.writerow([f'검증_{lbl}공란', f'=COUNTIFS($A$2:$A${last},"?*",${col}$2:${col}${last},"")'])
w.writerow(['검증_LABDIP보유', f'=COUNTIF($D$2:$D${last},"?*")'])
open(OUT, 'w', encoding='utf-8').write(buf.getvalue())
print('saved:', OUT, len(buf.getvalue()), 'chars')
