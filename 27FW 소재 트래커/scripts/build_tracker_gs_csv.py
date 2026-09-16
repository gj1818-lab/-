# 27FW 소재역산 트래커 — 구글시트 운영판 생성기 (CSV -> Google Sheet 자동 변환 업로드용)
# 단일 탭: A~AG 트래커(425 스타일, xlsx판과 동일 열) + AI1~AX1 기준정보(수정 가능 셀)
# 수식은 2행 앵커의 ARRAYFORMULA 스필 (행별 수식 없음 → CSV 소형화, 27SS 백데이터 방식)
# 주의: CSV 변환기는 열린 범위(A2:A)를 #NAME? 처리 → 전부 닫힌 범위 / MIN·MAX는 배열 브로드캐스트
#       안 되므로 IF로 원소별 구현 / 날짜 연산은 TO_DATE로 감쌈
import csv, io

SRC = "data/style_master_27fw.csv"
OUT = "output/27FW_소재역산_트래커_GS.csv"

import re
def week_key(w):
    m=re.match(r'(\d+)월(\d+)주',w or '')
    return (int(m.group(1)),int(m.group(2))) if m else (99,99)
def sort_key(r):
    if r['시즌코드']=='A': grp=0
    elif r.get('비수기','')=='O': grp=1
    else: grp=2
    so={'W':0,'X':1}.get(r['시즌코드'],2) if grp==2 else 0
    return (grp,so,week_key(r['출고일(주차)']),r['StyleCode'])
def money(v):
    v=(v or '').replace(',','').strip()
    return int(v) if v.isdigit() else 0

rows=list(csv.DictReader(open(SRC)))
rows.sort(key=sort_key)
N=len(rows); FIRST=2; LAST=FIRST+N-1  # 데이터 2~426행

# 기준정보 셀 (1행 우측): AJ1 기준일 / AL1 입고리드 / AN1 소재리드 / AP1 연도
#   AR1 띤쨘시작 / AT1 띤쨘끝 / AV1 르바란시작 / AX1 르바란끝
AJ,AL,AN,AP,AR,AT,AV,AX='$AJ$1','$AL$1','$AN$1','$AP$1','$AR$1','$AT$1','$AV$1','$AX$1'
def rng(col): return f'${col}${FIRST}:${col}${LAST}'
D,I,U,H,J,K,L,M,P,R,AD=[rng(c) for c in ['D','I','U','H','J','K','L','M','P','R','AD']]

F_G=(f'=ARRAYFORMULA(IF({D}="","",IF(MID({D},5,1)="A","A(가을)",IF(MID({D},5,1)="W","W(겨울)",'
     f'IF(MID({D},5,1)="X","X(사계절)",MID({D},5,1))))))')
F_J=(f'=ARRAYFORMULA(IF({I}="","",IFERROR(DATE({AP},VALUE(LEFT({I},FIND("월",{I})-1)),'
     f'1+7*(VALUE(SUBSTITUTE(MID({I},FIND("월",{I})+1,5),"주",""))-1)),"")))')
F_K=(f'=ARRAYFORMULA(IF({J}="","",TO_DATE('
     f'IF(ISNUMBER(SEARCH("미얀마",{U}))*({J}-{AL}>={AR})*({J}-{AL}<={AT})=1,{AR}-1,'
     f'IF(ISNUMBER(SEARCH("인도네시아",{U}))*({J}-{AL}>={AV})*({J}-{AL}<={AX})=1,{AV}-1,{J}-{AL})))))')
def ov(start,end):
    hi=f'IF({K}<{end},{K},{end})'; lo=f'IF({K}-{AN}>{start},{K}-{AN},{start})'
    return f'IF({hi}-{lo}+1>0,{hi}-{lo}+1,0)'
F_L=(f'=ARRAYFORMULA(IF({K}="","",IF(ISNUMBER(SEARCH("미얀마",{U})),{ov(AR,AT)},'
     f'IF(ISNUMBER(SEARCH("인도네시아",{U})),{ov(AV,AX)},0))))')
F_M=f'=ARRAYFORMULA(IF({K}="","",TO_DATE({K}-{AN}-{L})))'
F_N=f'=ARRAYFORMULA(IF({M}="","",{M}-{AJ}))'
F_S=f'=ARRAYFORMULA(IF({H}="","",IF({H}="C","본사(소재팀)",IF({H}="완","생산처","확인 필요"))))'
F_AD=(f'=ARRAYFORMULA(IF({D}="","",IF({J}="","출고일 미정",IF({P}="X","작지 미발행",IF({P}="부분","작지 일부 미발행",'
      f'IF(({R}<>"")*({R}>{M})=1,"소재입고 마감("&TEXT({M},"MM/DD")&") 초과",'
      f'IF({P}="","발행여부 확인 필요",IF({R}="","소재 입고일정 확인 필요","정상"))))))))')
F_AE=(f'=ARRAYFORMULA(IF({D}="","",IF({J}="","",IF((LEFT({AD},2)="작지")+(LEFT({AD},4)="소재입고")>0,'
      f'IF(MID({D},5,1)="A","★ 납기위험(A코드)","◆ 지연주의"),IF({AD}="정상","✓ 정상","? 확인필요")))))')

NCOL=50  # A~AX
hdr=[""]*NCOL
labels=["NO","MD","DS","Style Code","품명","컬러 구성","시즌","생산형태(수/C/완)","출고일(주차)","출고예정일",
        "제품입고 시점","공장휴무 보정(일)","소재입고 마감일","마감까지 D-day","비수기 선정","작지발행(일매)","사양확정",
        "소재입고 예정/실제일","소재입고 책임","생산처","원산지","원단처(대기)","겉감정보(대기)","안감정보(대기)","LAB DIP CFM(대기)",
        "기획수량","원가합(원)","소매가합(원)","입고예정일(생진테)","리스크 사유","판정","조율상태","메모/조치사항"]
hdr[:len(labels)]=labels
cfg=[("AI","기준일→"),("AJ","=TODAY()"),("AK","입고리드→"),("AL",14),("AM","소재리드→"),("AN",100),
     ("AO","연도→"),("AP",2027),("AQ","띤쨘시작→"),("AR","=DATE(2027,4,9)"),("AS","띤쨘끝→"),("AT","=DATE(2027,4,18)"),
     ("AU","르바란시작→"),("AV","=DATE(2027,3,4)"),("AW","르바란끝→"),("AX","=DATE(2027,3,17)")]
colnum=lambda cl: (ord(cl[0])-64)*26+(ord(cl[1])-64) if len(cl)==2 else ord(cl)-64
for cl,v in cfg: hdr[colnum(cl)-1]=v

buf=io.StringIO(); w=csv.writer(buf, lineterminator='\n')
w.writerow(hdr)
ANCHOR={'G':F_G,'J':F_J,'K':F_K,'L':F_L,'M':F_M,'N':F_N,'S':F_S,'AD':F_AD,'AE':F_AE}
for i,r in enumerate(rows):
    qty=money(r['기획수량']); cost=money(r['입고원가']); retail=money(r['예상소매가'])
    row=[""]*NCOL
    row[0]=i+1; row[1]=r['MD']; row[2]=r['DS']; row[3]=r['StyleCode']; row[4]=r['품명']; row[5]=r['컬러구성']
    row[7]=r['수/C/완']; row[8]=r['출고일(주차)']
    row[19]=r['생산처']; row[20]=r['원산지']
    row[25]=qty or ""; row[26]=qty*cost or ""; row[27]=qty*retail or ""; row[28]=r['입고예정일']
    if i==0:
        for cl,f in ANCHOR.items(): row[colnum(cl)-1]=f
    w.writerow(row)
open(OUT,'w',encoding='utf-8').write(buf.getvalue())
print("saved:",OUT,len(buf.getvalue()),"chars,",N,"styles, rows",FIRST,"~",LAST)
