# 27FW 스타일 마스터 추출 + 검증
# 입력: data/saengjinte_extract_raw.csv (27FW_생진테_추출 시트의 CSV 다운로드본, 품번+컬러 행 단위)
# 출력: data/style_master_27fw.csv (스타일 단위 집계) — 원가 포함이라 git 커밋 금지(.gitignore)
import csv, re, sys
from collections import OrderedDict, Counter

RAW = "data/saengjinte_extract_raw.csv"
OUT = "data/style_master_27fw.csv"
STYLE_RE = re.compile(r"^([A-Z0-9]{2}[A-Z]{2}[A-Z]\d{2}\d{3})\s*([A-Z0-9]{0,3})")  # 컬러 0~3자·뒤 잡음 허용, 년도 전체(이월 확인용)
WEEK_RE = re.compile(r"^(\d{1,2})월\s*(\d)주$")  # "9월 2주" 공백 변형 허용

rows = list(csv.reader(open(RAW, encoding="utf-8")))
# 열: A StyleNo+컬러, B 기획군, C 품종, D 시즌, E 년도, F 컬러, G 품명, H 기획수량,
#     I 수/C/완, J 원산지명, K 생산처명, L 입고원가, M 예상소매가, N 납기1차, O MIX1,
#     P 종결여부, Q 입고수량, R 총수불원가, S 지수, T 입고예정일, U 출고일, V 라벨, W MD, X DS
styles = OrderedDict()
bad_codes, week_bad = [], []
for r in rows:
    if not r or len(r) < 24:
        continue
    m = STYLE_RE.match(r[0].strip())
    if not m:
        if r[0].strip() and re.match(r"^[A-Z0-9]{10,}", r[0].strip()):
            bad_codes.append(r[0].strip())
        continue
    code, color = m.groups()
    s = styles.setdefault(code, {
        "StyleCode": code, "기획군": r[1].strip(), "품종": r[2].strip(), "시즌코드": r[3].strip(),
        "품명": r[6].strip(), "기획수량": 0, "수/C/완": "", "원산지": "", "생산처": "",
        "입고원가": "", "예상소매가": "", "종결여부": "", "입고예정일": "", "출고일(주차)": "",
        "라벨": "", "MD": "", "DS": "", "컬러구성": [], "출고일_상이": "",
    })
    s["컬러구성"].append(color)
    qty = r[7].replace(",", "").strip()
    if qty.isdigit():
        s["기획수량"] += int(qty)
    for key, idx in [("수/C/완", 8), ("원산지", 9), ("생산처", 10), ("입고원가", 11), ("예상소매가", 12),
                     ("종결여부", 15), ("입고예정일", 19), ("라벨", 21), ("MD", 22), ("DS", 23)]:
        v = r[idx].strip()
        if v and not s[key]:
            s[key] = v
    w = r[20].strip().replace(" ", "")
    if w:
        if not WEEK_RE.match(w):
            week_bad.append((code, w))
        if not s["출고일(주차)"]:
            s["출고일(주차)"] = w
        elif s["출고일(주차)"] != w:
            s["출고일_상이"] = f'{s["출고일(주차)"]} / {w}'

for s in styles.values():
    s["컬러구성"] = "/".join(s["컬러구성"])

with open(OUT, "w", newline="", encoding="utf-8") as f:
    w = csv.DictWriter(f, fieldnames=list(next(iter(styles.values())).keys()))
    w.writeheader()
    w.writerows(styles.values())

# ---- 검증 리포트 ----
n = len(styles)
season = Counter(s["시즌코드"] for s in styles.values())
with_week = [s for s in styles.values() if s["출고일(주차)"]]
closed = [s for s in styles.values() if s["종결여부"]]
diff = [s for s in styles.values() if s["출고일_상이"]]
print(f"스타일 {n} (시즌 {dict(season)}) / 출고일 보유 {len(with_week)} / 미입력 {n-len(with_week)}")
print(f"컬러행 총 {sum(len(s['컬러구성'].split('/')) for s in styles.values())} / 주차형식 오류 {len(week_bad)} / 컬러간 출고일 상이 {len(diff)}")
print(f"종결여부 값 분포: {Counter(s['종결여부'] for s in styles.values())}")
print(f"수/C/완 분포: {Counter(s['수/C/완'] for s in styles.values())}")
print(f"원산지 분포: {Counter(s['원산지'] for s in styles.values())}")
print(f"출고일 주차 분포: {sorted(Counter(s['출고일(주차)'] for s in with_week).items(), key=lambda x:(int(x[0].split('월')[0]), x[0]))}")
if week_bad[:5]: print("주차형식 오류 샘플:", week_bad[:5])
if bad_codes[:5]: print("품번형식 불일치 샘플:", bad_codes[:5], "총", len(bad_codes))
if diff[:5]: print("출고일 상이 샘플:", [(s['StyleCode'], s['출고일_상이']) for s in diff[:5]])
