"""샘플 판매 데이터(엑셀) 생성 스크립트.

실행: python3 generate_sample.py
결과: sales_data.xlsx (RawData 시트 + 컬럼설명 시트)
      sample_data.js  (dashboard.html 이 기본으로 띄우는 동일 데이터)
"""
import json
import random
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

random.seed(42)

TEAMS = {
    "영업1팀": ["김민준", "이서연", "박지호"],
    "영업2팀": ["최하은", "정도윤", "강서진", "윤예린"],
    "영업3팀": ["임건우", "한지민", "오세훈"],
}
# 제품: (단가, 월 기본수량, 분기 성장 추세)
PRODUCTS = {
    "프리미엄 A": (48000, 60, 0.06),
    "스탠다드 B": (22000, 140, 0.01),
    "라이트 C": (12000, 220, -0.03),
    "케어 D": (35000, 70, 0.09),
    "플러스 E": (61000, 30, 0.12),
    "베이직 F": (8000, 300, -0.05),
}
CLIENT_NAMES = [
    "한빛유통", "대성상사", "미래마트", "서울약품", "그린푸드", "동방무역", "우리헬스", "새롬상회",
    "태평양물산", "하나로스토어", "청솔유통", "명성메디칼", "누리마켓", "가온상사", "다온유통",
    "성원약국체인", "바른식품", "온누리몰", "제일상회", "한결유통", "신세계약품", "로뎀상사",
]
MONTHS = [f"{y}-{m:02d}" for y in (2025, 2026) for m in range(1, 13)][:21]  # 2025-01 ~ 2026-09

members = [(t, m) for t, ms in TEAMS.items() for m in ms]
# 거래처 → 담당자 배정, 거래처별 취급 제품 3~5개
clients = []
for i, name in enumerate(CLIENT_NAMES):
    team, member = members[i % len(members)]
    prods = random.sample(list(PRODUCTS), random.randint(3, 5))
    clients.append((name, team, member, prods, random.uniform(0.6, 1.6), random.uniform(-0.04, 0.06)))

rows = []
for mi, ym in enumerate(MONTHS):
    month = int(ym[5:])
    season = 1 + 0.12 * (month in (11, 12)) - 0.08 * (month in (1, 2)) + 0.05 * (month in (6, 7))
    q_idx = mi / 3
    for name, team, member, prods, scale, client_trend in clients:
        for p in prods:
            price, base, trend = PRODUCTS[p]
            qty = base * scale * season * (1 + trend + client_trend) ** q_idx * random.uniform(0.75, 1.25)
            qty = max(0, int(round(qty)))
            if random.random() < 0.04:  # 가끔 주문 없음
                continue
            unit = int(price * random.uniform(0.95, 1.03) / 10) * 10
            rows.append([ym, team, member, name, p, qty, qty * unit])

wb = Workbook()
ws = wb.active
ws.title = "RawData"
header = ["연월", "팀", "담당자", "거래처", "제품", "수량", "공급가액"]
ws.append(header)
for r in rows:
    ws.append(r)
hfill = PatternFill("solid", fgColor="1F4E9E")
for c in range(1, len(header) + 1):
    cell = ws.cell(row=1, column=c)
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = hfill
    cell.alignment = Alignment(horizontal="center")
    ws.column_dimensions[get_column_letter(c)].width = [10, 10, 10, 14, 12, 10, 16][c - 1]
for r in range(2, ws.max_row + 1):
    ws.cell(row=r, column=6).number_format = "#,##0"
    ws.cell(row=r, column=7).number_format = "#,##0"
ws.freeze_panes = "A2"
ws.auto_filter.ref = ws.dimensions

guide = wb.create_sheet("컬럼설명")
guide.append(["컬럼", "형식", "설명"])
for row in [
    ["연월", "YYYY-MM (텍스트) 또는 날짜", "판매 월. 분기는 자동 계산됨"],
    ["팀", "텍스트", "영업팀 이름"],
    ["담당자", "텍스트", "팀원(영업 담당자) 이름"],
    ["거래처", "텍스트", "거래처명"],
    ["제품", "텍스트", "제품명"],
    ["수량", "숫자", "판매 수량"],
    ["공급가액", "숫자(원)", "부가세 제외 공급가액. 대시보드의 매출·성장률은 이 값 기준"],
    [],
    ["사용법", "", "이 형식 그대로 RawData 시트에 행을 추가한 뒤, dashboard.html 의 '엑셀 불러오기'로 열면 됩니다."],
]:
    guide.append(row)
for c in range(1, 4):
    guide.cell(row=1, column=c).font = Font(bold=True, color="FFFFFF")
    guide.cell(row=1, column=c).fill = hfill
guide.column_dimensions["A"].width = 10
guide.column_dimensions["B"].width = 26
guide.column_dimensions["C"].width = 70

wb.save("sales_data.xlsx")
with open("sample_data.js", "w", encoding="utf-8") as f:
    f.write("// generate_sample.py 로 생성된 샘플 데이터 (sales_data.xlsx 와 동일)\n")
    f.write("window.SAMPLE_ROWS = ")
    json.dump([dict(zip(header, r)) for r in rows], f, ensure_ascii=False, separators=(",", ":"))
    f.write(";\n")
print(f"{len(rows)} rows written")
