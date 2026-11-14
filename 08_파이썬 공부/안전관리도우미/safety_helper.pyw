# -*- coding: utf-8 -*-
"""
안전관리 도우미 — 명령어 입력 없이 버튼만 눌러 쓰는 안전관리 파이썬 도구
  ① 위험성평가 위험도 자동계산 · 등급 색칠 · 개선 필요 목록
  ② 점검기록(여러 엑셀) 하나로 합치기 + 요약
  ③ 교육·검사·자격 만료일 알림 (D-day)
  ④ 내 PC 자료 키워드 검색 (PDF·엑셀·워드·한글·텍스트)
결과 파일은 이 프로그램이 있는 폴더의 [결과] 폴더에 저장됩니다.
"""
import os
import re
import sys
import zlib
import struct
import zipfile
import threading
import subprocess
from datetime import datetime, date
from pathlib import Path

# ======================================================================
#  사업장 기준에 맞게 바꿀 수 있는 설정 (이 부분만 수정)
# ======================================================================
# 위험성 = 가능성(빈도) × 중대성(강도)  →  등급 기준 (예시: 5×5 행렬)
# (상한값, 등급명) 순서. 사업장 위험성평가 실시규정에 맞게 수정하세요. [확인 필요]
GRADES = [(4, "낮음"), (9, "보통"), (14, "높음"), (25, "매우높음")]
ACTION_GRADES = ["높음", "매우높음"]          # 감소대책이 반드시 있어야 하는 등급
SCALE_MAX = 5                                  # 가능성·중대성 최대값 (3×3이면 3)

# 만료일 알림 구간 (일)
WARN_DAYS = [30, 60]

# 키워드 검색 기본값
DEFAULT_KEYWORDS = "중대재해, 위험성평가, 경영책임자, 안전보건관리체계"
# ======================================================================

APP_DIR = Path(sys.argv[0]).resolve().parent if sys.argv and sys.argv[0] else Path.cwd()
RESULT_DIR = APP_DIR / "결과"
FORM_DIR = APP_DIR / "양식"

GRADE_COLOR = {"낮음": "C6EFCE", "보통": "FFEB9C", "높음": "F8CBAD", "매우높음": "FF7C80"}
STATE_COLOR = {"기한 지남": "FF7C80", "30일 이내": "F8CBAD", "60일 이내": "FFEB9C", "정상": "C6EFCE",
               "날짜 없음": "D9D9D9"}


def need_libs():
    try:
        import pandas, openpyxl  # noqa: F401
        return None
    except ImportError as e:
        return str(e)


def stamp():
    return datetime.now().strftime("%Y%m%d_%H%M")


def grade_of(score):
    if score is None:
        return ""
    for limit, name in GRADES:
        if score <= limit:
            return name
    return GRADES[-1][1]


# ----------------------------------------------------------------------
#  공통: 서식
# ----------------------------------------------------------------------
def style_sheet(ws, header_row=1, widths=None, wrap_cols=()):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    thin = Side(style="thin", color="BFBFBF")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    for cell in ws[header_row]:
        cell.font = Font(name="맑은 고딕", bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E79")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = box
    for row in ws.iter_rows(min_row=header_row + 1, max_row=ws.max_row):
        for cell in row:
            cell.font = Font(name="맑은 고딕", size=10)
            cell.border = box
            cell.alignment = Alignment(vertical="top", wrap_text=cell.column_letter in wrap_cols)
    if widths:
        for col, w in widths.items():
            ws.column_dimensions[col].width = w
    ws.freeze_panes = ws.cell(row=header_row + 1, column=1)


def fill(cell, hex_color):
    from openpyxl.styles import PatternFill
    if hex_color:
        cell.fill = PatternFill("solid", fgColor=hex_color)


def find_header(path, must_have, sheet=0):
    """머리글 줄 위치를 자동으로 찾기 (위에 제목 줄이 있어도 됨)"""
    import pandas as pd
    raw = pd.read_excel(path, sheet_name=sheet, header=None, nrows=15, dtype=str)
    for i, row in raw.iterrows():
        vals = [str(v).replace(" ", "").replace("\n", "") for v in row.dropna()]
        if all(any(m in v for v in vals) for m in must_have):
            return i
    return None


def pick_col(df, *names):
    """열 이름이 조금 달라도 찾기 (띄어쓰기·괄호 무시)"""
    norm = {re.sub(r"[\s()（）\n]", "", str(c)): c for c in df.columns}
    for n in names:
        n2 = re.sub(r"[\s()（）\n]", "", n)
        for k, c in norm.items():
            if k.startswith(n2) or n2 in k:
                return c
    return None


# ----------------------------------------------------------------------
#  ① 위험성평가
# ----------------------------------------------------------------------
RA_HEADERS = ["No", "공정(작업)", "유해·위험요인", "재해형태", "현재 안전대책",
              "가능성(빈도)", "중대성(강도)", "위험성", "등급",
              "감소대책", "개선후 가능성", "개선후 중대성", "개선후 위험성", "개선후 등급",
              "담당자", "완료예정일", "완료일"]


def ra_grade_formula(cell):
    expr = f'"{GRADES[-1][1]}"'
    for limit, name in reversed(GRADES[:-1]):
        expr = f'IF({cell}<={limit},"{name}",{expr})'
    return f'=IF({cell}="","",{expr})'


def make_ra_form():
    from openpyxl import Workbook
    from openpyxl.styles import Font
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.styles import PatternFill
    from openpyxl.worksheet.datavalidation import DataValidation
    FORM_DIR.mkdir(exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "위험성평가"
    ws["A1"] = "위험성평가표"
    ws["A1"].font = Font(name="맑은 고딕", size=15, bold=True, color="1F4E79")
    ws["A2"] = (f"노란 칸만 입력 · 가능성/중대성 1~{SCALE_MAX} · 위험성·등급은 자동 계산 · "
                "등급 기준은 [기준] 시트 (예시 — 사업장 규정에 맞게 수정)")
    ws["A2"].font = Font(name="맑은 고딕", size=9, color="C00000")
    ws.append(RA_HEADERS)
    examples = [
        [1, "자재 운반", "지게차 후진 시 작업자 충돌", "부딪힘", "후방경보기 설치", 3, 4, None, None,
         "보행자 통로 분리, 유도자 배치", 2, 4, None, None, "홍길동", "2026-11-30", None],
        [2, "고소작업", "이동식 사다리 작업 중 추락", "떨어짐", "2인 1조 작업", 3, 5, None, None,
         "", None, None, None, None, "", "", None],
        [3, "용접", "용접 불티에 의한 화재", "화재", "소화기 비치", 2, 3, None, None,
         "불티방지포 설치, 화기작업허가", 1, 3, None, None, "김철수", "2026-10-31", None],
    ]
    for r, ex in enumerate(examples, start=4):
        for c, v in enumerate(ex, start=1):
            ws.cell(row=r, column=c, value=v)
    for r in range(4, 104):
        ws.cell(row=r, column=1, value=r - 3)
        ws.cell(row=r, column=8, value=f'=IF(OR(F{r}="",G{r}=""),"",F{r}*G{r})')
        ws.cell(row=r, column=9, value=ra_grade_formula(f"H{r}"))
        ws.cell(row=r, column=13, value=f'=IF(OR(K{r}="",L{r}=""),"",K{r}*L{r})')
        ws.cell(row=r, column=14, value=ra_grade_formula(f"M{r}"))
    style_sheet(ws, header_row=3,
                widths={"A": 5, "B": 14, "C": 32, "D": 10, "E": 24, "F": 8, "G": 8, "H": 8, "I": 9,
                        "J": 30, "K": 8, "L": 8, "M": 8, "N": 9, "O": 10, "P": 12, "Q": 12},
                wrap_cols=("C", "E", "J"))
    for r in range(4, 104):
        for col in "BCDEFGJKLOPQ":
            fill(ws[f"{col}{r}"], "FFF9E5")
    for col in ("I", "N"):
        for name, color in GRADE_COLOR.items():
            ws.conditional_formatting.add(
                f"{col}4:{col}103",
                FormulaRule(formula=[f'{col}4="{name}"'], fill=PatternFill("solid", fgColor=color)))
    dv = DataValidation(type="whole", operator="between", formula1="1", formula2=str(SCALE_MAX),
                        showErrorMessage=True, error=f"1~{SCALE_MAX} 숫자만 입력")
    ws.add_data_validation(dv)
    for col in "FGKL":
        dv.add(f"{col}4:{col}103")

    lg = wb.create_sheet("기준")
    lg["A1"] = "위험성 판단 기준 (예시 — 사업장 위험성평가 실시규정에 맞게 수정, 확인 필요)"
    lg["A1"].font = Font(name="맑은 고딕", size=12, bold=True, color="C00000")
    lg.append([])
    lg.append(["위험성(가능성×중대성)", "등급", "조치"])
    prev = 1
    acts = {"낮음": "현 상태 유지·관리", "보통": "계획적 개선", "높음": "감소대책 수립 후 작업",
            "매우높음": "즉시 개선, 개선 전 작업중지 검토"}
    for limit, name in GRADES:
        lg.append([f"{prev} ~ {limit}", name, acts.get(name, "")])
        fill(lg.cell(row=lg.max_row, column=2), GRADE_COLOR.get(name))
        prev = limit + 1
    style_sheet(lg, header_row=3, widths={"A": 24, "B": 12, "C": 40})
    lg.freeze_panes = None
    out = FORM_DIR / "위험성평가_양식.xlsx"
    wb.save(out)
    return out


def run_ra(path):
    import pandas as pd
    from openpyxl import load_workbook
    hr = find_header(path, ["유해", "가능성", "중대성"])
    if hr is None:
        raise ValueError("머리글을 찾지 못했습니다. '유해·위험요인', '가능성', '중대성' 열이 있어야 합니다.")
    df = pd.read_excel(path, header=hr)
    c_haz = pick_col(df, "유해·위험요인", "유해위험요인", "위험요인")
    c_p, c_s = pick_col(df, "가능성"), pick_col(df, "중대성")
    c_p2, c_s2 = pick_col(df, "개선후가능성"), pick_col(df, "개선후중대성")
    c_act = pick_col(df, "감소대책", "개선대책")
    if c_p2 == c_p:
        c_p2 = None
    if c_s2 == c_s:
        c_s2 = None
    df = df[df[c_haz].notna() & (df[c_haz].astype(str).str.strip() != "")].copy()

    def num(s):
        return pd.to_numeric(s, errors="coerce")

    log = []
    bad = df[(num(df[c_p]).isna()) | (num(df[c_s]).isna()) |
             (num(df[c_p]) > SCALE_MAX) | (num(df[c_s]) > SCALE_MAX)]
    for i in bad.index:
        log.append(f"입력 확인: '{str(df.at[i, c_haz])[:20]}' — 가능성/중대성 값이 비었거나 범위 밖")

    df["위험성"] = (num(df[c_p]) * num(df[c_s]))
    df["등급"] = df["위험성"].apply(lambda v: grade_of(v) if pd.notna(v) else "")
    if c_p2 and c_s2:
        df["개선후 위험성"] = num(df[c_p2]) * num(df[c_s2])
        df["개선후 등급"] = df["개선후 위험성"].apply(lambda v: grade_of(v) if pd.notna(v) else "")
    else:
        df["개선후 위험성"], df["개선후 등급"] = None, ""

    need = df[df["등급"].isin(ACTION_GRADES)]
    no_action = need[need[c_act].isna() | (need[c_act].astype(str).str.strip() == "")] if c_act else need
    still = df[df["개선후 등급"].isin(ACTION_GRADES)]

    order = [g for _, g in GRADES]
    summary = pd.DataFrame({
        "등급": order,
        "개선 전 (건)": [int((df["등급"] == g).sum()) for g in order],
        "개선 후 (건)": [int((df["개선후 등급"] == g).sum()) for g in order],
    })
    # 같은 이름의 열(위험성·등급 등)은 위에서 계산값으로 덮어씀, 없으면 맨 뒤에 추가됨
    full = df[[c for c in df.columns if not str(c).startswith("Unnamed")]]

    RESULT_DIR.mkdir(exist_ok=True)
    out = RESULT_DIR / f"위험성평가_결과_{stamp()}.xlsx"
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        summary.to_excel(w, sheet_name="요약", index=False)
        full.to_excel(w, sheet_name="전체(계산)", index=False)
        no_action[[c_haz, c_p, c_s, "위험성", "등급"]].to_excel(w, sheet_name="감소대책 미작성", index=False)
        if len(still):
            still[[c_haz, "개선후 위험성", "개선후 등급"]].to_excel(w, sheet_name="개선후에도 높음", index=False)
    wb = load_workbook(out)
    for ws in wb.worksheets:
        style_sheet(ws, widths={col: 14 for col in "ABCDEFGHIJKLMNOPQRS"}, wrap_cols=("C", "E", "J"))
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                if isinstance(cell.value, str) and cell.value in GRADE_COLOR:
                    fill(cell, GRADE_COLOR[cell.value])
    wb.save(out)
    msg = [f"평가 항목 {len(df)}건 계산 완료",
           "등급별: " + ", ".join(f"{g} {n}건" for g, n in zip(order, summary['개선 전 (건)'])),
           f"감소대책 미작성 ({'/'.join(ACTION_GRADES)}): {len(no_action)}건",
           f"개선 후에도 {'/'.join(ACTION_GRADES)}: {len(still)}건"] + log[:10]
    return out, msg


# ----------------------------------------------------------------------
#  ② 점검기록 합치기
# ----------------------------------------------------------------------
def run_merge(folder, group_col, sum_col):
    import pandas as pd
    frames, log = [], []
    files = sorted(p for p in Path(folder).rglob("*.xls*") if not p.name.startswith("~$"))
    if not files:
        raise ValueError("폴더에 엑셀 파일(.xlsx)이 없습니다.")
    for p in files:
        try:
            sheets = pd.read_excel(p, sheet_name=None)
        except Exception as e:
            log.append(f"읽기 실패: {p.name} ({type(e).__name__})")
            continue
        for sname, df in sheets.items():
            df = df.dropna(how="all")
            if df.empty:
                continue
            df.insert(0, "출처파일", p.name)
            df.insert(1, "출처시트", sname)
            frames.append(df)
            log.append(f"{p.name} [{sname}] {len(df)}건")
    if not frames:
        raise ValueError("읽을 수 있는 데이터가 없습니다.")
    total = pd.concat(frames, ignore_index=True)
    total = total[[c for c in total.columns if not str(c).startswith("Unnamed")]]
    by_file = total.groupby("출처파일").size().reset_index(name="건수")

    RESULT_DIR.mkdir(exist_ok=True)
    out = RESULT_DIR / f"점검기록_통합_{stamp()}.xlsx"
    pivot = None
    g = pick_col(total, group_col) if group_col else None
    s = pick_col(total, sum_col) if sum_col else None
    if g is not None:
        if s is not None:
            total[s] = pd.to_numeric(total[s], errors="coerce")
            pivot = total.groupby(g).agg(건수=(s, "size"), 합계=(s, "sum")).reset_index()
        else:
            pivot = total.groupby(g).size().reset_index(name="건수")
        pivot = pivot.sort_values(pivot.columns[-1], ascending=False)
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        total.to_excel(w, sheet_name="전체기록", index=False)
        by_file.to_excel(w, sheet_name="파일별건수", index=False)
        if pivot is not None:
            pivot.to_excel(w, sheet_name=f"{g}별요약"[:31], index=False)
    from openpyxl import load_workbook
    wb = load_workbook(out)
    for ws in wb.worksheets:
        style_sheet(ws, widths={col: 15 for col in "ABCDEFGHIJKLMNOP"})
    wb.save(out)
    msg = [f"파일 {len(files)}개, 합계 {len(total)}건 통합"]
    if group_col and g is None:
        msg.append(f"'{group_col}' 열이 없어 요약표는 만들지 않음 (열 이름 확인)")
    elif pivot is not None:
        msg.append(f"'{g}'별 요약표 생성")
    return out, msg + log[:15]


# ----------------------------------------------------------------------
#  ③ 만료일 알림
# ----------------------------------------------------------------------
EX_HEADERS = ["No", "구분", "대상(설비·인원)", "내용", "최근 실시일", "주기(개월)", "차기 예정일", "담당자", "비고"]


def make_expiry_form():
    from openpyxl import Workbook
    from openpyxl.styles import Font
    FORM_DIR.mkdir(exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "만료관리"
    ws["A1"] = "교육·검사·자격 만료 관리대장"
    ws["A1"].font = Font(name="맑은 고딕", size=15, bold=True, color="1F4E79")
    ws["A2"] = "노란 칸 입력 · 차기 예정일은 자동 계산(최근 실시일+주기) · 주기는 법정·사내 기준을 담당자가 확인해 입력 [확인 필요]"
    ws["A2"].font = Font(name="맑은 고딕", size=9, color="C00000")
    ws.append(EX_HEADERS)
    today = date.today()
    ex = [
        [1, "검사", "OO 설비 (예시)", "정기 검사", date(today.year - 1, today.month, 1), 12],
        [2, "교육", "OOO 외 5명 (예시)", "정기 교육", date(today.year, max(1, today.month - 2), 15), 3],
        [3, "자격", "OOO (예시)", "자격 갱신", date(today.year - 2, 12, 31), 36],
    ]
    for r, row in enumerate(ex, start=4):
        for c, v in enumerate(row, start=1):
            ws.cell(row=r, column=c, value=v)
    for r in range(4, 204):
        ws.cell(row=r, column=1, value=r - 3)
        ws.cell(row=r, column=7, value=f'=IF(OR(E{r}="",F{r}=""),"",EDATE(E{r},F{r}))')
        ws.cell(row=r, column=5).number_format = "yyyy-mm-dd"
        ws.cell(row=r, column=7).number_format = "yyyy-mm-dd"
    style_sheet(ws, header_row=3,
                widths={"A": 5, "B": 10, "C": 24, "D": 26, "E": 13, "F": 10, "G": 13, "H": 10, "I": 24})
    for r in range(4, 204):
        for col in "BCDEFHI":
            fill(ws[f"{col}{r}"], "FFF9E5")
    out = FORM_DIR / "만료관리_양식.xlsx"
    wb.save(out)
    return out


def run_expiry(path):
    import pandas as pd
    hr = find_header(path, ["대상", "실시"]) or find_header(path, ["대상", "예정"])
    if hr is None:
        raise ValueError("머리글을 찾지 못했습니다. '대상', '최근 실시일' 또는 '차기 예정일' 열이 필요합니다.")
    df = pd.read_excel(path, header=hr)
    c_tgt = pick_col(df, "대상")
    c_last, c_cycle, c_next = pick_col(df, "최근실시일", "실시일"), pick_col(df, "주기"), pick_col(df, "차기예정일", "예정일", "만료일")
    df = df[df[c_tgt].notna() & (df[c_tgt].astype(str).str.strip() != "")].copy()
    nxt = pd.to_datetime(df[c_next], errors="coerce") if c_next else pd.Series(pd.NaT, index=df.index)
    if c_last and c_cycle:
        last = pd.to_datetime(df[c_last], errors="coerce")
        cyc = pd.to_numeric(df[c_cycle], errors="coerce")
        calc = [l + pd.DateOffset(months=int(c)) if pd.notna(l) and pd.notna(c) else pd.NaT
                for l, c in zip(last, cyc)]
        nxt = nxt.fillna(pd.Series(calc, index=df.index))
    today = pd.Timestamp(date.today())
    df["차기 예정일(계산)"] = nxt.dt.date
    df["남은 일수"] = (nxt - today).dt.days

    def state(d):
        if pd.isna(d):
            return "날짜 없음"
        if d < 0:
            return "기한 지남"
        for w in WARN_DAYS:
            if d <= w:
                return f"{w}일 이내"
        return "정상"
    df["상태"] = df["남은 일수"].apply(state)
    df = df.sort_values("남은 일수", na_position="last")
    df = df[[c for c in df.columns if not str(c).startswith("Unnamed")]]
    if c_next:
        df = df.drop(columns=[c_next])
    alert = df[df["상태"] != "정상"]

    RESULT_DIR.mkdir(exist_ok=True)
    out = RESULT_DIR / f"만료알림_{stamp()}.xlsx"
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        alert.to_excel(w, sheet_name="조치필요", index=False)
        df.to_excel(w, sheet_name="전체", index=False)
    from openpyxl import load_workbook
    wb = load_workbook(out)
    for ws in wb.worksheets:
        style_sheet(ws, widths={col: 14 for col in "ABCDEFGHIJKLM"})
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                if isinstance(cell.value, str) and cell.value in STATE_COLOR:
                    fill(cell, STATE_COLOR[cell.value])
                if hasattr(cell.value, "year") and not isinstance(cell.value, str):
                    cell.number_format = "yyyy-mm-dd"
    wb.save(out)
    counts = df["상태"].value_counts()
    msg = [f"기준일 {date.today()} · 전체 {len(df)}건"] + [f"{k}: {v}건" for k, v in counts.items()]
    for _, r in alert.head(10).iterrows():
        msg.append(f"  · {r[c_tgt]} — {r['상태']} ({'' if pd.isna(r['남은 일수']) else int(r['남은 일수'])}일)")
    return out, msg


# ----------------------------------------------------------------------
#  ④ 자료 키워드 검색
# ----------------------------------------------------------------------
def _pdf(p):
    from pypdf import PdfReader
    return [(f"{i}쪽", pg.extract_text() or "") for i, pg in enumerate(PdfReader(p).pages, 1)]


def _xlsx(p):
    import pandas as pd
    out = []
    for name, df in pd.read_excel(p, sheet_name=None, header=None, dtype=str).items():
        for r, row in df.iterrows():
            line = " | ".join(v for v in row.dropna())
            if line:
                out.append((f"{name}시트 {r + 1}행", line))
    return out


def _docx(p):
    import docx
    d = docx.Document(p)
    out = [(f"{i}번째 문단", x.text) for i, x in enumerate(d.paragraphs, 1) if x.text.strip()]
    for t, tb in enumerate(d.tables, 1):
        for r, row in enumerate(tb.rows, 1):
            out.append((f"표{t} {r}행", " | ".join(c.text for c in row.cells)))
    return out


def _hwpx(p):
    out = []
    with zipfile.ZipFile(p) as z:
        for n in sorted(x for x in z.namelist() if re.match(r"Contents/section\d+\.xml", x)):
            xml = z.read(n).decode("utf-8", errors="ignore")
            for i, para in enumerate(re.findall(r"<hp:p\b.*?</hp:p>", xml, re.S), 1):
                t = re.sub(r"<[^>]+>", "", "".join(re.findall(r"<hp:t\b[^>]*>(.*?)</hp:t>", para, re.S)))
                if t.strip():
                    out.append((f"{Path(n).stem} {i}문단", t))
    return out


def _hwp(p):
    import olefile
    f = olefile.OleFileIO(str(p))
    comp = f.openstream("FileHeader").read()[36] & 1
    out = []
    for n in sorted(int(d[1][7:]) for d in f.listdir() if d[0] == "BodyText"):
        data = f.openstream(f"BodyText/Section{n}").read()
        if comp:
            data = zlib.decompress(data, -15)
        i = 0
        while i < len(data):
            h = struct.unpack_from("<I", data, i)[0]
            tag, size = h & 0x3FF, (h >> 20) & 0xFFF
            i += 4
            if size == 0xFFF:
                size = struct.unpack_from("<I", data, i)[0]
                i += 4
            if tag == 67:
                t = "".join(c for c in data[i:i + size].decode("utf-16-le", errors="ignore") if c.isprintable())
                if t.strip():
                    out.append((f"구역{n}", t))
            i += size
    f.close()
    return out


def _txt(p):
    t = Path(p).read_text(encoding="utf-8", errors="ignore")
    return [(f"{i}행", x) for i, x in enumerate(t.splitlines(), 1)]


READERS = {".pdf": _pdf, ".xlsx": _xlsx, ".xlsm": _xlsx, ".docx": _docx,
           ".hwpx": _hwpx, ".hwp": _hwp, ".txt": _txt, ".csv": _txt}


def run_search(folder, keywords, progress=None):
    import pandas as pd
    kws = [k.strip() for k in re.split(r"[,，\n]", keywords) if k.strip()]
    if not kws:
        raise ValueError("검색어를 입력하세요.")
    files = [p for p in Path(folder).rglob("*") if p.is_file() and not p.name.startswith("~$")]
    res, err = [], []
    for n, p in enumerate(files, 1):
        if progress and n % 10 == 0:
            progress(f"검색 중... {n}/{len(files)}")
        for k in kws:
            if k in p.name:
                res.append([k, "파일명", p.name, "-", p.name, str(p)])
        rd = READERS.get(p.suffix.lower())
        if not rd:
            continue
        try:
            for where, text in rd(p):
                for k in kws:
                    if k in text:
                        t = re.sub(r"\s+", " ", text)
                        j = t.find(k)
                        res.append([k, "내용", p.name, where, t[max(0, j - 40): j + len(k) + 40], str(p)])
        except Exception as e:
            err.append([p.name, type(e).__name__, str(p)])
    df = pd.DataFrame(res, columns=["키워드", "찾은곳", "파일명", "위치", "앞뒤 문장", "전체경로"])
    RESULT_DIR.mkdir(exist_ok=True)
    out = RESULT_DIR / f"자료검색_{stamp()}.xlsx"
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="검색결과", index=False)
        if len(df):
            df.groupby("키워드").size().reset_index(name="건수").to_excel(w, sheet_name="키워드별", index=False)
            df.groupby("파일명").size().reset_index(name="건수").sort_values("건수", ascending=False) \
              .to_excel(w, sheet_name="파일별", index=False)
        pd.DataFrame(err, columns=["파일명", "오류", "전체경로"]).to_excel(w, sheet_name="읽기실패", index=False)
    from openpyxl import load_workbook
    from openpyxl.styles import Font
    wb = load_workbook(out)
    for ws in wb.worksheets:
        style_sheet(ws, widths={"A": 14, "B": 10, "C": 30, "D": 16, "E": 70, "F": 50}, wrap_cols=("E",))
    ws = wb["검색결과"]
    for r in range(2, ws.max_row + 1):
        c = ws.cell(row=r, column=6)
        c.hyperlink = c.value
        c.font = Font(name="맑은 고딕", size=10, color="0563C1", underline="single")
    wb.save(out)
    return out, [f"파일 {len(files)}개 검색 · 찾은 건수 {len(df)}건 · 읽기 실패 {len(err)}개",
                 "결과의 '전체경로'를 클릭하면 파일이 열립니다."]


# ----------------------------------------------------------------------
#  화면 (버튼)
# ----------------------------------------------------------------------
def open_path(p):
    try:
        if sys.platform.startswith("win"):
            os.startfile(str(p))  # noqa
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(p)])
        else:
            subprocess.Popen(["xdg-open", str(p)])
    except Exception:
        pass


def main():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, simpledialog

    root = tk.Tk()
    root.title("안전관리 도우미")
    root.geometry("760x640")
    root.configure(bg="#F4F6F8")
    FONT = ("맑은 고딕", 10)
    BIG = ("맑은 고딕", 15, "bold")

    missing = need_libs()

    tk.Label(root, text="안전관리 도우미", font=BIG, bg="#F4F6F8", fg="#1F4E79").pack(anchor="w", padx=16, pady=(14, 0))
    tk.Label(root, text="명령어 입력 없이 버튼만 누르세요 · 결과는 [결과] 폴더에 저장됩니다",
             font=FONT, bg="#F4F6F8", fg="#595959").pack(anchor="w", padx=16, pady=(0, 8))

    logbox = tk.Text(root, height=10, font=("맑은 고딕", 9), bg="white", relief="solid", bd=1)

    def log(lines):
        if isinstance(lines, str):
            lines = [lines]
        for ln in lines:
            logbox.insert("end", ln + "\n")
        logbox.see("end")
        root.update_idletasks()

    def guard():
        if missing:
            messagebox.showerror("설치 필요", "필요한 프로그램이 설치되지 않았습니다.\n\n"
                                 "[1_처음설치(setup).bat] 을 더블클릭해 설치한 뒤 다시 실행하세요.\n\n" + missing)
            return False
        return True

    def done(out, msg):
        log(msg + [f"→ 저장: {out}", ""])
        if messagebox.askyesno("완료", "\n".join(msg[:6]) + f"\n\n결과 파일을 열까요?\n{out.name}"):
            open_path(out)

    def task(fn, *args):
        def work():
            try:
                out, msg = fn(*args)
                root.after(0, lambda: done(out, msg))
            except PermissionError:
                root.after(0, lambda: messagebox.showerror("오류", "엑셀에서 파일이 열려 있습니다. 닫고 다시 실행하세요."))
            except Exception as e:
                err = f"{type(e).__name__}: {e}"
                root.after(0, lambda: (log("오류 — " + err), messagebox.showerror("오류", err)))
        log("처리 중...")
        threading.Thread(target=work, daemon=True).start()

    def form(fn):
        if not guard():
            return
        try:
            out = fn()
            log(f"양식 생성: {out}")
            open_path(out)
        except PermissionError:
            messagebox.showerror("오류", "같은 이름의 양식 파일이 열려 있습니다. 닫고 다시 누르세요.")

    def ra():
        if not guard():
            return
        p = filedialog.askopenfilename(title="위험성평가 엑셀 선택", filetypes=[("엑셀", "*.xlsx *.xlsm")])
        if p:
            log(f"[위험성평가] {Path(p).name}")
            task(run_ra, p)

    def merge():
        if not guard():
            return
        d = filedialog.askdirectory(title="합칠 엑셀 파일들이 있는 폴더 선택")
        if not d:
            return
        g = simpledialog.askstring("요약 기준", "무엇별로 요약할까요? (열 이름, 비워두면 생략)\n예: 구역, 부서, 협력업체",
                                   initialvalue="구역", parent=root)
        s = simpledialog.askstring("합계 열", "합계를 낼 숫자 열 이름 (비워두면 건수만)\n예: 지적건수",
                                   initialvalue="지적건수", parent=root)
        log(f"[점검기록 통합] {d}")
        task(run_merge, d, (g or "").strip(), (s or "").strip())

    def expiry():
        if not guard():
            return
        p = filedialog.askopenfilename(title="만료관리 엑셀 선택", filetypes=[("엑셀", "*.xlsx *.xlsm")])
        if p:
            log(f"[만료일 알림] {Path(p).name}")
            task(run_expiry, p)

    def search():
        if not guard():
            return
        d = filedialog.askdirectory(title="검색할 폴더 선택 (하위 폴더 포함)")
        if not d:
            return
        k = simpledialog.askstring("검색어", "찾을 단어 (쉼표로 구분)", initialvalue=DEFAULT_KEYWORDS, parent=root)
        if k:
            log(f"[자료 검색] {d} — {k}")
            task(run_search, d, k, lambda m: root.after(0, lambda: log(m)))

    frame = tk.Frame(root, bg="#F4F6F8")
    frame.pack(fill="x", padx=16)
    items = [
        ("① 위험성평가 계산", "가능성×중대성 → 위험성·등급 자동 계산, 감소대책 누락 확인", ra, make_ra_form),
        ("② 점검기록 합치기", "폴더 안 엑셀 여러 개를 한 파일로 + 구역·부서별 요약", merge, None),
        ("③ 만료일 알림", "교육·검사·자격 차기 예정일 D-day, 30·60일 이내 표시", expiry, make_expiry_form),
        ("④ 자료 키워드 검색", "폴더 안 PDF·엑셀·워드·한글에서 단어 찾기", search, None),
    ]
    style = ttk.Style()
    style.configure("Run.TButton", font=("맑은 고딕", 10, "bold"))
    style.configure("Form.TButton", font=FONT)
    for r, (title, desc, fn, fform) in enumerate(items):
        box = tk.Frame(frame, bg="white", highlightbackground="#BFBFBF", highlightthickness=1)
        box.grid(row=r, column=0, sticky="ew", pady=4)
        frame.columnconfigure(0, weight=1)
        tk.Label(box, text=title, font=("맑은 고딕", 11, "bold"), bg="white", fg="#1F4E79",
                 width=18, anchor="w").pack(side="left", padx=10, pady=10)
        tk.Label(box, text=desc, font=FONT, bg="white", anchor="w", justify="left",
                 wraplength=300).pack(side="left", fill="x", expand=True)
        ttk.Button(box, text="실행", style="Run.TButton", command=fn, width=8).pack(side="right", padx=8)
        if fform:
            ttk.Button(box, text="양식 만들기", style="Form.TButton",
                       command=lambda f=fform: form(f), width=10).pack(side="right")

    bar = tk.Frame(root, bg="#F4F6F8")
    bar.pack(fill="x", padx=16, pady=(8, 2))
    tk.Label(bar, text="진행 기록", font=("맑은 고딕", 10, "bold"), bg="#F4F6F8").pack(side="left")
    ttk.Button(bar, text="결과 폴더 열기",
               command=lambda: (RESULT_DIR.mkdir(exist_ok=True), open_path(RESULT_DIR))).pack(side="right")
    logbox.pack(fill="both", expand=True, padx=16, pady=(0, 14))
    if missing:
        log("※ 설치가 필요합니다 → [1_처음설치(setup).bat] 더블클릭")
    else:
        log("준비 완료. 처음이면 ①·③의 [양식 만들기]로 양식부터 받아 보세요.")
    root.mainloop()


if __name__ == "__main__":
    main()
