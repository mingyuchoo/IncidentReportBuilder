# -*- coding: utf-8 -*-
"""
장애 월간보고 - PPT 생성 프로그램 (make_monthly_ppt.py)
======================================================

동료 주간 도구(make_ppt.py)의 헬퍼를 재사용해, update_excel.py 로 갱신된 엑셀을 읽어
지난달 "장애현황" PPT(=템플릿 겸 비교본)의 차트/서술문구/제목을 채워
이번달 월간 "장애현황" PPT를 만든다.

  · 기존 make_ppt.py / 주간 흐름은 전혀 건드리지 않는다 (헬퍼만 import).
  · 기준일 해석 = 기준일이 속한 달 1일 ~ 기준일 (해당 월 진행분, MTD).
  · 결과 파일명 = "26년 {M}월 장애현황_기준일(YYYYMMDD).pptx"

── 갱신 대상 (월간 15장 기준, v0.1 = 차트 + 서술문구 + 제목/날짜) ──
  · 슬라이드 0 (표지): 사업부별 서술 요약 박스 4개 + 등급내역 + 전사 누적(요약 시트)
  · 슬라이드 1   : 사업부 4개 추이차트 (사업부별 시트)  + 연도 라벨 1월 정렬
  · 슬라이드 2   : 사업부 4개 추이차트 (사업부별_계열사별 시트) + 연도 라벨 1월 정렬
  · 별첨1(3·7·10·12): 각 사업부 대형 추이차트 + 서술문구 + 전사누적 + 날짜 라벨
  · 각 슬라이드의 'M/D 기준' 등 날짜 표기를 기준일로 갱신
  * 별첨2 상세표(등급확정/협의중)는 다음 단계에서 작성 (v0.1 미포함)

요구사항: python-pptx, openpyxl  (make_ppt.py, appconfig.py 와 같은 폴더에서 실행)
사용법:
    python make_monthly_ppt.py --date 2026-08-11 --prev "26년 7월 장애현황_v1.0.pptx"
"""

import argparse
import copy
import datetime as dt
import os
import re
import sys

import openpyxl
from pptx import Presentation
from pptx.oxml.ns import qn

import appconfig
import make_ppt as W          # 주간 도구의 헬퍼 재사용 (import 시 부작용 없음: main 은 __main__ 가드)


# ─────────────────────────────────────────────────────────────────────────────
# 요약 시트 셀 매핑 (역설계) — 서술형 문구는 모두 '요약' 시트에 이미 존재
# ─────────────────────────────────────────────────────────────────────────────
# 요약 시트 셀 매핑은 회사 고유값이므로 소스에 하드코딩하지 않고 report_config.json 의
# "monthly" 섹션에서 로드한다 (appconfig 가 이미 로드한 설정 재사용).
#   total_cells     : 전사 누적   (narr/avg/delta = C22/C23/C24)
#   affiliate_cells : 계열사 전사 (C135/C136/C137)
#   division_cells  : 사업부별    {사업부명: {narr, grade}}
# 값이 없으면 안전한 기본값(전사/계열사 셀)만 두고 사업부 매핑은 비운다.
_MONTHLY = dict(getattr(appconfig, "_cfg", {}).get("monthly", {}))
TOTAL_CELLS = dict(_MONTHLY.get("total_cells", {"narr": "C22", "avg": "C23", "delta": "C24"}))
AFFIL_CELLS = dict(_MONTHLY.get("affiliate_cells", {"narr": "C135", "avg": "C136", "delta": "C137"}))
DIV_CELLS = dict(_MONTHLY.get("division_cells", {}))
DIVISIONS = list(DIV_CELLS.keys())


def cell(ws, coord):
    v = ws[coord].value
    return "" if v is None else str(v)


def polish_delta_m(text):
    """월간 요약 시트의 증감 문구 정리.
    - 'N% 증가or감소' / 'N% 감소/증가' 등 미정 표현을 부호(-)로 방향 확정
    - 백분율 앞 부호 제거, 천단위 콤마 적용
    (주간 make_ppt.polish_delta 는 '감소/증가' 만 처리 → 월간의 '증가or감소' 토큰까지 확장)"""
    def repl(m):
        sign, num = m.group(1), m.group(2)
        direction = "감소" if sign == "-" else "증가"
        return f"{num}% {direction}"

    text = re.sub(
        r"([+-]?)(\d+(?:\.\d+)?)\s*%\s*(?:증가or감소|감소or증가|감소/증가|증가/감소|감소|증가)",
        repl, text)
    return W.group_commas(text)


def clean_cell(ws, coord):
    """요약 셀 값을 읽어 증감/콤마 정리한 문자열 반환 (#REF! 등 오류는 빈 문자열)."""
    v = cell(ws, coord)
    if not v or v.startswith("#"):
        return ""
    return polish_delta_m(v)


def which_division(text):
    """텍스트가 어떤 사업부로 시작/포함하는지 판별 (없으면 None)."""
    for d in DIVISIONS:
        if d in text:
            return d
    return None


def update_div_slide(slide, divs, categories, label):
    """슬라이드의 사업부 차트(제목=사업부명)를 해당 시리즈로 갱신.
    (make_ppt.main() 안의 중첩함수와 동일 로직 — import 불가라 월간 모듈에 인라인)"""
    cnt = 0
    for sh in slide.shapes:
        if not sh.has_chart:
            continue
        ch = sh.chart
        title = ch.chart_title.text_frame.text if ch.has_title else None
        if title in divs:
            W.replace_chart(sh, categories, divs[title])
            cnt += 1
    W.log(f"  [{label}] 사업부 차트 {cnt}개 갱신")


# ─────────────────────────────────────────────────────────────────────────────
# 슬라이드 0 (표지): 사업부 서술박스 + 등급내역 + 전사 누적
# ─────────────────────────────────────────────────────────────────────────────
def fill_cover(slide, ws, base_date):
    n_div = n_grade = n_total = 0
    for sh in slide.shapes:
        if not sh.has_text_frame:
            continue
        txt = sh.text_frame.text.strip()
        if not txt:
            continue
        # 전사 누적 박스 (동일 문구 4개)
        if "전사 누적" in txt:
            W.set_shape_text(sh, clean_cell(ws, TOTAL_CELLS["narr"]))
            n_total += 1
            continue
        # 등급내역 소박스: "(2등급 …건 …분, 3등급 이하 …건 …분)" — 사업부 박스 바로 위/옆
        if re.match(r"^\(.*등급.*건", txt):
            d = _nearest_division_box(slide, sh)
            if d:
                W.set_shape_text(sh, clean_cell(ws, DIV_CELLS[d]["grade"]))
                sh.text_frame.word_wrap = False   # 글상자가 안 커져 2줄 되는 문제 → 한 줄 유지
                n_grade += 1
            continue
        # 사업부 서술 박스: 사업부명으로 시작
        d = which_division(txt)
        if d and txt.startswith(d):
            W.set_shape_text(sh, clean_cell(ws, DIV_CELLS[d]["narr"]))
            n_div += 1
            continue
        # 표지 제목 "0. 7월 현장 점검 결과" → 이번달로
        if "현장 점검 결과" in txt or re.search(r"\d+월", txt) and "Page" not in txt and len(txt) < 25:
            W.set_shape_text(sh, re.sub(r"\d+\s*월", f"{base_date.month}월", txt))
    W.log(f"  [슬라이드0] 표지 갱신: 사업부박스 {n_div} · 등급내역 {n_grade} · 전사누적 {n_total}")


def _nearest_division_box(slide, small):
    """소박스(등급내역) 위치에서 세로로 가장 가까운 사업부 서술박스의 사업부명을 찾는다."""
    best, bestd = None, None
    for sh in slide.shapes:
        if not sh.has_text_frame or sh is small:
            continue
        d = which_division(sh.text_frame.text.strip())
        if not d or not sh.text_frame.text.strip().startswith(d):
            continue
        dist = abs((sh.top or 0) - (small.top or 0))
        if best is None or dist < best:
            best, bestd = dist, d
    return bestd


# ─────────────────────────────────────────────────────────────────────────────
# 별첨1 (사업부 대형 추이차트 슬라이드): 차트 + 서술문구 + 전사누적 + 날짜라벨
# ─────────────────────────────────────────────────────────────────────────────
def align_attach1_bands(slide, n_tpl, n_new):
    """별첨1 대형차트: 개월 수가 n_tpl→n_new 로 늘면, 회색 연도밴드(큰 직사각형)와
    '’24/’25/’26년' 라벨을 2024년 1월(왼쪽 앵커) 기준으로 개월 비율만큼 좌측 압축한다.
    (차트가 카테고리를 동일 폭에 재분배하는 것과 같은 변환 → 밴드 의미와 무관하게 정합)."""
    if n_tpl < 2 or n_new < 2 or n_tpl == n_new:
        return 0
    yr = re.compile(r"[’'’]?(\d{2}|\d{4})\s*년$")
    anchor = None                                  # 2024년 1월 = '’24년' 라벨 left
    for sh in slide.shapes:
        if sh.has_text_frame:
            m = yr.fullmatch(sh.text_frame.text.strip())
            if m and (m.group(1) in ("24", "2024")):
                anchor = sh.left
                break
    if anchor is None:
        return 0
    r = n_tpl / float(n_new)

    def remap(x):
        return int(round(anchor + (x - anchor) * r))

    moved = 0
    for sh in slide.shapes:
        if sh.shape_type == 1 and (sh.height or 0) > 914400:      # 큰 직사각형 = 연도밴드
            sh.width = int(round(sh.width * r))
            sh.left = remap(sh.left)
            moved += 1
        elif sh.has_text_frame and yr.fullmatch(sh.text_frame.text.strip()):
            sh.left = remap(sh.left)
            moved += 1
    return moved


def fill_attach1(slide, ws, divs0, categories, base_date):
    # 슬라이드가 담당하는 사업부 = 제목 '[별첨1] {사업부} …'
    div = None
    for sh in slide.shapes:
        if sh.has_text_frame and "[별첨1]" in sh.text_frame.text:
            div = which_division(sh.text_frame.text)
            break
    if not div:
        return
    # 대형 차트(제목 무제) 갱신 — 교체 전 템플릿 개월 수(n_tpl) 확보
    n_tpl = None
    for sh in slide.shapes:
        if sh.has_chart:
            n_tpl = len(list(sh.chart.plots[0].categories))
            if div in divs0:
                W.replace_chart(sh, categories, divs0[div])
    if n_tpl:
        m = align_attach1_bands(slide, n_tpl, len(categories))
        W.log(f"  [별첨1] {div} 연도밴드/라벨 정렬 ({m}개)")
    # 서술문구 / 전사누적 / 날짜 라벨
    for sh in slide.shapes:
        if not sh.has_text_frame:
            continue
        t = sh.text_frame.text.strip()
        if not t or "[별첨1]" in t:
            continue
        if "전사 누적" in t:
            W.set_shape_text(sh, clean_cell(ws, TOTAL_CELLS["narr"]))
        elif t.startswith(div) or (div in t and "기준" in t and "등급장애" in t):
            # "M/D 기준 {사업부} 등급장애는 …" — 문구는 요약셀, 날짜는 기준일
            W.set_shape_text(sh, W.sub_date(clean_cell(ws, DIV_CELLS[div]["narr"]), base_date))
        elif "기준)" in t:                       # "(M/D {사업부} 기준)" 우측 라벨
            W.set_shape_text(sh, W.sub_date(t, base_date))
    W.log(f"  [별첨1] {div} 차트/서술 갱신")


# ─────────────────────────────────────────────────────────────────────────────
# 날짜 표기 일괄 갱신 ('M/D 기준', 'N월D일 기준' 등)
# ─────────────────────────────────────────────────────────────────────────────
def refresh_dates(prs, base_date):
    n = 0
    pat_kor = re.compile(r"\d{1,2}\s*월\s*\d{1,2}\s*일\s*기준")   # "5월27일 기준"
    for slide in prs.slides:
        for sh in slide.shapes:
            if not sh.has_text_frame:
                continue
            t = sh.text_frame.text
            if pat_kor.search(t):
                W.set_shape_text(sh, pat_kor.sub(f"{base_date.month}월{base_date.day}일 기준", t))
                n += 1
    if n:
        W.log(f"  [전체] 'N월D일 기준' 날짜 라벨 {n}개 갱신")


# ─────────────────────────────────────────────────────────────────────────────
# 별첨2 상세표: 동료 주간 결과물(YYYYMMDD_장애공유.pptx)의 '장애 목록 상세' /
#              '등급 협의 중 장애' 에서 해당 사업부 행만 가져와 채운다.
#   주간 표(9열): 장애명·발생일·시간·고객·내용·원인·[사업부]·담당·팀
#   월간 표(8열): 장애명·발생일·시간·고객·내용·원인·담당·팀  (사업부 열만 제외)
# ─────────────────────────────────────────────────────────────────────────────
WEEKLY_DETAIL_KW = "장애 목록 상세"
WEEKLY_HYEOBUI_KW = "등급 협의 중 장애"


def _clean(s):
    """셀 텍스트 정제: 세로탭(소프트 줄바꿈)→개행, 그 외 제어문자 제거 (XML 호환)."""
    s = (s or "").replace("\x0b", "\n").replace("\r", "\n")
    return re.sub(r"[\x00-\x08\x0c\x0e-\x1f]", "", s).strip()


def _weekly_rows(table):
    """주간 표 데이터행(2행부터) → [{div, vals[8]}] (사업부 열 제외한 8값)."""
    rows = []
    for ri in range(2, len(table.rows)):
        cells = [_clean(c.text) for c in table.rows[ri].cells]
        if len(cells) < 9 or not cells[0]:
            continue
        rows.append({
            "div": cells[6],
            "vals": [cells[0], cells[1], cells[2], cells[3],
                     cells[4], cells[5], cells[7], cells[8]],
        })
    return rows


def read_weekly(weekly_prs):
    """주간 결과물에서 상세목록/협의중 행을 모아 반환 (슬라이드 순서 유지)."""
    detail, hyeobui = [], []
    for slide in weekly_prs.slides:
        titles = " ".join(sh.text_frame.text for sh in slide.shapes if sh.has_text_frame)
        is_detail = WEEKLY_DETAIL_KW in titles
        is_hyeobui = WEEKLY_HYEOBUI_KW in titles
        if not (is_detail or is_hyeobui):
            continue
        for sh in slide.shapes:
            if not sh.has_table:
                continue
            (detail if is_detail else hyeobui).extend(_weekly_rows(sh.table))
    W.log(f"  [주간소스] 상세목록 {len(detail)}건 · 협의중 {len(hyeobui)}건 로드")
    return detail, hyeobui


def _is_hyeobui_table(table):
    """월간 별첨2 표 분류: col5 헤더에 '예상등급'이 있으면 협의중 표."""
    try:
        return "예상등급" in table.rows[0].cells[5].text
    except Exception:
        return False


def _attach2_slides(prs, div):
    """해당 사업부 [별첨2] 슬라이드의 (확정표 리스트, 협의중표 리스트) — 슬라이드 순서."""
    conf, hyb, last_slide = [], [], None
    for slide in prs.slides:
        titles = " ".join(sh.text_frame.text for sh in slide.shapes if sh.has_text_frame)
        if "[별첨2]" not in titles or div not in titles:
            continue
        last_slide = slide
        for sh in slide.shapes:
            if not sh.has_table:
                continue
            (hyb if _is_hyeobui_table(sh.table) else conf).append((slide, sh))
    return conf, hyb, last_slide


def _fill_tr(tr, vals):
    tcs = tr.findall(qn("a:tc"))
    for ci, v in enumerate(vals):
        if ci < len(tcs):
            W._set_tc_text(tcs[ci], v)


def _move_after(prs, moved, ref):
    """moved 슬라이드를 ref 슬라이드 바로 뒤로 이동(sldIdLst 재배치)."""
    lst = prs.slides._sldIdLst

    def sldid(sl):
        for sid in lst:
            rId = sid.get(qn("r:id"))
            if prs.part.rels[rId].target_part is sl.part:
                return sid
        return None

    m, r = sldid(moved), sldid(ref)
    if m is not None and r is not None:
        lst.remove(m)
        r.addnext(m)


def _strip_other_tables(slide, keep_tsh):
    """복제 슬라이드에서 대상 표(keep_tsh) 외의 다른 표는 제거(협의중/확정 혼입 방지).
    python-pptx가 shape 접근마다 새 래퍼를 만들므로 내부 XML 요소로 비교한다."""
    keep = keep_tsh._element
    for sh in list(slide.shapes):
        if sh.has_table and sh._element is not keep:
            sh._element.getparent().remove(sh._element)


def _distribute(prs, tables, rows):
    """확정/협의중 행(rows=[[8값]])을 표들에 용량 기준으로 분배.
    초과 시 마지막 슬라이드를 복제해 해당 사업부 자리 바로 뒤에 순서대로 삽입."""
    if not tables:
        return 0
    def data_trs(tsh):
        return tsh.table._tbl.findall(qn("a:tr"))
    style_tr = copy.deepcopy(data_trs(tables[0][1])[2])          # 서식 소스(첫 표 데이터행)
    budget = max((sum(W._tr_est_height(tr) for tr in data_trs(t)[2:]) or 1)
                 for _, t in tables)
    built = []
    for v in rows:
        tr = copy.deepcopy(style_tr)
        _fill_tr(tr, v)
        built.append(tr)
    # 용량 기준 페이지 그룹핑
    groups, cur, used = [], [], 0
    for tr in built:
        h = W._tr_est_height(tr)
        if cur and used + h > budget:
            groups.append(cur); cur, used = [], 0
        cur.append(tr); used += h
    if cur:
        groups.append(cur)
    if not groups:
        groups = [[]]
    # 페이지 부족 시: 마지막 슬라이드 복제 → 직전 페이지 바로 뒤로 이동 → 대상 표만 남김
    anchor = tables[-1][0]
    while len(tables) < len(groups):
        s, tsh = tables[-1]
        ns, ntsh = W._duplicate_detail_slide(prs, s)
        _strip_other_tables(ns, ntsh)
        _move_after(prs, ns, anchor)
        anchor = ns
        tables.append((ns, ntsh))
    # 배치
    for i, (s, tsh) in enumerate(tables):
        W._clear_data_rows(tsh.table)
        if i < len(groups):
            tbl = tsh.table._tbl
            for tr in groups[i]:
                tbl.append(tr)
    return len(built)


def _table_data_recs(table):
    """비교표 데이터행 → [{date,cust,name}] (동일장애 매칭용)."""
    recs = []
    for ri in range(2, len(table.rows)):
        cells = [_clean(c.text) for c in table.rows[ri].cells]
        if not cells or not cells[0]:
            continue
        recs.append({"name": cells[0],
                     "date": W._norm_date(cells[1]) if len(cells) > 1 else None,
                     "cust": cells[3] if len(cells) > 3 else ""})
    return recs


def read_compare(compare_prs):
    """비교 파일(전월 월간보고 등)에서 사업부별 등급확정 항목 레코드를 만든다."""
    out = {}
    for div in DIVISIONS:
        conf, _, _ = _attach2_slides(compare_prs, div)
        recs = []
        for _, tsh in conf:
            recs += _table_data_recs(tsh.table)
        out[div] = recs
    W.log("  [비교] " + " · ".join(f"{d[:4]} {len(out[d])}건" for d in DIVISIONS))
    return out


def _update_new_count(prs, div, n):
    """해당 사업부 별첨2 슬라이드의 '전월 보고 이후 … 신규 장애는 N건입니다' 문구 갱신."""
    phrase = "없습니다" if n == 0 else f"{n}건입니다"
    for slide in prs.slides:
        titles = " ".join(sh.text_frame.text for sh in slide.shapes if sh.has_text_frame)
        if "[별첨2]" not in titles or div not in titles:
            continue
        for sh in slide.shapes:
            if sh.has_text_frame and "확정된 신규 장애는" in sh.text_frame.text:
                new = re.sub(r"신규 장애는\s*(?:없습니다|\d+\s*건입니다)\.?",
                             f"신규 장애는 {phrase}.", sh.text_frame.text)
                W.set_shape_text(sh, new)


def fill_attach2(prs, detail_rows, hyeobui_rows, compare_keys=None):
    """사업부별로 별첨2(등급확정 상세 + 협의중)를 주간 소스에서 채운다.
    compare_keys 가 주어지면 '전월 대비 신규 확정 N건'을 항목 비교로 산정해 문구 갱신."""
    for div in DIVISIONS:
        conf_tabs, hyb_tabs, _ = _attach2_slides(prs, div)
        d_rows = [r["vals"] for r in detail_rows if which_division(r["div"]) == div]
        h_rows = [r["vals"] for r in hyeobui_rows if which_division(r["div"]) == div]
        nc = _distribute(prs, conf_tabs, d_rows)
        nh = _distribute(prs, hyb_tabs, h_rows)
        # 협의중 건수 문구 갱신
        for slide in prs.slides:
            titles = " ".join(sh.text_frame.text for sh in slide.shapes if sh.has_text_frame)
            if "[별첨2]" in titles and div in titles:
                for sh in slide.shapes:
                    if sh.has_text_frame and "협의 중인 장애는 총" in sh.text_frame.text:
                        W.set_shape_text(sh, re.sub(r"총\s*\d+\s*건",
                                                    f"총 {len(h_rows)}건", sh.text_frame.text))
        # 전월 대비 신규 확정 N건 (비교 파일에 '동일 장애'로 매칭되지 않는 항목 수)
        msg_new = ""
        if compare_keys is not None:
            prev = compare_keys.get(div, [])
            new_names = [v[0] for v in d_rows
                         if not any(W._same_incident(W._norm_date(v[1]), v[3], v[0], r)
                                    for r in prev)]
            n_new = len(new_names)
            _update_new_count(prs, div, n_new)
            msg_new = f" · 신규확정 {n_new}건(전월 {len(prev)}→현재 {len(d_rows)})"
            for nm in new_names:
                W.log(f"       + 신규: {nm[:40]}")
        W.log(f"  [별첨2] {div}: 등급확정 {nc}건 · 협의중 {nh}건{msg_new}")


def renumber_attach2(prs):
    """사업부별 [별첨2] 슬라이드 제목의 '(i/N)' 페이지 라벨을 실제 슬라이드 수로 재넘버링."""
    order = []
    for slide in prs.slides:
        titles = " ".join(sh.text_frame.text for sh in slide.shapes if sh.has_text_frame)
        if "[별첨2]" in titles:
            order.append((which_division(titles), slide))
    perdiv = {}
    for div, slide in order:
        perdiv.setdefault(div, []).append(slide)
    for div, slides in perdiv.items():
        total = len(slides)
        for i, slide in enumerate(slides, 1):
            for sh in slide.shapes:
                if not (sh.has_text_frame and "[별첨2]" in sh.text_frame.text):
                    continue
                base = re.sub(r"\s*\(\s*\d+\s*/\s*\d+\s*\)\s*$", "",
                              sh.text_frame.text.rstrip())
                W.set_shape_text(sh, base + (f" ({i}/{total})" if total > 1 else ""))
                break
    if any(len(v) > 1 for v in perdiv.values()):
        W.log("  [별첨2] 페이지 라벨 (i/N) 재넘버링 완료")


# ─────────────────────────────────────────────────────────────────────────────
# 메인
# ─────────────────────────────────────────────────────────────────────────────
def parse_args():
    p = argparse.ArgumentParser(description="장애 월간보고 PPT 생성")
    p.add_argument("--xlsx", default=None, help="갱신된 엑셀 경로(미지정 시 자동 탐색)")
    p.add_argument("--prev", required=True, help="지난달 장애현황 PPT (템플릿 겸 비교용, 필수)")
    p.add_argument("--weekly", default=None,
                   help="별첨2 소스 = 동료 주간 결과물 YYYYMMDD_장애공유.pptx (미지정 시 별첨2 미갱신)")
    p.add_argument("--compare", default=None,
                   help="'전월 대비 신규 확정 N건' 비교 기준 파일 (전월 월간보고 등). 미지정 시 --prev 사용")
    p.add_argument("--date", default=None, help="기준일 (YYYY-MM-DD, 기본: 오늘)")
    p.add_argument("--out", default=None, help="결과 PPT 경로(미지정 시 지난달 PPT 폴더)")
    p.add_argument("--outdir", default=None, help="결과 폴더(미지정 시 지난달 PPT 폴더)")
    return p.parse_args()


def main():
    args = parse_args()
    here = os.path.dirname(os.path.abspath(__file__))

    if args.date:
        base_date = dt.datetime.strptime(args.date, "%Y-%m-%d")
    else:
        t = dt.date.today()
        base_date = dt.datetime(t.year, t.month, t.day)
    by, bm = base_date.year, base_date.month

    xlsx = W.resolve_xlsx(args, base_date, here)
    if not xlsx or not os.path.exists(xlsx):
        W.log("[오류] 갱신된 엑셀을 찾을 수 없습니다. update_excel.py 실행 또는 --xlsx 지정 필요.")
        sys.exit(1)
    if not os.path.exists(args.prev):
        W.log(f"[오류] 지난달 PPT를 찾을 수 없습니다: {args.prev}")
        sys.exit(1)

    if args.out:
        out = os.path.abspath(args.out)
    else:
        outdir = args.outdir or os.path.dirname(os.path.abspath(args.prev))
        os.makedirs(outdir, exist_ok=True)
        out = os.path.join(outdir, f"{base_date:%y}년 {bm}월 장애현황_기준일({base_date:%Y%m%d}).pptx")

    W.log("=" * 70)
    W.log("장애 월간보고 - PPT 생성  (v0.2: 차트 + 서술문구 + 별첨2 상세표)")
    W.log("=" * 70)
    W.log(f"  엑셀     : {xlsx}")
    W.log(f"  템플릿   : {args.prev}  (지난달 PPT)")
    W.log(f"  기준일   : {base_date:%Y-%m-%d}  (해당월 {by}-{bm:02d} 진행분, 최근 3개년 {by-2}~{by})")
    W.log(f"  결과 PPT : {out}")
    W.log("-" * 70)

    wb = openpyxl.load_workbook(xlsx, data_only=True)
    ws_sum = wb["요약"]

    months = W.month_range(by - 2, 1, by, bm)
    categories = [f"{m}월" for (_, m) in months]
    W.log(f"  차트 구간: {by-2}-01 ~ {by}-{bm:02d}  ({len(months)}개월)")

    def series_for(cfg_key):
        cfg = W.SRC[cfg_key]
        ws = wb[cfg["sheet"]]
        divs = {name: W.read_series(ws, row, cfg["div_start_col"], months)
                for name, row in cfg["divisions"].items()}
        return divs

    divs0 = series_for("사업부별")             # 슬라이드1 · 별첨1
    divs1 = series_for("사업부별_계열사별")     # 슬라이드2

    prs = Presentation(args.prev)

    # 슬라이드 0: 표지
    fill_cover(prs.slides[0], ws_sum, base_date)

    # 슬라이드 1,2: 사업부 4개 추이차트
    update_div_slide(prs.slides[1], divs0, categories, "슬라이드1")
    n1 = W.align_year_labels_grid(prs.slides[1], len(months), by)
    W.log(f"  [슬라이드1] 연도 라벨 1월 정렬 ({n1}개)")
    update_div_slide(prs.slides[2], divs1, categories, "슬라이드2")
    n2 = W.align_year_labels_grid(prs.slides[2], len(months), by)
    W.log(f"  [슬라이드2] 연도 라벨 1월 정렬 ({n2}개)")

    # 별첨1: 각 사업부 대형 추이차트 슬라이드
    for slide in prs.slides:
        if any(sh.has_text_frame and "[별첨1]" in sh.text_frame.text for sh in slide.shapes):
            fill_attach1(slide, ws_sum, divs0, categories, base_date)

    # 별첨2: 주간 결과물에서 사업부별 상세목록/협의중 채우기
    if args.weekly:
        if not os.path.exists(args.weekly):
            W.log(f"  [경고] 주간 결과물을 찾을 수 없어 별첨2 미갱신: {args.weekly}")
        else:
            detail_rows, hyeobui_rows = read_weekly(Presentation(args.weekly))
            # 전월 대비 신규확정 비교 파일 (미지정 시 --prev)
            compare_path = args.compare or args.prev
            compare_keys = None
            if compare_path and os.path.exists(compare_path):
                compare_keys = read_compare(Presentation(compare_path))
                W.log(f"  [비교] 기준 파일: {os.path.basename(compare_path)}")
            fill_attach2(prs, detail_rows, hyeobui_rows, compare_keys)
            renumber_attach2(prs)               # 페이지 라벨 (i/N) 재넘버링
            W.normalize_table_borders(prs)      # 행 증감으로 흐트러진 표 테두리 정규화
    else:
        W.log("  [별첨2] --weekly 미지정 → 상세표는 템플릿 그대로 유지")

    # 날짜 라벨 일괄 갱신
    refresh_dates(prs, base_date)

    prs.save(out)
    W.log("-" * 70)
    W.log(f"  저장 완료: {out}")
    W.log("=" * 70)
    W.log("완료.")


if __name__ == "__main__":
    main()
