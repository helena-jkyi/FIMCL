import calendar as _cal
import csv
import datetime
import html
import io
import json
import os
import re
import ssl
import urllib.parse
import urllib.request
import uuid

import streamlit as st

st.set_page_config(page_title="FIMCL", page_icon="🔬", layout="wide")

# ══════════════════════════════════════════════════
#  ✏️  랩 기본 정보
# ══════════════════════════════════════════════════
LAB = {
    "name": "FIMCL",
    "full_name": "박선아 교수님 연구실",
    "tagline": "주문/행정 관리 페이지 입니다",
    "intro": "주문/행정 현황을 파악할 수 있습니다",
    "affiliation": "KAIST 화학과",
    "location": "대전광역시",
    "email": "jaekyung12@kaist.ac.kr",
}
RESEARCH = [
    {"icon": "🔬", "title": "연구 분야 1", "desc": "이 연구 방향을 한두 문장으로 소개하세요."},
    {"icon": "⚗️", "title": "연구 분야 2", "desc": "두 번째 연구 주제 설명을 여기에 적어주세요."},
    {"icon": "📊", "title": "연구 분야 3", "desc": "세 번째 연구 주제 설명을 여기에 적어주세요."},
]
MEMBERS = [
    {"name": "지도교수 성함", "role": "Principal Investigator", "detail": "연구 관심 분야를 적어주세요."},
    {"name": "구성원 이름", "role": "박사과정 · Ph.D. Student", "detail": "담당 연구 주제를 적어주세요."},
    {"name": "구성원 이름", "role": "석사과정 · M.S. Student", "detail": "담당 연구 주제를 적어주세요."},
]

# ══════════════════════════════════════════════════
#  🌴 휴가 구글 시트 연결  ← 여기에 링크만 붙여넣으면 됩니다
# ══════════════════════════════════════════════════
# ▸ 시트 [공유] → '링크 복사' 한 주소를 그대로 넣으세요. gid 가 없어도 괜찮아요.
#     예) "https://docs.google.com/spreadsheets/d/1AbCdEf...xyz/edit?usp=sharing"
# ▸ 공유 설정은 '링크가 있는 모든 사용자: 뷰어' (또는 서비스 계정 이메일에 시트 공유).
# ▸ .xlsx 파일을 드라이브에 그대로 올린 경우엔 [파일] → 'Google 스프레드시트로 저장' 후
#   새로 생긴 시트의 링크를 쓰세요. (엑셀 원본 상태로는 읽히지 않아요)
# ▸ 앱은 '휴가 입력' 탭만 읽고(시트는 수정하지 않음), 월별 요약·달력은 앱에서 계산합니다.
VAC_SHEET_LINK = "https://docs.google.com/spreadsheets/d/18f6OeOsjEQE7Rgj7QAKRsh_JK9KXAaNO/edit?gid=1847544715#gid=1847544715"
VAC_TAB = "휴가 입력"        # 휴가 기록이 있는 탭 이름
VAC_HEADER_ROWS = 3          # 그 탭에서 머리글(이름/사용일/…)이 있는 행 번호
# 월별 요약에 항상 보일 구성원 (시트 '월별 요약' 순서). 기록에만 있는 이름은 자동으로 뒤에 붙어요.
VAC_ROSTER = ["박경철", "이재경", "박근찬", "김문정", "전재민", "박지원", "정유리", "김강민", "문상원",
              "박민서", "하승연", "조수용", "박성현", "김수린", "김다연", "김주현", "이건우"]


# ══════════════════════════════════════════════════
#  공강표 데이터
# ══════════════════════════════════════════════════
DAYS = ["월", "화", "수", "목", "금"]
TIMES = ["09:00–09:30", "09:30–10:00", "10:00–10:30", "10:30–11:00", "11:00–11:30",
         "11:30–12:00", "12:00–12:30", "12:30–13:00", "13:00–13:30", "13:30–14:00",
         "14:00–14:30", "14:30–15:00", "15:00–15:30", "15:30–16:00", "16:00–16:30",
         "16:30–17:00", "17:00–17:30", "17:30–18:00"]
SLOTS = [(9 * 60 + 30 * i, 9 * 60 + 30 * (i + 1)) for i in range(18)]

GROUPS = {
    "2d조":     ["수린", "다연", "유리", "상원", "주현", "승연", "지원", "근찬"],
    "골드조":   ["문정", "성현", "민서", "경철"],
    "mofcof조": ["강민", "수용", "재민", "재경", "건우"],
}

# 수업 듣는 5명만 입력 (나머지 12명은 자동으로 '평일 항상 가능')
# 형식: (요일, 시작, 끝, 과목/일정)  — 요일은 "월·수", "화·목", "금" 처럼
CLASSES = {
    "수린": [("화·목", "09:00", "10:30", "고체의구조및결함"),
             ("월·수", "13:00", "14:30", "고체화학개론"),
             ("화·목", "14:30", "16:00", "전기화학분석"),
             ("수", "16:30", "18:00", "세미나(석사)"),
             ("금", "16:00", "18:00", "리더십강좌")],
    "다연": [("월·수", "09:00", "10:30", "분자분광학개론"),
             ("월·수", "13:00", "14:30", "고체화학개론"),
             ("화·목", "14:30", "16:00", "전기화학분석"),
             ("금", "13:00", "16:00", "일반화학실험 조교"),
             ("월", "16:00", "17:30", "화학교육실습"),
             ("수", "16:30", "18:00", "세미나(석사)")],
    "주현": [("월·수", "13:00", "14:30", "고체화학개론"),
             ("월·수", "14:30", "16:00", "인공지능 화학"),
             ("화·목", "14:30", "16:00", "소재분석"),
             ("목", "09:00", "12:00", "일반화학실험 조교"),
             ("월", "16:00", "17:30", "화학교육실습"),
             ("수", "16:30", "18:00", "세미나(석사)")],
    "유리": [("화·목", "14:00", "15:30", "고급 나노 공정·생산 (온라인)"),
             ("목", "16:00", "16:50", "초청세미나 (온라인)")],
    "건우": [("화·목", "10:30", "12:00", "Scientific Writing"),
             ("화·목", "14:30", "16:00", "전기화학분석"),
             ("수", "16:30", "18:00", "세미나(석사)")],
}

# 조별 '다 같이 되는 시간' 표 — 엑셀 파일 값 그대로 (숫자 = 참석 가능 인원)
GRIDS = {
    "전체 랩": {"total": 17, "rows": [
        [16, 16, 16, 15, 17], [16, 16, 16, 15, 17], [16, 16, 16, 15, 17], [17, 16, 17, 15, 17],
        [17, 16, 17, 15, 17], [17, 16, 17, 15, 17], [17, 17, 17, 17, 17], [17, 17, 17, 17, 17],
        [14, 17, 14, 17, 16], [14, 17, 14, 17, 16], [14, 16, 14, 16, 16], [16, 12, 16, 13, 16],
        [16, 12, 16, 13, 16], [16, 13, 16, 14, 16], [15, 17, 17, 16, 16], [15, 17, 13, 16, 16],
        [15, 17, 13, 17, 16], [17, 17, 13, 17, 16]]},
    "2d조": {"total": 8, "rows": [
        [7, 7, 7, 6, 8], [7, 7, 7, 6, 8], [7, 7, 7, 6, 8], [8, 8, 8, 7, 8],
        [8, 8, 8, 7, 8], [8, 8, 8, 7, 8], [8, 8, 8, 8, 8], [8, 8, 8, 8, 8],
        [5, 8, 5, 8, 7], [5, 8, 5, 8, 7], [5, 7, 5, 7, 7], [7, 4, 7, 5, 7],
        [7, 4, 7, 5, 7], [7, 5, 7, 6, 7], [6, 8, 8, 7, 7], [6, 8, 5, 7, 7],
        [6, 8, 5, 8, 7], [8, 8, 5, 8, 7]]},
    "골드조": {"total": 4, "rows": [[4, 4, 4, 4, 4] for _ in range(18)]},
    "mofcof조": {"total": 5, "rows": [
        [5, 5, 5, 5, 5], [5, 5, 5, 5, 5], [5, 5, 5, 5, 5], [5, 4, 5, 4, 5],
        [5, 4, 5, 4, 5], [5, 4, 5, 4, 5], [5, 5, 5, 5, 5], [5, 5, 5, 5, 5],
        [5, 5, 5, 5, 5], [5, 5, 5, 5, 5], [5, 5, 5, 5, 5], [5, 4, 5, 4, 5],
        [5, 4, 5, 4, 5], [5, 4, 5, 4, 5], [5, 5, 5, 5, 5], [5, 5, 4, 5, 5],
        [5, 5, 4, 5, 5], [5, 5, 4, 5, 5]]},
}
FULL_TIMES = {  # 전원 가능 시간대
    "전체 랩": {"월": "10:30–13:00, 17:30–18:00", "화": "12:00–14:00, 16:00–18:00",
              "수": "10:30–13:00, 16:00–16:30", "목": "12:00–14:00, 17:00–18:00", "금": "09:00–13:00"},
    "2d조": {"월": "10:30–13:00, 17:30–18:00", "화": "10:30–14:00, 16:00–18:00",
            "수": "10:30–13:00, 16:00–16:30", "목": "12:00–13:00, 17:00–18:00", "금": "09:00–13:00"},
    "골드조": {"월": "09:00–18:00", "화": "09:00–18:00", "수": "09:00–18:00", "목": "09:00–18:00", "금": "09:00–18:00"},
    "mofcof조": {"월": "09:00–18:00", "화": "09:00–10:30, 12:00–14:30, 16:00–18:00",
                "수": "09:00–16:30", "목": "09:00–10:30, 12:00–14:30, 16:00–18:00", "금": "09:00–18:00"},
}
RECS = {
    "전체 랩": ["금 09:00–13:00 (240분)", "월 10:30–13:00 (150분)", "수 10:30–13:00 (150분)", "화 16:00–18:00 (120분)"],
    "2d조": ["금 09:00–13:00 (240분)", "화 10:30–14:00 (210분)", "월 10:30–13:00 (150분)", "수 10:30–13:00 (150분)"],
    "골드조": ["화 09:00–18:00 (540분)", "월 09:00–18:00 (540분)", "수 09:00–18:00 (540분)", "목 09:00–18:00 (540분)"],
    "mofcof조": ["월 09:00–18:00 (540분)", "금 09:00–18:00 (540분)", "수 09:00–16:30 (450분)", "화 12:00–14:30 (150분)"],
}

# ══════════════════════════════════════════════════
#  계산 (개별 그리드용)
# ══════════════════════════════════════════════════
ALL_MEMBERS = [m for g in GROUPS.values() for m in g]


def group_of(n):
    for g, ms in GROUPS.items():
        if n in ms:
            return g
    return ""


def _to_min(t):
    h, m = t.split(":")
    return int(h) * 60 + int(m)


_busy = {}
for _name, _lst in CLASSES.items():
    _iv = []
    for _dt, _s, _e, _c in _lst:
        for _d in _dt.split("·"):
            _iv.append((_d, _to_min(_s), _to_min(_e), _c))
    _busy[_name] = _iv


def busy_course(name, day, i):
    ss, se = SLOTS[i]
    for d, s, e, c in _busy.get(name, []):
        if d == day and s < se and e > ss:
            return c
    return None


# ══════════════════════════════════════════════════
#  HTML 렌더러
# ══════════════════════════════════════════════════
def render_group(key):
    g = GRIDS[key]
    T = g["total"]
    head = "<tr><th class='time'>시간</th>" + "".join(f"<th>{d}</th>" for d in DAYS) + "</tr>"
    body = ""
    for i, tlabel in enumerate(TIMES):
        cells = ""
        for j in range(5):
            c = g["rows"][i][j]
            lvl = min(max(T - c, 0), 4)
            cells += f"<td class='lvl{lvl}'>{c}</td>"
        body += f"<tr><td class='time'>{tlabel}</td>{cells}</tr>"
    table = f"<table class='sched'><thead>{head}</thead><tbody>{body}</tbody></table>"

    legend = ("<div class='legend'><span class='lbl'>참석 가능 인원</span>"
              "<span class='lg lvl0'>전원</span><span class='lg lvl1'>-1</span>"
              "<span class='lg lvl2'>-2</span><span class='lg lvl3'>-3</span>"
              "<span class='lg lvl4'>그이하</span></div>")

    ft = FULL_TIMES[key]
    ftrows = "".join(f"<div class='ftrow'><b>{d}</b> {ft[d]}</div>" for d in DAYS)
    recs = "".join(f"<li>{r}</li>" for r in RECS[key])
    summary = (f"<div class='summary'>"
               f"<div class='sumcol'><h4>전원({T}명) 가능 시간대</h4>{ftrows}</div>"
               f"<div class='sumcol'><h4>추천 (긴 순)</h4><ol>{recs}</ol></div></div>")
    caption = f"<div class='ind-note'>총 {T}명 · 숫자는 참석 가능 인원, 초록이 전원 가능 시간대예요.</div>"
    return caption + table + legend + summary


def render_person(name):
    grp = group_of(name)
    has = name in CLASSES
    head = "<tr><th class='time'>시간</th>" + "".join(f"<th>{d}</th>" for d in DAYS) + "</tr>"
    body = ""
    for i, tlabel in enumerate(TIMES):
        cells = ""
        for d in DAYS:
            cur = busy_course(name, d, i)
            above = busy_course(name, d, i - 1) if i > 0 else "\x00"
            top = cur != above
            if cur is None:
                cells += f"<td class='grid-free'>{'가능' if top else ''}</td>"
            else:
                cells += f"<td class='grid-busy'>{cur if top else ''}</td>"
        body += f"<tr><td class='time'>{tlabel}</td>{cells}</tr>"
    table = f"<table class='sched'><thead>{head}</thead><tbody>{body}</tbody></table>"

    if has:
        rows = "".join(f"<div class='clrow'><span class='cld'>{c[0]}</span>"
                       f"<span class='clt'>{c[1]}–{c[2]}</span><span class='cln'>{c[3]}</span></div>"
                       for c in CLASSES[name])
        note = f"<div class='clist'><h4>{name} · {grp} · 수업/일정</h4>{rows}</div>"
    else:
        note = f"<div class='freeall'>{name} · {grp} · 수업 없음 → 평일 09:00–18:00 언제나 가능</div>"
    return table + note


# ══════════════════════════════════════════════════
#  스타일
# ══════════════════════════════════════════════════
CSS = """
<style>
@import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.css');
:root{--bg:#F8FAFC;--surface:#FFFFFF;--ink:#0F172A;--muted:#64748B;--line:#E7ECF2;
--accent:#0D9488;--accent-dark:#0F766E;--accent-soft:#E6F4F1;}

[data-testid="stHeader"],[data-testid="stToolbar"],[data-testid="stDecoration"]{display:none;}
#MainMenu,footer{visibility:hidden;}
.block-container,.stMainBlockContainer{max-width:1080px !important;padding:0 24px 48px !important;margin:0 auto !important;}
[data-testid="stAppViewContainer"]{background:var(--bg);}
html,body,[class*="css"]{font-family:'Pretendard',-apple-system,BlinkMacSystemFont,'Malgun Gothic',sans-serif;}
html{scroll-behavior:smooth;}
*{box-sizing:border-box;}

/* 내비 */
.nav{position:sticky;top:0;z-index:50;background:rgba(248,250,252,0.9);backdrop-filter:blur(10px);
border-bottom:1px solid var(--line);margin:0 -24px;padding:0 24px;}
.nav-in{display:flex;align-items:center;justify-content:space-between;height:64px;}
.brand{font-weight:800;font-size:1.15rem;letter-spacing:-0.02em;color:var(--ink);}
.brand .dot{color:var(--accent);}
.links{display:flex;gap:24px;}
.links a{color:var(--muted);text-decoration:none;font-weight:500;font-size:0.95rem;}
.links a:hover{color:var(--ink);}

/* 히어로 */
.hero{padding:72px 0 56px;}
.eyebrow{display:inline-block;margin-bottom:20px;color:var(--accent-dark);background:var(--accent-soft);
padding:6px 15px;border-radius:999px;font-weight:600;font-size:0.9rem;}
.hero-name{font-size:clamp(3rem,8vw,4.6rem);font-weight:800;letter-spacing:-0.04em;line-height:1.05;margin:0 0 6px;color:var(--ink);}
.hero-role{font-size:clamp(1.1rem,2.5vw,1.45rem);font-weight:600;color:var(--accent);margin:0 0 22px;}
.lead{font-size:clamp(1.02rem,2vw,1.2rem);color:var(--muted);max-width:600px;margin:0 0 32px;}
.cta-row{display:flex;gap:12px;flex-wrap:wrap;}
.btn{display:inline-flex;align-items:center;text-decoration:none;font-weight:600;font-size:1rem;padding:12px 24px;border-radius:12px;}
.btn-primary{background:var(--accent);color:#fff;}
.btn-primary:hover{background:var(--accent-dark);}
.btn-ghost{background:transparent;color:var(--ink);border:1px solid var(--line);}
.btn-ghost:hover{border-color:var(--ink);}

/* 섹션 */
.section{padding:56px 0;}
.section.sched-head{padding:60px 0 18px;}
.section-head{margin-bottom:34px;max-width:640px;}
.section-head h2{font-size:clamp(1.7rem,3.4vw,2.3rem);font-weight:800;letter-spacing:-0.03em;margin:0 0 10px;color:var(--ink);}
.section-head p{color:var(--muted);font-size:1.05rem;margin:0;}
.about-text{font-size:1.12rem;color:#334155;max-width:720px;line-height:1.85;}

/* 카드 */
.grid{display:grid;grid-template-columns:repeat(3,1fr);gap:20px;}
.card{background:var(--surface);border:1px solid var(--line);border-radius:18px;padding:28px 26px;transition:border-color .15s,box-shadow .15s;}
.card:hover{border-color:var(--accent);box-shadow:0 10px 26px rgba(15,23,42,0.06);}
.card .ico{width:46px;height:46px;display:grid;place-items:center;background:var(--accent-soft);border-radius:12px;font-size:1.4rem;margin-bottom:16px;}
.card h3{font-size:1.15rem;font-weight:700;margin:0 0 8px;color:var(--ink);}
.card p{color:var(--muted);margin:0;font-size:0.96rem;line-height:1.7;}
.people-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:16px;}
.member{background:var(--surface);border:1px solid var(--line);border-radius:16px;padding:22px;}
.m-name{font-size:1.08rem;font-weight:700;margin:0 0 4px;color:var(--ink);}
.m-role{font-size:0.88rem;color:var(--accent);font-weight:600;margin:0 0 10px;}
.m-detail{font-size:0.9rem;color:var(--muted);margin:0;line-height:1.6;}

/* 공강표 */
.sched{width:100%;border-collapse:collapse;font-size:0.85rem;}
.sched th,.sched td{border:1px solid var(--line);padding:6px 5px;text-align:center;}
.sched thead th{background:#F1F5F9;font-weight:700;color:var(--ink);}
.sched .time{background:#F8FAFC;color:var(--muted);font-weight:600;font-size:0.75rem;white-space:nowrap;}
.lvl0{background:#0D9488;color:#fff;font-weight:700;}
.lvl1{background:#99F6E4;color:#0F766E;font-weight:600;}
.lvl2{background:#FEF3C7;color:#92400E;}
.lvl3{background:#FED7AA;color:#9A3412;}
.lvl4{background:#FEE2E2;color:#991B1B;}
.grid-free{background:#E6F4F1;color:var(--accent-dark);font-weight:600;}
.grid-busy{background:#EEF1F5;color:#64748B;font-size:0.74rem;line-height:1.25;}
.legend{display:flex;gap:7px;margin:12px 0 2px;flex-wrap:wrap;align-items:center;}
.legend .lbl{font-size:0.8rem;color:var(--muted);margin-right:4px;}
.lg{display:inline-block;padding:2px 10px;border-radius:6px;font-size:0.76rem;font-weight:600;}
.summary{display:grid;grid-template-columns:1fr 1fr;gap:24px;margin-top:22px;}
.summary h4{margin:0 0 10px;font-size:0.96rem;font-weight:700;color:var(--ink);}
.ftrow{font-size:0.9rem;color:#334155;margin-bottom:5px;}
.ftrow b{display:inline-block;width:22px;color:var(--accent-dark);}
.summary ol{margin:0;padding-left:20px;}
.summary li{font-size:0.9rem;color:#334155;margin-bottom:5px;}
.ind-note{font-size:0.88rem;color:var(--muted);margin:2px 0 12px;}
.clist{margin-top:20px;}
.clist h4{margin:0 0 8px;font-size:0.96rem;font-weight:700;color:var(--ink);}
.clrow{display:flex;gap:12px;align-items:baseline;padding:7px 0;border-bottom:1px solid var(--line);font-size:0.92rem;}
.clrow:last-child{border-bottom:none;}
.cld{width:46px;color:var(--accent-dark);font-weight:600;flex:none;}
.clt{width:108px;color:var(--muted);flex:none;}
.cln{color:var(--ink);}
.freeall{margin-top:20px;padding:16px 18px;background:#E6F4F1;border-radius:12px;color:var(--accent-dark);font-weight:600;}

/* 연락처 */
.contact{background:var(--ink);border-radius:22px;padding:46px 40px;margin-top:24px;}
.contact .section-head{margin-bottom:20px;}
.contact .section-head h2{color:#fff;}
.contact .section-head p{color:#94A3B8;}
.contact-meta{color:#94A3B8;margin:0 0 20px;font-size:0.96rem;}
.contact-links{display:flex;gap:12px;flex-wrap:wrap;}
.contact-links a{color:#fff;text-decoration:none;background:var(--accent);border:1px solid var(--accent);
padding:12px 24px;border-radius:12px;font-weight:600;}
.contact-links a:hover{background:var(--accent-dark);}
.contact-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:22px 30px;margin:0 0 8px;}
.cblock .ck{color:#5EEAD4;font-size:0.78rem;font-weight:700;letter-spacing:0.03em;margin:0 0 9px;}
.cblock .cv{color:#E2E8F0;font-size:0.92rem;line-height:1.55;margin:4px 0;}
.cblock .cv.en{color:#94A3B8;font-size:0.82rem;}
.cblock .cv a{color:#5EEAD4;text-decoration:none;border-bottom:1px solid #134E4A;font-weight:600;}
.cblock .cmut{color:#94A3B8;font-size:0.8rem;margin:8px 0 0;line-height:1.5;}
.foot{text-align:center;color:var(--muted);padding:26px 0 6px;font-size:0.88rem;}

/* 네이티브 탭 */
[data-testid="stTabs"] [data-baseweb="tab-list"]{gap:4px;}
[data-testid="stTabs"] button[data-baseweb="tab"]{font-family:'Pretendard',sans-serif;font-weight:600;color:var(--muted);}
[data-testid="stTabs"] button[aria-selected="true"]{color:var(--accent);}
[data-testid="stTabs"] [data-baseweb="tab-highlight"]{background-color:var(--accent);}
[data-testid="stTabs"] [data-baseweb="tab-border"]{background-color:var(--line);}

/* ── 주문 관리: 요약 카드 ── */
.osum{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin:4px 0 22px;}
.ocard{background:var(--surface);border:1px solid var(--line);border-radius:16px;padding:18px 20px;}
.ocard .lab{font-size:0.82rem;color:var(--muted);font-weight:600;margin:0 0 6px;}
.ocard .val{font-size:1.55rem;font-weight:800;color:var(--ink);letter-spacing:-0.02em;}
.ocard .val.accent{color:var(--accent-dark);}

/* ── 달력 ── */
.monthlabel{text-align:center;font-size:1.25rem;font-weight:800;color:var(--ink);letter-spacing:-0.02em;padding-top:4px;}
.calwrap{margin:8px 0 6px;}
.cal{width:100%;border-collapse:collapse;table-layout:fixed;}
.cal th{padding:9px 6px;font-size:0.82rem;font-weight:700;color:var(--muted);border-bottom:2px solid var(--line);}
.cal th.sun{color:#DC2626;}
.cal th.sat{color:#2563EB;}
.cal td{border:1px solid var(--line);vertical-align:top;height:106px;width:14.28%;padding:6px 6px 8px;background:var(--surface);}
.cal td.empty{background:#FBFCFE;}
.cal td.today{background:var(--accent-soft);}
.cal .daynum{font-size:0.8rem;font-weight:700;color:#334155;margin-bottom:4px;}
.cal .daynum.sun{color:#DC2626;}
.cal .daynum.sat{color:#2563EB;}
.cal .chip{display:block;font-size:0.72rem;font-weight:600;padding:2px 7px;border-radius:6px;margin-top:3px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.clg{display:inline-block;padding:3px 10px;border-radius:6px;font-size:0.78rem;font-weight:600;}

/* ── 진행 단계 pill ── */
.stagerow{display:flex;flex-wrap:wrap;gap:6px;margin:2px 0;}
.spill{display:inline-flex;align-items:center;gap:6px;font-size:0.82rem;font-weight:600;padding:5px 12px 5px 6px;border-radius:999px;border:1px solid var(--line);color:var(--muted);background:#fff;}
.spill .sdot{width:19px;height:19px;border-radius:50%;display:grid;place-items:center;font-size:0.72rem;background:#E2E8F0;color:#94A3B8;flex:none;}
.spill.done{background:var(--accent-soft);border-color:transparent;color:var(--accent-dark);}
.spill.done .sdot{background:var(--accent);color:#fff;}
.spill.cur{border-color:var(--accent);color:var(--accent-dark);box-shadow:0 0 0 3px var(--accent-soft);}
.spill.cur .sdot{background:#fff;color:var(--accent-dark);border:1.5px solid var(--accent);}
.spill.todo{opacity:0.65;}

/* ── 주문 카드 ── */
.ordcard{background:var(--surface);border:1px solid var(--line);border-radius:16px;padding:16px 18px 12px;margin-top:6px;}
.ordtop{display:flex;justify-content:space-between;align-items:baseline;flex-wrap:wrap;gap:6px;margin-bottom:11px;}
.orditem{font-size:1.06rem;font-weight:700;color:var(--ink);}
.ordmeta{font-size:0.86rem;color:var(--muted);}
.ordmeta b{color:var(--accent-dark);font-weight:700;}
.ordsub{font-size:0.8rem;color:#94A3B8;margin:-4px 0 11px;}
.ordnote{font-size:0.85rem;color:#64748B;margin:9px 0 2px;}
.syncbar{font-size:0.9rem;color:var(--muted);background:#F0FDFA;border:1px solid #CCFBF1;
  border-radius:12px;padding:11px 16px;margin:2px 0 16px;line-height:1.6;}
.syncbar a{color:var(--accent-dark);font-weight:700;text-decoration:none;border-bottom:1px solid #99F6E4;}
.syncbar b{color:var(--ink);font-weight:700;}
.osrc{display:inline-block;font-size:0.7rem;font-weight:700;padding:2px 8px;border-radius:999px;margin-left:8px;vertical-align:middle;}
.osrc.sheet{background:#ECFEFF;color:#0E7490;border:1px solid #A5F3FC;}
.osrc.manual{background:#F5F3FF;color:#6D28D9;border:1px solid #DDD6FE;}
.ordcard.manual{border-left:3px solid #A78BFA;}
.ordhint{font-size:0.78rem;color:#94A3B8;margin:6px 0 0;}
.catpill{display:inline-block;font-size:0.72rem;font-weight:600;color:#475569;background:#F1F5F9;border:1px solid var(--line);padding:1px 8px;border-radius:6px;}
.guidebox{background:var(--surface);border:1px solid var(--line);border-radius:14px;padding:14px 18px;margin:2px 0;}
.guidebox h4{margin:2px 0 8px;font-size:0.96rem;color:var(--ink);font-weight:800;}
.guidebox h4.mt{margin-top:18px;}
.gflow{display:flex;flex-wrap:wrap;align-items:center;gap:8px;margin:4px 0 2px;}
.gflow .gtag{background:var(--accent-soft);color:var(--accent-dark);font-weight:700;font-size:0.82rem;padding:4px 11px;border-radius:999px;}
.gflow .garr{color:var(--muted);font-weight:800;}
.gtype{font-size:0.85rem;color:var(--muted);margin:3px 0 2px;}
.gtype b{color:var(--ink);}
.gex{font-size:0.8rem;color:var(--muted);font-weight:700;margin:10px 0 0;}
.gimg{width:100%;max-width:560px;border-radius:10px;border:1px solid var(--line);margin:6px 0 2px;display:block;}
.gcaut{font-size:0.88rem;color:#475569;margin:6px 0;padding-left:18px;position:relative;line-height:1.5;}
.gcaut::before{content:"•";position:absolute;left:4px;color:var(--accent);font-weight:700;}
.gcaut .warn{color:#B45309;font-weight:600;}
.vend{font-size:0.9rem;color:var(--ink);margin:8px 0;line-height:1.6;}
.vend .vn{font-weight:700;}
.vend .vc{color:var(--muted);}
.vend a{color:var(--accent-dark);text-decoration:none;border-bottom:1px solid #99F6E4;}
.stButton>button{border-radius:10px;font-family:'Pretendard',sans-serif;font-weight:600;}

/* ── 휴가 현황 ── */
.ocard .sub{font-size:0.78rem;color:var(--muted);margin:5px 0 0;line-height:1.45;}
.vscroll{overflow-x:auto;margin:6px 0 4px;}
.vsum{width:100%;border-collapse:collapse;font-size:0.86rem;background:var(--surface);}
.vsum th,.vsum td{border:1px solid var(--line);padding:7px 6px;text-align:center;white-space:nowrap;}
.vsum thead th{background:#F1F5F9;font-weight:700;color:var(--ink);}
.vsum td.nm{text-align:left;font-weight:600;color:var(--ink);padding-left:12px;}
.vsum td.nm .grp{font-size:0.7rem;color:#94A3B8;font-weight:500;margin-left:6px;}
.vsum td.z{color:#CBD5E1;}
.vsum td.v1{background:#E6F4F1;color:#0F766E;font-weight:600;}
.vsum td.v2{background:#99F6E4;color:#0F766E;font-weight:700;}
.vsum td.v3{background:#0D9488;color:#fff;font-weight:700;}
.vsum td.tot{background:#F8FAFC;font-weight:800;color:var(--ink);}
.vsum td.sk{color:#BE123C;font-weight:600;}
.vsum tfoot td{background:#F1F5F9;font-weight:700;color:var(--ink);}
.vnote{font-size:0.8rem;color:var(--muted);margin:8px 0 0;}
.vstat{display:flex;flex-wrap:wrap;gap:8px;margin:12px 0 10px;}
.vstat span{font-size:0.84rem;font-weight:600;padding:5px 12px;border-radius:999px;background:#F1F5F9;color:#475569;}
.vstat span b{color:var(--ink);}
.vrow{display:flex;gap:12px;align-items:center;padding:8px 2px;border-bottom:1px solid var(--line);font-size:0.92rem;}
.vrow:last-child{border-bottom:none;}
.vrow .vd{width:96px;color:var(--ink);font-weight:600;flex:none;}
.vrow .vt{width:88px;flex:none;}
.vrow .vdy{width:74px;color:var(--muted);flex:none;}
.vrow .vx{color:#64748B;}
.vempty{padding:16px 18px;background:#F8FAFC;border:1px dashed #CBD5E1;border-radius:12px;color:var(--muted);font-size:0.92rem;}

/* ── 다크 모드에서도 스트림릿 기본 위젯 글자가 보이도록 강제 (배경은 항상 밝음) ── */
[data-testid="stAppViewContainer"]{color:var(--ink);}
.stExpander summary,.stExpander summary p,.stExpander summary span,.stExpander summary div,
[data-testid="stExpander"] summary,[data-testid="stExpander"] summary p{color:var(--ink) !important;}
.stExpander summary svg,[data-testid="stExpander"] summary svg{fill:var(--ink) !important;}
[data-testid="stTabs"] [data-testid="stTab"],[data-testid="stTabs"] button[data-baseweb="tab"]{color:var(--muted) !important;}
[data-testid="stTabs"] [data-testid="stTab"] *,[data-testid="stTabs"] button[data-baseweb="tab"] *{color:inherit !important;}
[data-testid="stTabs"] [data-testid="stTab"][aria-selected="true"],
[data-testid="stTabs"] button[data-baseweb="tab"][aria-selected="true"],
[data-testid="stTabs"] [data-testid="stTab"][aria-selected="true"] *{color:var(--accent-dark) !important;}
[data-testid="stWidgetLabel"],[data-testid="stWidgetLabel"] p,label[data-testid="stWidgetLabel"] p{color:var(--ink) !important;}
[data-testid="stCaption"],[data-testid="stCaption"] p{color:var(--muted) !important;}

@media (max-width:820px){
  .links{display:none;}
  .grid,.people-grid,.summary{grid-template-columns:1fr;}
  .contact-grid{grid-template-columns:1fr;gap:20px;}
  .osum{grid-template-columns:repeat(2,1fr);}
  .cal td{height:78px;}
  .cal .chip{font-size:0.62rem;padding:1px 5px;}
  .vrow{gap:8px;font-size:0.86rem;}
  .vrow .vd{width:80px;}
  .vrow .vdy{width:58px;}
  .hero{padding:56px 0 44px;}
  .section{padding:44px 0;}
}
</style>
"""

# ── 반복 카드 HTML ──
members_html = "".join(
    f'<div class="member"><p class="m-name">{m["name"]}</p><p class="m-role">{m["role"]}</p>'
    f'<p class="m-detail">{m["detail"]}</p></div>' for m in MEMBERS)

# ══════════════════════════════════════════════════
#  📦 주문 데이터 & 헬퍼 (구글 시트 읽기 + 로컬 진행/직접 추가)
# ══════════════════════════════════════════════════
# 이 표는 두 가지를 합쳐서 보여줍니다.
#   ① 구글 시트(Park Group…, 시약)를 "읽기 전용"으로 읽어 요청→교수확인→주문→입고 단계를 자동 반영.
#      (시트는 절대 수정하지 않습니다.)
#   ② 입고 이후의 검수 / 견적서·검수 마무리 단계는 직접 체크(로컬 저장), 시트에 없는
#      소모품·장비·기타 주문은 직접 추가(로컬 저장)합니다.  → data.json 파일에 저장돼요.
#
# ▸ 시트 읽는 방법은 자동 선택: secrets 에 서비스 계정([gcp_service_account])이 있으면
#   gspread(비공개 시트)로, 없으면 공개 CSV 링크로 읽습니다. 자세한 건 아래 주석/안내 참고.
SHEET_ID = "18b88VxnCWBw-mYZBu292q1MSGbbIrf1XDRI_8lu3JyU"
GID = "2103689234"          # 화면에 보여줄 탭(sheet)의 gid — 시트 URL의 gid= 뒤 숫자
SHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit#gid={GID}"
CSV_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={GID}"
LOCAL_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data.json")

# 시트 열 이름 → 내부 필드 (열 순서가 바뀌어도 '이름'으로 찾습니다)
COLMAP = {
    "name": "Your Name", "vendor": "Vendor", "amt": "Amount/unit",
    "units": "# of units", "total": "Total (₩)", "cas": "CAS #",
    "prodno": "제품번호", "desc": "Product description", "date": "Date",
    "prof": "Prof. Check", "notes": "Extra notes", "ordered": "Ordered date",
    "arrived": "왔나요",
}

# 6단계.  0~3 = 시트에서 자동 판별 / 4~5 = 입고 후 직접 체크(내가 처리하는 부분)
STAGES = ["주문 요청", "교수 확인", "주문 완료", "입고 완료", "검수", "견적서·검수 마무리"]
SHEET_MAX = 3               # 시트가 채워주는 마지막 단계(입고 완료)의 인덱스
LAST = len(STAGES) - 1
CATEGORIES = ["시약", "소모품", "장비", "기타"]
STAGE_COLORS = [
    ("#E2E8F0", "#475569"),   # 0 요청
    ("#DBEAFE", "#1D4ED8"),   # 1 교수 확인
    ("#E0E7FF", "#4338CA"),   # 2 주문 완료
    ("#FEF3C7", "#B45309"),   # 3 입고 완료
    ("#FCE7F3", "#BE185D"),   # 4 검수 (내 처리)
    ("#CCFBF1", "#0F766E"),   # 5 마무리 완료
]


def _esc(s):
    return html.escape(str(s)) if s else ""


def _today():
    return datetime.date.today().isoformat()


def _parse_date(s):
    s = (s or "").strip()
    if not s:
        return None
    parts = s.replace("-", ".").replace("/", ".").split(".")
    try:
        y, m, d = int(parts[0]), int(parts[1]), int(parts[2])
        if y < 100:
            y += 2000
        return datetime.date(y, m, d)
    except Exception:
        return None


def _fmt(iso):
    try:
        return datetime.date.fromisoformat(iso).strftime("%m/%d")
    except Exception:
        return iso or "-"


def _won(s):
    digits = "".join(ch for ch in (s or "") if ch.isdigit())
    return int(digits) if digits else 0


# ── 로컬 저장소 (직접 추가한 주문 + 시트 주문의 입고후 단계 오버레이) ──
def _local_load():
    if os.path.exists(LOCAL_PATH):
        try:
            with open(LOCAL_PATH, "r", encoding="utf-8") as f:
                d = json.load(f)
            d.setdefault("manual", [])
            d.setdefault("overlay", {})
            return d
        except Exception:
            pass
    return {"manual": [], "overlay": {}}


def _local_save(d):
    try:
        with open(LOCAL_PATH, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False, indent=2)
    except Exception as e:
        st.warning(f"로컬 데이터를 저장하지 못했어요: {e}")


def _order_key(o):
    """시트 주문의 안정적 키 (행 순서가 바뀌어도 유지되도록 내용 기반)."""
    return "|".join([o.get("request_date", ""), o.get("requester", ""),
                     (o.get("cas") or o.get("item") or "")])


# ── 구글 시트 읽기 ──
@st.cache_data(ttl=300, show_spinner=False)
def _fetch_sheet_rows():
    """구글 시트를 표(행들의 리스트)로 읽어옵니다.
    (1) secrets 에 서비스 계정이 있으면 비공개 시트도 gspread 로,
    (2) 없으면 공개 CSV 링크로 읽습니다."""
    use_sa = False
    try:
        use_sa = "gcp_service_account" in st.secrets
    except Exception:
        use_sa = False
    if use_sa:
        import gspread
        gc = gspread.service_account_from_dict(dict(st.secrets["gcp_service_account"]))
        sh = gc.open_by_key(SHEET_ID)
        ws = next((w for w in sh.worksheets() if str(w.id) == str(GID)), sh.sheet1)
        return ws.get_all_values()
    req = urllib.request.Request(CSV_URL, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=15, context=ssl.create_default_context()) as resp:
        raw = resp.read().decode("utf-8")
    return list(csv.reader(io.StringIO(raw)))


def parse_sheet_orders(rows):
    """시트 행 → 주문 dict 목록 (stage 0~3 = 시트 자동 판별)."""
    hidx = None
    for i, r in enumerate(rows):
        cells = [c.strip() for c in r]
        if COLMAP["name"] in cells or COLMAP["date"] in cells:
            hidx = i
            break
    if hidx is None:
        return []
    hdr = [c.strip() for c in rows[hidx]]
    idx = {k: (hdr.index(v) if v in hdr else None) for k, v in COLMAP.items()}
    arrived_j = idx.get("arrived")

    def cell(r, key):
        j = idx.get(key)
        return r[j].strip() if (j is not None and j < len(r)) else ""

    out = []
    for n, r in enumerate(rows[hidx + 1:]):
        if not any(c.strip() for c in r):
            continue
        name, desc = cell(r, "name"), cell(r, "desc")
        d = _parse_date(cell(r, "date"))
        if not (name or desc) or d is None:
            continue
        arrived, ordered, prof = cell(r, "arrived"), cell(r, "ordered"), cell(r, "prof")
        if arrived.upper().startswith("O"):
            base = 3
        elif ordered:
            base = 2
        elif prof.upper().startswith("O"):
            base = 1
        else:
            base = 0
        extra = ""
        if arrived_j is not None:
            extra = " · ".join(c.strip() for c in r[arrived_j + 1:] if c.strip())
        note = " · ".join(x for x in [cell(r, "notes"), extra] if x)
        qty, units = cell(r, "amt"), cell(r, "units")
        out.append({
            "item": desc or "(품명 미기재)",
            "requester": name or "-",
            "vendor": cell(r, "vendor") or "-",
            "amount": _won(cell(r, "total")),
            "request_date": d.isoformat(),
            "ordered_date": cell(r, "ordered"),
            "stage": base,
            "category": "시약",
            "qty": f"{qty} × {units}" if (units and units not in ("1", "")) else qty,
            "cas": cell(r, "cas"),
            "prodno": cell(r, "prodno"),
            "note": note,
        })
    return out


def build_orders():
    """시트 주문(자동 0~3 + 입고후 오버레이) + 직접 추가 주문을 합칩니다.
    반환: (orders, sheet_error_or_None). 시트 실패해도 직접 추가분은 계속 보입니다."""
    local = _local_load()
    overlay = local.get("overlay", {})
    err = None
    sheet_orders = []
    try:
        sheet_orders = parse_sheet_orders(_fetch_sheet_rows())
    except Exception as e:
        err = str(e)
    out = []
    for o in sheet_orders:
        o = dict(o)
        o["source"] = "sheet"
        base = o["stage"]
        o["base_stage"] = base
        ov = overlay.get(_order_key(o))
        if ov and ov.get("stage", base) > base:
            o["stage"] = min(ov["stage"], LAST)
        o["ov_dates"] = (ov or {}).get("dates", {})
        out.append(o)
    for m in local.get("manual", []):
        m = dict(m)
        m["source"] = "manual"
        m["base_stage"] = 0
        m.setdefault("category", "기타")
        m.setdefault("stage", 0)
        m.setdefault("ordered_date", "")
        out.append(m)
    return out, err


# ── 단계 진행 / 되돌리기 (콜백은 로컬 파일만 수정) ──
def _advance(o):
    local = _local_load()
    if o["source"] == "manual":
        for r in local["manual"]:
            if r["id"] == o["id"] and r.get("stage", 0) < LAST:
                r["stage"] = r.get("stage", 0) + 1
                r.setdefault("stage_dates", {})[str(r["stage"])] = _today()
    else:
        key = _order_key(o)
        ov = local["overlay"].get(key, {"stage": o["base_stage"], "dates": {}})
        cur = max(ov.get("stage", o["base_stage"]), o["base_stage"])
        if cur < LAST:
            cur += 1
            ov["stage"] = cur
            ov.setdefault("dates", {})[str(cur)] = _today()
            local["overlay"][key] = ov
    _local_save(local)


def _revert(o):
    local = _local_load()
    if o["source"] == "manual":
        for r in local["manual"]:
            if r["id"] == o["id"] and r.get("stage", 0) > 0:
                r.get("stage_dates", {}).pop(str(r["stage"]), None)
                r["stage"] = r["stage"] - 1
    else:
        key = _order_key(o)
        ov = local["overlay"].get(key)
        if ov and ov.get("stage", o["base_stage"]) > o["base_stage"]:
            ov["dates"].pop(str(ov["stage"]), None)
            ov["stage"] -= 1
            if ov["stage"] <= o["base_stage"]:
                local["overlay"].pop(key, None)   # 시트 기본값으로 복귀 → 오버레이 제거
            else:
                local["overlay"][key] = ov
    _local_save(local)


def _add_manual(item, category, requester, vendor, amount, req_date, note):
    local = _local_load()
    local["manual"].append({
        "id": uuid.uuid4().hex[:8],
        "item": item, "category": category, "requester": requester,
        "vendor": vendor, "amount": int(amount), "request_date": req_date,
        "ordered_date": "", "stage": 0, "stage_dates": {"0": req_date}, "note": note,
    })
    _local_save(local)


def _save_manual(oid, **fields):
    local = _local_load()
    for r in local["manual"]:
        if r["id"] == oid:
            new_stage = fields.pop("stage", r.get("stage", 0))
            r.update(fields)
            sd = r.setdefault("stage_dates", {})
            if new_stage != r.get("stage", 0):
                for i in range(new_stage + 1):
                    sd.setdefault(str(i), r.get("request_date", _today()))
                for i in range(new_stage + 1, len(STAGES)):
                    sd.pop(str(i), None)
                r["stage"] = new_stage
    _local_save(local)


def _delete_manual(oid):
    local = _local_load()
    local["manual"] = [r for r in local["manual"] if r["id"] != oid]
    _local_save(local)


def _change_month(delta):
    m = st.session_state.cal_month + delta
    y = st.session_state.cal_year
    while m < 1:
        m += 12
        y -= 1
    while m > 12:
        m -= 12
        y += 1
    st.session_state.cal_month = m
    st.session_state.cal_year = y


# ── HTML 조각 ──
def stepper_html(stage):
    out = ""
    for i, name in enumerate(STAGES):
        if i <= stage:
            cls, mark = "done", "✓"
        elif i == stage + 1:
            cls, mark = "cur", str(i + 1)
        else:
            cls, mark = "todo", str(i + 1)
        out += f"<span class='spill {cls}'><b class='sdot'>{mark}</b>{name}</span>"
    return f"<div class='stagerow'>{out}</div>"


def order_card_html(o):
    amt = f"₩{o['amount']:,}" if o.get("amount") else "-"
    badge = ("<span class='osrc sheet'>🔗 시트</span>" if o["source"] == "sheet"
             else "<span class='osrc manual'>✍️ 직접 추가</span>")
    bits = [f"<span class='catpill'>{_esc(o.get('category', ''))}</span>",
            f"요청자 {_esc(o['requester'])}", _esc(o["vendor"])]
    if o.get("qty"):
        bits.append(_esc(o["qty"]))
    bits.append(f"<b>{amt}</b>")
    sub = [f"요청일 {_fmt(o['request_date'])}"]
    _od = _parse_date(o.get("ordered_date", ""))
    if _od:
        sub.append(f"주문일 {_od.strftime('%m/%d')}")
    if o.get("cas"):
        sub.append(f"CAS {_esc(o['cas'])}")
    if o.get("prodno"):
        sub.append(f"#{_esc(o['prodno'])}")
    note = f"<div class='ordnote'>📝 {_esc(o['note'])}</div>" if o.get("note") else ""
    cls = "ordcard manual" if o["source"] == "manual" else "ordcard"
    return (f"<div class='{cls}'><div class='ordtop'>"
            f"<div class='orditem'>{_esc(o['item'])}{badge}</div>"
            f"<div class='ordmeta'>{' · '.join(bits)}</div></div>"
            f"<div class='ordsub'>{' · '.join(sub)}</div>"
            f"{stepper_html(o['stage'])}{note}</div>")


def render_calendar(year, month, orders):
    by_day = {}
    for o in orders:
        try:
            dt = datetime.date.fromisoformat(o["request_date"])
        except Exception:
            continue
        if dt.year == year and dt.month == month:
            by_day.setdefault(dt.day, []).append(o)
    wd = ["일", "월", "화", "수", "목", "금", "토"]
    head = "<tr>" + "".join(
        f"<th class='{'sun' if i == 0 else 'sat' if i == 6 else ''}'>{w}</th>"
        for i, w in enumerate(wd)) + "</tr>"
    today = datetime.date.today()
    weeks = _cal.Calendar(firstweekday=6).monthdayscalendar(year, month)
    body = ""
    for week in weeks:
        body += "<tr>"
        for i, day in enumerate(week):
            if day == 0:
                body += "<td class='empty'></td>"
                continue
            is_today = (year == today.year and month == today.month and day == today.day)
            daycls = "day today" if is_today else "day"
            numcls = "sun" if i == 0 else "sat" if i == 6 else ""
            chips = ""
            for o in by_day.get(day, []):
                bg, fg = STAGE_COLORS[o["stage"]]
                pre = "✍️ " if o["source"] == "manual" else ""
                raw = o["item"]
                nm = raw if len(raw) <= 8 else raw[:7] + "…"
                chips += (f"<span class='chip' style='background:{bg};color:{fg}'>"
                          f"{pre}{_esc(nm)} · {o['stage'] + 1}/{len(STAGES)}</span>")
            body += f"<td class='{daycls}'><div class='daynum {numcls}'>{day}</div>{chips}</td>"
        body += "</tr>"
    return f"<table class='cal'><thead>{head}</thead><tbody>{body}</tbody></table>"


def cal_legend():
    items = ""
    for i, name in enumerate(STAGES):
        bg, fg = STAGE_COLORS[i]
        items += f"<span class='clg' style='background:{bg};color:{fg}'>{i + 1}. {name}</span>"
    return f"<div class='legend'><span class='lbl'>진행 단계</span>{items}</div>"


# ══════════════════════════════════════════════════
#  🌴 휴가 데이터 & 헬퍼 (구글 시트 읽기 전용 — 설정은 코드 맨 위 VAC_SHEET_LINK)
# ══════════════════════════════════════════════════
VAC_KINDS = {   # 분류 → (범례 이름, 배경색, 글자색)
    "annual":  ("연차", "#CCFBF1", "#0F766E"),
    "half":    ("반차", "#E0F2FE", "#0369A1"),
    "sick":    ("병가", "#FFE4E6", "#BE123C"),
    "special": ("특별휴가", "#EDE9FE", "#6D28D9"),
    "other":   ("기타", "#F1F5F9", "#475569"),
}
_KST = datetime.timezone(datetime.timedelta(hours=9))
_WD = ["월", "화", "수", "목", "금", "토", "일"]


def _kst_today():
    """서버가 해외(UTC)에 있어도 한국 날짜 기준으로 '오늘'을 계산."""
    return datetime.datetime.now(_KST).date()


def _short(name):
    """김수린 → 수린 (공강표의 조 편성 이름과 맞추기용)."""
    return name[1:] if len(name) == 3 else name


def _num(x):
    return f"{x:g}"


def _vac_kind(vtype):
    t = (vtype or "").replace(" ", "")
    if "병가" in t:
        return "sick"
    if "특별" in t:
        return "special"
    if "반차" in t:
        return "half"
    if "연차" in t or not t:
        return "annual"
    return "other"


def _vac_days(vtype, raw):
    """시트의 '차감 일수' 값이 있으면 그대로, 비어 있으면 시트 수식과 같은 규칙
    (병가·특별휴가 0 / 반차 0.5 / 그 외 1)으로 계산."""
    try:
        return float((raw or "").replace(",", "."))
    except ValueError:
        k = _vac_kind(vtype)
        return 0.0 if k in ("sick", "special") else 0.5 if k == "half" else 1.0


def _vac_date(s):
    """'2026-02-23', '2026. 2. 23', '2/23/2026', 엑셀 일련번호(46076) 등을 모두 날짜로."""
    s = (s or "").strip()
    if not s:
        return None
    if re.fullmatch(r"\d{5}(\.0+)?", s):
        return datetime.date(1899, 12, 30) + datetime.timedelta(days=int(float(s)))
    nums = [int(x) for x in re.findall(r"\d+", s)]
    if len(nums) < 3:
        return None
    a, b, c = nums[:3]
    try:
        if a > 31:
            return datetime.date(a, b, c)          # 2026-02-23
        if c > 31:
            return datetime.date(c, a, b)          # 2/23/2026
        return datetime.date(2000 + a, b, c)       # 26.02.23
    except ValueError:
        return None


def _sheet_parts(link):
    """공유 링크 → (시트 ID, gid). 링크 대신 ID만 넣어도 동작."""
    link = (link or "").strip()
    m = re.search(r"/spreadsheets/d/(?!e/)([\w-]{20,})", link)
    sid = m.group(1) if m else (link if re.fullmatch(r"[\w-]{30,}", link) else "")
    g = re.search(r"[#&?]gid=(\d+)", link)
    return sid, (g.group(1) if g else "")


def _http_rows(url):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=15, context=ssl.create_default_context()) as resp:
        raw = resp.read().decode("utf-8-sig", errors="replace")
    if raw.lstrip()[:1] == "<":
        raise PermissionError("시트가 공개되어 있지 않아요 (로그인 페이지가 돌아옴)")
    return list(csv.reader(io.StringIO(raw)))


def parse_vacations(rows):
    """'휴가 입력' 탭의 행들 → ([{name, date, type, kind, days, note}], 머리글_찾음여부).
    머리글을 못 찾으면 A~E 열(이름·사용일·유형·차감·비고) 순서로 가정합니다."""
    col = {"name": 0, "date": 1, "type": 2, "days": 3, "note": 4}
    start, found = 0, False
    for i, r in enumerate(rows[:20]):
        cells = [c.strip() for c in r]

        def find(test, cells=cells):
            return next((j for j, c in enumerate(cells) if test(c)), None)

        ni, di = find(lambda c: c.endswith("이름")), find(lambda c: c.endswith("사용일"))
        if ni is not None and di is not None:
            col = {"name": ni, "date": di, "type": find(lambda c: "유형" in c),
                   "days": find(lambda c: "차감" in c), "note": find(lambda c: c.endswith("비고"))}
            start, found = i + 1, True
            break

    def get(r, k):
        j = col.get(k)
        return r[j].strip() if (j is not None and j < len(r)) else ""

    out = []
    for r in rows[start:]:
        name, d = get(r, "name"), _vac_date(get(r, "date"))
        if not name or d is None:
            continue
        vtype = get(r, "type") or "연차"
        if not found and _vac_kind(vtype) == "other":
            continue        # 머리글 없이 열 위치로 읽을 땐 휴가 유형이 확실한 행만 (엉뚱한 탭 방지)
        out.append({"name": name, "date": d, "type": vtype, "kind": _vac_kind(vtype),
                    "days": _vac_days(vtype, get(r, "days")), "note": get(r, "note")})
    return out, found


@st.cache_data(ttl=300, show_spinner=False)
def load_vacations(link):
    """휴가 시트 → 휴가 목록. 아래 순서로 시도해서 '휴가 입력' 표가 나오는 첫 방법을 씁니다.
    ① secrets 에 서비스 계정이 있으면 gspread 로 탭 이름 검색
    ② 링크에 gid 가 있으면 그 탭의 CSV
    ③ 탭 이름(VAC_TAB)으로 CSV  ← 일반 공유 링크(gid 없음)는 여기서 읽혀요."""
    sid, gid = _sheet_parts(link)
    if not sid:
        raise ValueError("링크에서 시트 ID를 찾지 못했어요. …/spreadsheets/d/<ID>/… 형태의 링크인지 확인해 주세요.")
    errs = []
    try:
        use_sa = "gcp_service_account" in st.secrets
    except Exception:
        use_sa = False
    if use_sa:
        try:
            import gspread
            gc = gspread.service_account_from_dict(dict(st.secrets["gcp_service_account"]))
            wss = gc.open_by_key(sid).worksheets()
            ws = (next((w for w in wss if w.title.strip() == VAC_TAB), None)
                  or next((w for w in wss if str(w.id) == gid), None))
            if ws is not None:
                got, found = parse_vacations(ws.get_all_values())
                if found or got:
                    return got
        except Exception as e:
            errs.append(f"서비스 계정: {e}")
    base = f"https://docs.google.com/spreadsheets/d/{sid}"
    urls = [f"{base}/export?format=csv&gid={gid}"] if gid else []
    urls.append(f"{base}/gviz/tq?tqx=out:csv&headers={VAC_HEADER_ROWS}"
                f"&sheet={urllib.parse.quote(VAC_TAB)}")
    for u in urls:
        try:
            got, found = parse_vacations(_http_rows(u))
            if found or got:
                return got
        except Exception as e:
            errs.append(str(e))
    raise RuntimeError(" / ".join(errs) or f"‘{VAC_TAB}’ 탭을 찾지 못했어요. 탭 이름을 확인해 주세요.")


def _vac_chip_style(e):
    _, bg, fg = VAC_KINDS[e["kind"]]
    return f"background:{bg};color:{fg}"


def _vac_label(e):
    return e["type"].replace(" ", "")


def _change_vac_month(delta):
    y, m = st.session_state.vac_cal_y, st.session_state.vac_cal_m + delta
    y, m = y + (m - 1) // 12, (m - 1) % 12 + 1
    st.session_state.vac_cal_y, st.session_state.vac_cal_m = y, m


def vac_legend():
    items = "".join(f"<span class='clg' style='background:{bg};color:{fg}'>{lab}</span>"
                    for lab, bg, fg in VAC_KINDS.values())
    return f"<div class='legend'><span class='lbl'>휴가 유형</span>{items}</div>"


def render_vac_calendar(year, month, entries):
    by_day = {}
    for e in entries:
        if e["date"].year == year and e["date"].month == month:
            by_day.setdefault(e["date"].day, []).append(e)
    wd = ["일", "월", "화", "수", "목", "금", "토"]
    head = "<tr>" + "".join(
        f"<th class='{'sun' if i == 0 else 'sat' if i == 6 else ''}'>{w}</th>"
        for i, w in enumerate(wd)) + "</tr>"
    today = _kst_today()
    body = ""
    for week in _cal.Calendar(firstweekday=6).monthdayscalendar(year, month):
        body += "<tr>"
        for i, day in enumerate(week):
            if day == 0:
                body += "<td class='empty'></td>"
                continue
            daycls = "day today" if datetime.date(year, month, day) == today else "day"
            numcls = "sun" if i == 0 else "sat" if i == 6 else ""
            chips = ""
            for e in sorted(by_day.get(day, []), key=lambda x: x["name"]):
                tip = f"{e['name']} {e['type']}" + (f" ({e['note']})" if e["note"] else "")
                chips += (f"<span class='chip' title='{_esc(tip)}' style='{_vac_chip_style(e)}'>"
                          f"{_esc(_short(e['name']))} {_esc(_vac_label(e))}</span>")
            body += f"<td class='{daycls}'><div class='daynum {numcls}'>{day}</div>{chips}</td>"
        body += "</tr>"
    return f"<table class='cal'><thead>{head}</thead><tbody>{body}</tbody></table>"


def render_vac_summary(entries, year, names):
    """시트 '월별 요약'과 같은 표: 이름 × 월 (차감 일수 합), 합계, 병가·특별휴가(건)."""
    ys = [e for e in entries if e["date"].year == year]
    ms = {e["date"].month for e in ys}
    today = _kst_today()
    if year == today.year:
        ms.add(today.month)
    months = list(range(min(ms), max(ms) + 1)) if ms else list(range(1, 13))

    def cls(v):
        return "z" if v == 0 else "v1" if v < 2 else "v2" if v < 4 else "v3"

    head = ("<tr><th>이름</th>" + "".join(f"<th>{m}월</th>" for m in months)
            + "<th>합계</th><th>병가·특별휴가</th></tr>")
    body = ""
    col_tot = [0.0] * len(months)
    all_tot, all_sick = 0.0, 0
    for n in names:
        mine = [e for e in ys if e["name"] == n]
        per = [sum(e["days"] for e in mine if e["date"].month == m) for m in months]
        tot = sum(per)
        sick = sum(1 for e in mine if e["kind"] in ("sick", "special"))
        col_tot = [a + b for a, b in zip(col_tot, per)]
        all_tot += tot
        all_sick += sick
        grp = group_of(_short(n))
        cells = "".join(f"<td class='{cls(v)}'>{_num(v) if v else '–'}</td>" for v in per)
        body += (f"<tr><td class='nm'>{_esc(n)}<span class='grp'>{_esc(grp)}</span></td>{cells}"
                 f"<td class='tot'>{_num(tot)}</td>"
                 f"<td class='{'sk' if sick else 'z'}'>{sick if sick else '–'}</td></tr>")
    foot = ("<tr><td class='nm'>합계</td>" + "".join(f"<td>{_num(v)}</td>" for v in col_tot)
            + f"<td>{_num(all_tot)}</td><td>{all_sick}</td></tr>")
    return (f"<div class='vscroll'><table class='vsum'><thead>{head}</thead>"
            f"<tbody>{body}</tbody><tfoot>{foot}</tfoot></table></div>"
            "<div class='vnote'>숫자는 차감 일수(연차 1 · 반차 0.5). 병가·특별휴가는 차감 없이 건수만 셉니다.</div>")


def render_vac_person(entries, name, year):
    mine = sorted((e for e in entries if e["name"] == name and e["date"].year == year),
                  key=lambda e: e["date"])
    if not mine:
        return f"<div class='vempty'>{_esc(name)} 님의 {year}년 휴가 기록이 없어요.</div>"
    cnt = {k: sum(1 for e in mine if e["kind"] == k) for k in VAC_KINDS}
    used = sum(e["days"] for e in mine)
    stats = [f"<span>차감 합계 <b>{_num(used)}일</b></span>",
             f"<span>연차 <b>{cnt['annual']}</b>회</span>",
             f"<span>반차 <b>{cnt['half']}</b>회</span>",
             f"<span>병가 <b>{cnt['sick']}</b>회</span>",
             f"<span>특별휴가 <b>{cnt['special']}</b>회</span>"]
    if cnt["other"]:
        stats.append(f"<span>기타 <b>{cnt['other']}</b>회</span>")
    today = _kst_today()
    rows = ""
    for e in mine:
        d = e["date"]
        when = f"{d.month:02d}/{d.day:02d} ({_WD[d.weekday()]})"
        if d > today:
            when += " 예정"
        dy = f"{_num(e['days'])}일" if e["days"] else "차감 없음"
        rows += (f"<div class='vrow'><span class='vd'>{when}</span>"
                 f"<span class='vt'><span class='clg' style='{_vac_chip_style(e)}'>{_esc(_vac_label(e))}</span></span>"
                 f"<span class='vdy'>{dy}</span><span class='vx'>{_esc(e['note'])}</span></div>")
    return (f"<div class='vstat'>{''.join(stats)}</div>"
            f"<div class='guidebox'>{rows}</div>")


def render_vacation_section():
    st.markdown("""
<div class="section" id="vacation" style="padding:64px 0 8px;">
  <div class="section-head">
    <h2>휴가 사용 현황</h2>
    <p>휴가 구글 시트 ‘휴가 입력’ 탭을 읽어 보여줍니다. 시트에 입력하면 반영돼요.</p>
  </div>
</div>
""", unsafe_allow_html=True)

    link = VAC_SHEET_LINK.strip()
    if not link:
        st.info("휴가 시트가 아직 연결되지 않았어요. 코드 맨 위의 `VAC_SHEET_LINK` 에 "
                "구글 시트 공유 링크를 붙여넣으면 이 영역이 채워집니다.")
        return
    try:
        entries = load_vacations(link)
    except Exception as e:
        st.warning(
            "휴가 시트를 불러오지 못했어요.\n\n"
            f"• 원인: `{e}`\n\n"
            "확인할 것: ① [공유] → ‘링크가 있는 모든 사용자: 뷰어’ "
            f"② 휴가 기록 탭 이름이 ‘{VAC_TAB}’ 인지 "
            "③ 엑셀 원본(.xlsx)이면 [파일] → ‘Google 스프레드시트로 저장’ 후 새 링크 사용")
        return

    sid, _ = _sheet_parts(link)
    sheet_url = f"https://docs.google.com/spreadsheets/d/{sid}/edit"
    today = _kst_today()
    last = max((e["date"] for e in entries), default=None)

    # ── 연동 바 ──
    vb1, vb2 = st.columns([4, 1])
    with vb1:
        st.markdown(
            f"<div class='syncbar'>🔗 <a href='{sheet_url}' target='_blank'>FIMCL 휴가 시트</a> 연동"
            f" &nbsp;·&nbsp; 전체 기록 <b>{len(entries)}건</b>"
            + (f" &nbsp;·&nbsp; 가장 최근 기록 <b>{last.isoformat()}</b>" if last else "")
            + "</div>", unsafe_allow_html=True)
    with vb2:
        st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
        st.button("🔄 새로고침", key="refresh_vac", on_click=st.cache_data.clear,
                  use_container_width=True)

    # ── 요약 카드 (한국 날짜 기준) ──
    todays = sorted((e for e in entries if e["date"] == today), key=lambda e: e["name"])
    mon = today - datetime.timedelta(days=today.weekday())
    week = sorted((e for e in entries if mon <= e["date"] <= mon + datetime.timedelta(days=6)),
                  key=lambda e: (e["date"], e["name"]))
    yr = [e for e in entries if e["date"].year == today.year]
    yr_used = sum(e["days"] for e in yr)
    yr_sick = sum(1 for e in yr if e["kind"] in ("sick", "special"))
    today_sub = (", ".join(f"{_short(e['name'])} {_vac_label(e)}" for e in todays)
                 if todays else "휴가자 없음")
    week_bits = [f"{_short(e['name'])} {e['date'].month}/{e['date'].day}" for e in week]
    week_sub = (", ".join(week_bits[:4]) + (f" 외 {len(week_bits) - 4}건" if len(week_bits) > 4 else "")
                if week_bits else "이번 주 휴가 없음")
    st.markdown(f"""
<div class="osum">
  <div class="ocard"><p class="lab">오늘 휴가</p><div class="val accent">{len({e['name'] for e in todays})}명</div>
    <p class="sub">{_esc(today_sub)}</p></div>
  <div class="ocard"><p class="lab">이번 주 휴가</p><div class="val">{len(week)}건</div>
    <p class="sub">{_esc(week_sub)}</p></div>
  <div class="ocard"><p class="lab">{today.year}년 연차 사용</p><div class="val">{_num(yr_used)}일</div>
    <p class="sub">랩 전체 차감 일수 합계</p></div>
  <div class="ocard"><p class="lab">{today.year}년 병가·특별휴가</p><div class="val">{yr_sick}건</div>
    <p class="sub">연차에서 차감되지 않음</p></div>
</div>
""", unsafe_allow_html=True)

    names = VAC_ROSTER + sorted({e["name"] for e in entries} - set(VAC_ROSTER))
    years = sorted({e["date"].year for e in entries} | {today.year}, reverse=True)

    vt = st.tabs(["📅 휴가 달력", "📊 월별 요약", "🔎 개인별 내역"])

    # ── 달력 ──
    with vt[0]:
        if "vac_cal_y" not in st.session_state:
            st.session_state.vac_cal_y, st.session_state.vac_cal_m = today.year, today.month
        VY, VM = st.session_state.vac_cal_y, st.session_state.vac_cal_m
        n1, n2, n3 = st.columns([1, 2, 1])
        with n1:
            st.button("◀ 이전 달", key="vac_prev", on_click=_change_vac_month, args=(-1,),
                      use_container_width=True)
        with n2:
            m_days = sum(e["days"] for e in entries if (e["date"].year, e["date"].month) == (VY, VM))
            st.markdown(f"<div class='monthlabel'>{VY}년 {VM}월</div>"
                        f"<div class='ind-note' style='text-align:center;margin:2px 0 0;'>"
                        f"이 달 차감 합계 {_num(m_days)}일</div>", unsafe_allow_html=True)
        with n3:
            st.button("다음 달 ▶", key="vac_next", on_click=_change_vac_month, args=(1,),
                      use_container_width=True)
        st.markdown(f"<div class='calwrap'>{render_vac_calendar(VY, VM, entries)}</div>",
                    unsafe_allow_html=True)
        st.markdown(vac_legend(), unsafe_allow_html=True)

    # ── 월별 요약 ──
    with vt[1]:
        s1, s2, _ = st.columns([1, 1, 2])
        with s1:
            sy = st.selectbox("연도", years, key="vac_sum_year")
        with s2:
            sg = st.selectbox("조", ["전체"] + list(GROUPS.keys()), key="vac_sum_grp")
        shown = names if sg == "전체" else [n for n in names if group_of(_short(n)) == sg]
        st.markdown(render_vac_summary(entries, sy, shown), unsafe_allow_html=True)

    # ── 개인별 ──
    with vt[2]:
        p1, p2, _ = st.columns([1.4, 1, 1.6])
        with p1:
            who = st.selectbox("구성원", names, key="vac_person",
                               format_func=lambda n: f"{n}  ·  {group_of(_short(n)) or '-'}")
        with p2:
            py = st.selectbox("연도", years, key="vac_person_year")
        st.markdown(render_vac_person(entries, who, py), unsafe_allow_html=True)

# ══════════════════════════════════════════════════
#  상단 (내비 + 히어로)
# ══════════════════════════════════════════════════
nav_hero_html = f"""
<div class="nav"><div class="nav-in">
  <div class="brand">{LAB['name']}<span class="dot">.</span></div>
  <div class="links">
    <a href="#orders">주문현황</a><a href="#vacation">휴가현황</a><a href="#schedule">공강표</a>
    <a href="#contact">문의</a>
  </div>
</div></div>

<div class="hero">
  <span class="eyebrow">{LAB['affiliation']}</span>
  <h1 class="hero-name">{LAB['name']}</h1>
  <p class="hero-role">{LAB['full_name']}</p>
  <p class="lead">{LAB['tagline']}</p>
  <div class="cta-row">
    <a class="btn btn-primary" href="#orders">주문 현황 보기</a>
    <a class="btn btn-ghost" href="#vacation">휴가 현황 보기</a>
    <a class="btn btn-ghost" href="#schedule">공강표 보기</a>
  </div>
</div>
"""
st.markdown(CSS + nav_hero_html, unsafe_allow_html=True)

# ══════════════════════════════════════════════════
#  📦 물품 주문 현황 (시트 읽기 + 단계 진행 + 직접 추가)
# ══════════════════════════════════════════════════
if "cal_year" not in st.session_state:
    _t = datetime.date.today()
    st.session_state.cal_year = _t.year
    st.session_state.cal_month = _t.month

orders, err = build_orders()

st.markdown("""
<div class="section" id="orders" style="padding:60px 0 8px;">
  <div class="section-head">
    <h2>물품 주문 현황</h2>
    <p>시약은 구글 시트에서 요청→교수확인→주문→입고까지 자동 반영, 입고 이후 검수·견적 단계와 소모품·기타 주문은 직접 관리합니다.</p>
  </div>
</div>
""", unsafe_allow_html=True)

if err is not None:
    st.warning(
        "구글 시트를 지금 불러오지 못했어요 — 아래에는 **직접 추가한 주문만** 표시됩니다.\n\n"
        f"• 원인: `{err}`\n\n"
        "시트를 표시하려면 둘 중 하나: ① [공유] → ‘링크가 있는 모든 사용자: 뷰어’로 변경, "
        "또는 ② 서비스 계정을 만들어 시트를 공유하고 secrets 의 [gcp_service_account] 에 키 등록 "
        "(코드 상단 주석 참고)."
    )

orders = orders or []

# ── 전체 누적 요약 (상단 바) ──
_all_n = len(orders)
_sheet_n = sum(1 for o in orders if o["source"] == "sheet")
_manual_n = _all_n - _sheet_n
_all_amt = sum(o["amount"] for o in orders)
_mine = sum(1 for o in orders if SHEET_MAX <= o["stage"] < LAST)   # 입고 후 내 처리 대기
sb1, sb2 = st.columns([4, 1])
with sb1:
    st.markdown(
        f"<div class='syncbar'>🔗 <a href='{SHEET_URL}' target='_blank'>Park Group 주문 시트</a>"
        f" 연동 &nbsp;·&nbsp; 시트 <b>{_sheet_n}</b> + 직접추가 <b>{_manual_n}</b> = 누적 <b>{_all_n}건</b>"
        f" &nbsp;·&nbsp; 총 <b>₩{_all_amt:,}</b>"
        f" &nbsp;·&nbsp; 검수·견적 <b>내 처리 대기 {_mine}건</b></div>",
        unsafe_allow_html=True)
with sb2:
    st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
    st.button("🔄 새로고침", key="refresh_sheet", on_click=st.cache_data.clear,
              use_container_width=True)

# ── 새 주문 추가 (시트에 없는 소모품·장비·기타) ──
with st.expander("➕ 새 주문 추가 · 소모품 / 장비 / 기타 (시트에 없는 주문)", expanded=False):
    st.caption("여기서 추가한 주문은 이 앱에만 저장됩니다 — 구글 시트에는 기록되지 않아요.")
    with st.form("add_order", clear_on_submit=True):
        ac1, ac2 = st.columns(2)
        with ac1:
            n_item = st.text_input("물품명 *", placeholder="예: 니트릴 장갑 M")
            n_cat = st.selectbox("분류", CATEGORIES, index=1)   # 기본값: 소모품
            n_req = st.selectbox("요청자", ALL_MEMBERS)
            n_vendor = st.text_input("구매처", placeholder="예: ○○사이언스")
        with ac2:
            n_amount = st.number_input("금액 (원)", min_value=0, step=1000, value=0)
            n_date = st.date_input("주문요청일", value=datetime.date.today())
            n_note = st.text_input("비고", placeholder="선택 입력")
        add_ok = st.form_submit_button("주문 등록", use_container_width=True)
    if add_ok:
        if not n_item.strip():
            st.warning("물품명을 입력해 주세요.")
        else:
            _add_manual(n_item.strip(), n_cat, n_req, n_vendor.strip(),
                        int(n_amount), n_date.isoformat(), n_note.strip())
            st.rerun()

# 검수 안내 예시 이미지 (base64 내장 — 별도 파일 없이 단독 실행)
EX_COURIER = "data:image/jpeg;base64,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"
EX_ITEM = "data:image/jpeg;base64,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"

# ── 물품 검수 안내 ──
with st.expander("📋 물품 검수 안내 (사진 제출 방법 · 주의사항)", expanded=False):
    st.markdown(f"""
<div class='guidebox'>
  <h4>🚚 택배로 온 것</h4>
  <div class='gflow'>
    <span class='gtag'>택배 사진</span><span class='garr'>›</span>
    <span class='gtag'>개봉 사진</span><span class='garr'>›</span>
    <span class='gtag'>물품 사진</span>
  </div>
  <div class='gtype'>+ <b>물품명</b> 함께 기재</div>
  <div class='gex'>예)</div>
  <img class='gimg' src='{EX_COURIER}' alt='택배 검수 예시'>

  <h4 class='mt'>📦 물품으로 온 것</h4>
  <div class='gflow'><span class='gtag'>물품 사진</span></div>
  <div class='gtype'>+ <b>물품명</b> 함께 기재</div>
  <div class='gex'>예)</div>
  <img class='gimg' src='{EX_ITEM}' alt='물품 검수 예시'>

  <h4 class='mt'>⚠️ 주의사항</h4>
  <div class='gcaut'>사진 파일이 아니라 <span class='warn'>PPT 파일</span>로 보내기<span class='warn'>(이재경 카카오톡)</span></div>
  <div class='gcaut'>Lot # 보이도록 촬영 및 주변을 정돈하여 촬영 <span class='warn'>(다른 물품이 보일경우 반려)</span></div>
  <div class='gcaut'><b>Thermo Fisher</b>는 <span class='warn'>주문 번호</span>도 작성하여 보내기</div>
</div>
""", unsafe_allow_html=True)

# ── 기타 구매처 연락처 ──
with st.expander("🏪 기타 구매처 연락처", expanded=False):
    st.markdown("""
<div class='guidebox'>
  <div class='vend'><span class='vn'>제이원</span> · <span class='vc'>김재훈 사장님</span><br>
    📞 010-5495-2526 &nbsp;·&nbsp; ✉️ <a href='mailto:jaehoon85@naver.com'>jaehoon85@naver.com</a></div>
  <div class='vend'><span class='vn'>석림랩텍</span> · <span class='vc'>성열민</span><br>
    ✉️ <a href='mailto:ymsung@sercim.com'>ymsung@sercim.com</a></div>
  <div class='vend'><span class='vn'>삼전순약</span> · <span class='vc'>삼전순약</span><br>
    ✉️ <a href='http://www.samchun.com/kr/sub/product/list.asp'>http://www.samchun.com/kr/sub/product/list.asp</a></div>
</div>
""", unsafe_allow_html=True)

# ── 필터 ──
fc1, fc2 = st.columns(2)
with fc1:
    f_cat = st.selectbox("분류", ["전체"] + CATEGORIES, key="f_cat")
with fc2:
    f_stat = st.selectbox("진행 상태", ["전체", "입고 전", "내 처리 대기", "마무리 완료"], key="f_stat")


def _match(o):
    if f_cat != "전체" and o.get("category") != f_cat:
        return False
    if f_stat == "입고 전" and o["stage"] >= SHEET_MAX:
        return False
    if f_stat == "내 처리 대기" and not (SHEET_MAX <= o["stage"] < LAST):
        return False
    if f_stat == "마무리 완료" and o["stage"] != LAST:
        return False
    return True


fset = [o for o in orders if _match(o)]

# ── 월 이동 ──
Y, M = st.session_state.cal_year, st.session_state.cal_month
mc1, mc2, mc3 = st.columns([1, 2, 1])
with mc1:
    st.button("◀ 이전 달", key="prev_month", on_click=_change_month, args=(-1,),
              use_container_width=True)
with mc2:
    st.markdown(f"<div class='monthlabel'>{Y}년 {M}월</div>", unsafe_allow_html=True)
with mc3:
    st.button("다음 달 ▶", key="next_month", on_click=_change_month, args=(1,),
              use_container_width=True)

month_orders = sorted(
    [o for o in fset if o["request_date"][:7] == f"{Y:04d}-{M:02d}"],
    key=lambda x: x["request_date"])

# ── 이번 달 요약 ──
_total = len(month_orders)
_wait = sum(1 for o in month_orders if SHEET_MAX <= o["stage"] < LAST)
_done = sum(1 for o in month_orders if o["stage"] == LAST)
_amount = sum(o["amount"] for o in month_orders)
st.markdown(f"""
<div class="osum">
  <div class="ocard"><p class="lab">이번 달 주문</p><div class="val">{_total}건</div></div>
  <div class="ocard"><p class="lab">내 처리 대기</p><div class="val">{_wait}건</div></div>
  <div class="ocard"><p class="lab">마무리 완료</p><div class="val accent">{_done}건</div></div>
  <div class="ocard"><p class="lab">이번 달 총액</p><div class="val">₩{_amount:,}</div></div>
</div>
""", unsafe_allow_html=True)

# ── 달력 ──
st.markdown(f"<div class='calwrap'>{render_calendar(Y, M, fset)}</div>",
            unsafe_allow_html=True)
st.markdown(cal_legend(), unsafe_allow_html=True)

# ── 이번 달 주문 목록 ──
st.markdown("<div class='section-head' style='margin:28px 0 12px;'>"
            "<h2 style='font-size:1.3rem;'>이번 달 주문 목록</h2></div>",
            unsafe_allow_html=True)

if not month_orders:
    st.info("이 달에는 표시할 주문이 없어요. 달을 이동하거나, 위의 ‘➕ 새 주문 추가’로 등록해 보세요.")
else:
    for o in month_orders:
        st.markdown(order_card_html(o), unsafe_allow_html=True)
        okey = o["id"] if o["source"] == "manual" else _order_key(o)
        if o["source"] == "sheet" and o["base_stage"] < SHEET_MAX:
            st.markdown("<div class='ordhint'>🔗 시트에서 자동 반영 중 · 입고되면 검수·견적 단계를 직접 체크할 수 있어요.</div>",
                        unsafe_allow_html=True)
        else:
            floor = 0 if o["source"] == "manual" else o["base_stage"]
            bc1, bc2, _sp = st.columns([1.15, 1.15, 5])
            with bc1:
                st.button("다음 단계 ▶", key=f"nx_{o['source']}_{okey}", on_click=_advance,
                          args=(o,), disabled=o["stage"] >= LAST, use_container_width=True)
            with bc2:
                st.button("◀ 되돌리기", key=f"pv_{o['source']}_{okey}", on_click=_revert,
                          args=(o,), disabled=o["stage"] <= floor, use_container_width=True)
        if o["source"] == "manual":
            with st.expander("✏️ 수정 / 삭제"):
                with st.form(f"edit_{o['id']}"):
                    ec1, ec2 = st.columns(2)
                    with ec1:
                        e_item = st.text_input("물품명", o["item"], key=f"ei_{o['id']}")
                        _ci = CATEGORIES.index(o["category"]) if o.get("category") in CATEGORIES else 0
                        e_cat = st.selectbox("분류", CATEGORIES, index=_ci, key=f"ec_{o['id']}")
                        _ri = ALL_MEMBERS.index(o["requester"]) if o.get("requester") in ALL_MEMBERS else 0
                        e_req = st.selectbox("요청자", ALL_MEMBERS, index=_ri, key=f"er_{o['id']}")
                        e_vendor = st.text_input("구매처", o.get("vendor", ""), key=f"ev_{o['id']}")
                    with ec2:
                        e_amount = st.number_input("금액 (원)", min_value=0, step=1000,
                                                   value=int(o.get("amount", 0)), key=f"ea_{o['id']}")
                        try:
                            _dv = datetime.date.fromisoformat(o["request_date"])
                        except Exception:
                            _dv = datetime.date.today()
                        e_date = st.date_input("주문요청일", value=_dv, key=f"ed_{o['id']}")
                        e_stage = st.select_slider("현재 단계", options=list(range(len(STAGES))),
                                                   value=o["stage"],
                                                   format_func=lambda i: f"{i + 1}. {STAGES[i]}",
                                                   key=f"es_{o['id']}")
                    e_note = st.text_input("비고", o.get("note", ""), key=f"en_{o['id']}")
                    sc1, sc2 = st.columns([3, 1])
                    save_e = sc1.form_submit_button("저장", use_container_width=True)
                    del_e = sc2.form_submit_button("🗑 삭제", use_container_width=True)
                if save_e:
                    _save_manual(o["id"], item=e_item.strip() or o["item"], category=e_cat,
                                 requester=e_req, vendor=e_vendor.strip(),
                                 amount=int(e_amount), request_date=e_date.isoformat(),
                                 note=e_note.strip(), stage=e_stage)
                    st.rerun()
                if del_e:
                    _delete_manual(o["id"])
                    st.rerun()

# ══════════════════════════════════════════════════
#  🌴 휴가 사용 현황 (구글 시트 읽기 전용)
# ══════════════════════════════════════════════════
render_vacation_section()

# ══════════════════════════════════════════════════
#  공강표 (네이티브 탭)
# ══════════════════════════════════════════════════
st.markdown("""
<div class="section sched-head" id="schedule">
  <div class="section-head">
    <h2>랩 공강표</h2>
    <p>조별로 다 같이 되는 시간과, 구성원별 available time을 확인할 수 있어요.</p>
  </div>
</div>
""", unsafe_allow_html=True)

tabs = st.tabs(["전체 랩", "2d조", "골드조", "mofcof조", "🔎 Available time (개별)"])
for tab, key in zip(tabs[:4], ["전체 랩", "2d조", "골드조", "mofcof조"]):
    with tab:
        st.markdown(render_group(key), unsafe_allow_html=True)
with tabs[4]:
    st.markdown("<div class='ind-note'>🟩 가능 · 회색 = 수업/일정. 아래에서 구성원을 선택하세요.</div>",
                unsafe_allow_html=True)
    c1, _ = st.columns([2, 3])
    with c1:
        sel = st.selectbox(
            "구성원 선택", ALL_MEMBERS,
            format_func=lambda n: f"{n}  ·  {group_of(n)}" + ("  · 수업 있음" if n in CLASSES else ""))
    st.markdown(render_person(sel), unsafe_allow_html=True)

# ══════════════════════════════════════════════════
#  하단 (구성원 + 문의)
# ══════════════════════════════════════════════════
year = datetime.date.today().year
footer_html = f"""
<div class="contact" id="contact">
  <div class="section-head">
    <h2>문의 · 연락처</h2>
    <p>주문·행정 관련 문의는 아래 연락처로 주세요.</p>
  </div>
  <div class="contact-grid">
    <div class="cblock">
      <p class="ck">랩실 주소</p>
      <p class="cv">대전광역시 유성구 대학로 291<br>자연과학동 E6-4 3131호 (34141)</p>
      <p class="cv en">Room 3131, Natural Sciences building (E6-4),<br>291 Daehak-ro, Yuseong-gu, Daejeon, 34141, Republic of Korea</p>
    </div>
    <div class="cblock">
      <p class="ck">전화</p>
      <p class="cv">랩 · E6-4 3131·3132<br><a href="tel:0423502872">042-350-2872</a></p>
      <p class="cv">한지선 선생님 · E7 6102<br><a href="tel:0423502864">042-350-2864</a> <span class="cmut">(업무 08–17시)</span></p>
      <p class="cv">교수님 오피스 · E6-4 4104<br><a href="tel:0423502832">042-350-2832</a></p>
      <p class="cmut">※ 교내 전화는 뒷자리 4자리만 입력하면 됩니다.</p>
    </div>
    <div class="cblock">
      <p class="ck">교수님 정보</p>
      <p class="cv">생년월일 · 1986. 01. 17.</p>
      <p class="cv">연구자번호 · 12513650</p>
    </div>
  </div>
</div>

<div class="foot">© {year} {LAB['name']} · {LAB['affiliation']} · Made with Streamlit</div>
"""
st.markdown(footer_html, unsafe_allow_html=True)
