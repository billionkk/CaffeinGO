# -*- coding: utf-8 -*-
"""CaffeinGO 중간발표 PPT v3 — Claude 디자인 설계안 구현"""
import math
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree as _etree

# ---- 팔레트 (Claude 설계) ----
ESPRESSO = RGBColor(0x2B, 0x1A, 0x12)
CREAM    = RGBColor(0xFA, 0xF7, 0xF2)
BEIGE    = RGBColor(0xF3, 0xE8, 0xD8)
MILK     = RGBColor(0x8A, 0x72, 0x62)
CARAMEL  = RGBColor(0xA4, 0x7E, 0x64)
AMBER    = RGBColor(0xE8, 0xA3, 0x3D)
MIDNIGHT = RGBColor(0x3D, 0x4A, 0x6B)
TERRA    = RGBColor(0xC8, 0x55, 0x3D)
WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
AMBER_LT = RGBColor(0xF7, 0xE4, 0xC2)  # 말풍선용 연앰버
MID_LT   = RGBColor(0xD9, 0xDE, 0xEA)  # 연미드나이트

FONT = "Pretendard"
LOGO = "/home/hatch/workspace/your_files/caffeingo/logo.png"
SHOT = "/tmp/site_shots"
OUT = "/home/hatch/workspace/your_files/caffeingo/CaffeinGO_중간발표.pptx"
EMU = 914400

prs = Presentation()
prs.slide_width = Inches(13.33)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]
TOTAL = 18

def bg(slide, color):
    f = slide.background.fill; f.solid(); f.fore_color.rgb = color

def tb(slide, x, y, w, h):
    return slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)).text_frame

def set_run(r, size=18, bold=False, color=ESPRESSO, name=FONT):
    r.font.name = name
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    rPr = r._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = _etree.SubElement(rPr, qn("w:rFonts"))
    rFonts.set(qn("w:eastAsia"), name)
    rFonts.set(qn("w:ascii"), name)
    rFonts.set(qn("w:hAnsi"), name)

def para(tf, text, size=18, bold=False, color=ESPRESSO, align=PP_ALIGN.LEFT, space_after=6):
    tf.word_wrap = True
    p = tf.paragraphs[0] if not tf.paragraphs[0].text else tf.add_paragraph()
    p.text = ""
    r = p.add_run(); r.text = text
    set_run(r, size, bold, color)
    p.alignment = align; p.space_after = Pt(space_after)
    return p

def addp(tf, text, size=18, bold=False, color=ESPRESSO, align=PP_ALIGN.LEFT, space_after=6):
    tf.word_wrap = True
    p = tf.add_paragraph(); p.text = ""
    r = p.add_run(); r.text = text
    set_run(r, size, bold, color)
    p.alignment = align; p.space_after = Pt(space_after)
    return p

def label(tf_or_slide, x, y, text, color=CARAMEL):
    """섹션 라벨: 02 · PROBLEM"""
    if hasattr(tf_or_slide, "add_paragraph"):
        tf = tf_or_slide
    else:
        tf = tb(tf_or_slide, x, y, 6, 0.4)
    para(tf, text, size=12, bold=True, color=color, space_after=2)
    return tf

def headline(slide, x, y, w, text, size=38, color=ESPRESSO, align=PP_ALIGN.LEFT):
    tf = tb(slide, x, y, w, 1.6)
    para(tf, text, size=size, bold=True, color=color, align=align, space_after=4)
    return tf

def dot(slide, x, y, d, color):
    sp = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(d), Inches(d))
    sp.fill.solid(); sp.fill.fore_color.rgb = color; sp.line.fill.background()
    return sp

def draw_curve(slide, x, y, w, h, color, lw=3.5, peak_x=0.28, decay=3.0, rise_end=0.22, t_end=1.0):
    """카페인 곡선: 상승 후 지수 감쇠. t_end로 곡선을 중간에 끊을 수 있음"""
    pts = []
    n = 56
    for i in range(n + 1):
        t = (i / n) * t_end
        if t < rise_end:
            fx = (t / rise_end) * peak_x
            fy = t / rise_end
        else:
            denom = max(1 - rise_end, 1e-6)
            fx = peak_x + ((t - rise_end) / denom) * (1 - peak_x)
            fy = math.exp(-((t - rise_end) / denom) * decay)
        px = int((x + fx * w) * EMU)
        py = int((y + h - fy * h * 0.92) * EMU)
        pts.append((px, py))
    fb = slide.shapes.build_freeform(pts[0][0], pts[0][1], scale=1.0)
    fb.add_line_segments(pts[1:], close=False)
    shp = fb.convert_to_shape()
    shp.fill.background()
    shp.line.color.rgb = color
    shp.line.width = Pt(lw)
    return shp

def rrect(slide, x, y, w, h, fill, line=None, radius=0.12):
    sp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    sp.fill.solid(); sp.fill.fore_color.rgb = fill
    if line:
        sp.line.color.rgb = line; sp.line.width = Pt(1.5)
    else:
        sp.line.fill.background()
    try:
        sp.adjustments[0] = radius
    except Exception:
        pass
    return sp

def progress(slide, idx):
    """하단 진행 표시 제거됨 (검수 결과 얼룩처럼 보여 삭제)"""
    return

def section_head(slide, label_text, head_text, dark=False):
    c = WHITE if dark else ESPRESSO
    lc = CARAMEL
    tf = tb(slide, 0.9, 0.55, 11.5, 1.7)
    label(tf, 0, 0, label_text, color=lc)
    para(tf, head_text, size=36, bold=True, color=c, space_after=2)

# ============ 1. 표지 (다크) ============
s = prs.slides.add_slide(BLANK); bg(s, ESPRESSO)
# 앰버 곡선: 왼쪽 아래에서 오르다 정점 직전에 끊김
draw_curve(s, 0.8, 2.9, 5.6, 2.9, AMBER, lw=5, peak_x=0.95, rise_end=0.6, t_end=0.6)
dot(s, 6.0, 2.98, 0.3, AMBER)
# 타이틀
tf = tb(s, 6.7, 2.55, 5.7, 1.4)
para(tf, "CaffeinGO", size=72, bold=True, color=WHITE, space_after=2)
addp(tf, "마신 커피를 입력하면, 잘 시간과 일어날 시간을 알려줘요", size=20, color=CARAMEL)
# 태그
tag = rrect(s, 10.9, 0.55, 1.85, 0.5, ESPRESSO, line=CARAMEL)
tf = tag.text_frame; tf.word_wrap = True
para(tf, "MID-TERM · 2026.10", size=11, bold=True, color=CARAMEL, align=PP_ALIGN.CENTER)
# 하단 팀 정보
tf = tb(s, 0.9, 6.15, 11.5, 0.8)
para(tf, "팀 세송이", size=17, bold=True, color=WHITE, space_after=6)
row = tb(s, 0.9, 6.62, 11.5, 0.5)
names = [("곽준성", "26615002"), ("배윤학", "26615026"), ("박현민", "26615025")]
for i, (nm, no) in enumerate(names):
    x = 0.9 + i * 2.2
    tf = tb(s, x, 6.62, 2.0, 0.5)
    para(tf, f"{nm}  {no}", size=13, color=MILK)
    if i < 2:
        ln = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x + 1.95), Inches(6.62), Pt(1), Inches(0.32))
        ln.fill.solid(); ln.fill.fore_color.rgb = CARAMEL; ln.line.fill.background()

# ============ 2. 목차 ============
s = prs.slides.add_slide(BLANK); bg(s, CREAM)
section_head(s, "CONTENTS", "18장으로 보는 CaffeinGO")
# 3개 챕터 (열)
chapters = [
    ("01", "왜 필요한가", "기획 배경 · 문제 심화 · 근거", "3–5장"),
    ("02", "무엇을 만들었나", "소개 · 기능 4장 · 개인 맞춤화 · 차별성", "6–12장"),
    ("03", "어떻게 만들고 어디로 가나", "핵심 기술 · 구조 · 개발 방법 · 시연 · 향후 계획", "13–18장"),
]
for i, (no, title, items, rng) in enumerate(chapters):
    x = 0.9 + i * 3.95
    ln = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(2.6), Inches(3.55), Pt(3))
    ln.fill.solid(); ln.fill.fore_color.rgb = AMBER if i == 1 else CARAMEL; ln.line.fill.background()
    tf = tb(s, x, 2.85, 3.55, 2.8)
    para(tf, no, size=54, bold=True, color=RGBColor(0xE3, 0xD2, 0xB8), space_after=2)
    addp(tf, title, size=22, bold=True, color=ESPRESSO, space_after=8)
    addp(tf, items, size=15, color=MILK, space_after=6)
    addp(tf, rng, size=13, bold=True, color=CARAMEL)
progress(s, 2)

# ============ 3. 기획 배경 ============
s = prs.slides.add_slide(BLANK); bg(s, CREAM)
section_head(s, "01 · BACKGROUND", "지금 커피\n마셔도 될까?")
# 말풍선 3개
bubbles = [
    ("오후에 커피 한 잔 할까?", "15:00", AMBER_LT, ESPRESSO),
    ("아직도 눈이 말똥말똥해", "19:30", RGBColor(0xE8, 0xC9, 0x8E), ESPRESSO),
    ("아... 잠이 안 와", "23:00", MIDNIGHT, WHITE),
]
for i, (txt, ts, fill, tc) in enumerate(bubbles):
    y = 2.1 + i * 1.55
    b = rrect(s, 8.3, y, 3.9, 1.25, fill, radius=0.25)
    tf = b.text_frame; tf.word_wrap = True
    tf.margin_left = Inches(0.25); tf.margin_right = Inches(0.25)
    para(tf, txt, size=17, bold=True, color=tc, space_after=2)
    addp(tf, ts, size=12, color=MILK if tc == ESPRESSO else BEIGE)
progress(s, 3)

# ============ 4. 문제 심화 ============
s = prs.slides.add_slide(BLANK); bg(s, CREAM)
section_head(s, "02 · PROBLEM", "오후 4시 커피, 밤 9시에도 절반이 남아 있다")
# 히어로 숫자
tf = tb(s, 0.9, 2.35, 4.5, 1.8)
para(tf, "5h", size=140, bold=True, color=AMBER, space_after=0)
addp(tf, "카페인 반감기", size=18, bold=True, color=ESPRESSO)
# 막대 3개
for i, (pct, lab) in enumerate([(1.0, "0시간"), (0.5, "5시간"), (0.25, "10시간")]):
    y = 5.35 + i * 0.6
    bg_bar = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.9), Inches(y), Inches(3.6), Inches(0.34))
    bg_bar.fill.solid(); bg_bar.fill.fore_color.rgb = BEIGE; bg_bar.line.fill.background()
    fg = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.9), Inches(y), Inches(3.6 * pct), Inches(0.34))
    fg.fill.solid(); fg.fill.fore_color.rgb = AMBER; fg.line.fill.background()
    tf = tb(s, 4.7, y - 0.06, 1.6, 0.45)
    para(tf, f"{lab}  {int(pct*100)}%", size=14, color=MILK)
# 오른쪽: 연구 근거 숫자
tf = tb(s, 6.6, 2.5, 5.8, 2.2)
para(tf, "수면 방해 — 본인은 모릅니다", size=22, bold=True, color=TERRA, space_after=8)
addp(tf, "취침 6시간 전 커피도 수면 40분 이상 단축, 깊은 잠 감소", size=16, color=MILK, space_after=4)
addp(tf, "참가자는 “잘 잤다”고 느꼈습니다 (Drake 2013)", size=16, color=MILK)
tf = tb(s, 6.6, 4.7, 5.8, 2.2)
para(tf, "개인차 — 반감기 3~7시간", size=22, bold=True, color=MIDNIGHT, space_after=8)
addp(tf, "같은 한 잔도 유전자(CYP1A2)에 따라 다르게 남는다", size=16, color=MILK, space_after=4)
addp(tf, "→ 내 몸 기준 계산이 필요하다", size=16, bold=True, color=MIDNIGHT)
tf = tb(s, 0.9, 7.02, 11.5, 0.4)
para(tf, "24개 연구 메타분석: 카페인은 평균 수면 45분 단축  ·  Gardiner 외, 2023", size=13.5, bold=True, color=CARAMEL, align=PP_ALIGN.CENTER)
progress(s, 4)

# ============ 5. 근거 ============
s = prs.slides.add_slide(BLANK); bg(s, CREAM)
section_head(s, "03 · EVIDENCE", "숫자마다 근거가 있습니다")
cards = [
    ("90분", "최적 시각", "기상 후 30~45분에 코르티솔이 정점을 찍어요. 90분 뒤면 내려가기 시작합니다.", "Clow 외, 2004"),
    ("7시간 30분", "수면시간", "수면 주기 약 90분 × 5번. 깊은 잠은 전반부, REM은 후반부에 몰려 있어요.", "NHLBI"),
    ("1.0mg/L", "효과 종료 기준", "취침 시 잔류 카페인이 수면과 관련된다는 연구에 근거한 실용적 기준입니다.", "Drake 2013 · Gardiner 2023"),
]
for i, (big, t, d, src) in enumerate(cards):
    x = 0.9 + i * 3.95
    box = rrect(s, x, 2.5, 3.55, 3.3, WHITE, radius=0.12)
    tf = tb(s, x + 0.3, 2.75, 2.95, 2.9)
    para(tf, big, size=36, bold=True, color=AMBER, space_after=4)
    addp(tf, t, size=18, bold=True, color=ESPRESSO, space_after=8)
    addp(tf, d, size=14.5, color=MILK, space_after=10)
    addp(tf, src, size=12.5, bold=True, color=CARAMEL)
tf = tb(s, 0.9, 6.15, 11.5, 0.6)
para(tf, "1.0mg/L은 연구의 직접 제시값이 아닙니다. 근거 있는 부분과 모델의 선택을 구분해서 밝힙니다.", size=14, bold=True, color=MIDNIGHT, align=PP_ALIGN.CENTER)
progress(s, 5)

# ============ 5. CaffeinGO 소개 (다크) ============
s = prs.slides.add_slide(BLANK); bg(s, ESPRESSO)
section_head(s, "04 · SOLUTION", "마신 커피를 입력하면, 잘 시간과 일어날 시간을 알려줘요", dark=True)
# 왼쪽: 수면 플랜 스크린샷 (크게)
s.shapes.add_picture(f"{SHOT}/shot_plan.png", Inches(0.9), Inches(2.9), width=Inches(5.2))
# 오른쪽: 알려주는 3가지 + 새로 추가된 2가지
tf = tb(s, 6.9, 2.9, 5.5, 4.1)
outputs = [
    ("최적 취침시간", "농도가 1.0mg/L 아래로 떨어지는 가장 이른 시각"),
    ("총 수면시간", "수면 주기 90분 × 5번 = 7시간 30분"),
    ("기상시간", "취침시간 + 7시간 30분"),
]
for i, (t, d) in enumerate(outputs):
    p = addp(tf, "", size=20, space_after=3)
    r = p.add_run(); r.text = f"{i+1}  "; set_run(r, 20, True, AMBER)
    r = p.add_run(); r.text = t; set_run(r, 20, True, WHITE)
    addp(tf, "      " + d, size=14, color=BEIGE, space_after=12)
p = addp(tf, "", size=14, space_after=3)
r = p.add_run(); r.text = "NEW  "; set_run(r, 14, True, AMBER)
r = p.add_run(); r.text = "내 몸 기준 비교 + 연구 근거 섹션"; set_run(r, 14, True, WHITE)

# ============ 6~8. 주요 기능 (공통 템플릿) ============
def feature_slide(num, head, bullets, shot_fn, annos, label="05 · FEATURES"):
    """annos: [(x_frac, y_frac, label)] 스크린샷 위 앰버 번호 (미사용, 호환용)"""
    s = prs.slides.add_slide(BLANK); bg(s, CREAM)
    section_head(s, label, head)
    # 왼쪽 스크린샷
    img_x, img_y, img_w = 0.9, 2.35, 5.3
    pic = s.shapes.add_picture(f"{SHOT}/{shot_fn}", Inches(img_x), Inches(img_y), width=Inches(img_w))
    # 주의: 스크린샷 위 번호 주석은 글씨를 덮어 삭제함
    # 오른쪽: 큰 번호 + 불릿
    tf = tb(s, 6.9, 2.15, 2.2, 1.3)
    para(tf, num, size=100, bold=True, color=RGBColor(0xE3, 0xD2, 0xB8), space_after=0)
    tf = tb(s, 6.9, 3.45, 5.5, 3.3)
    for i, (bt, bd) in enumerate(bullets):
        p = addp(tf, "", size=19, space_after=3)
        r = p.add_run(); r.text = f"{i+1}  "; set_run(r, 19, True, AMBER)
        r = p.add_run(); r.text = bt; set_run(r, 19, True, ESPRESSO)
        addp(tf, "      " + bd, size=15, color=MILK, space_after=12)
    return s

s = feature_slide("01", "마신 커피가 수면 시간을 정한다", [
    ("최적 취침시간", "농도가 1.0mg/L 아래로 떨어지는 가장 이른 시각"),
    ("총 수면시간", "수면 주기 90분 × 5번 = 7시간 30분"),
    ("기상시간", "취침시간 + 7시간 30분으로 자동 계산"),
], "shot_plan.png", [(0.5, 0.35), (0.5, 0.6), (0.5, 0.85)], label="05 · FEATURES")
progress(s, 7)

s = feature_slide("02", "누르는 순간 기록 끝", [
    ("원탭 기록", "커피 칩을 누르면 지금 시각으로 바로 기록"),
    ("직접 입력", "마신 시간과 카페인 함량(mg)을 직접 입력"),
    ("시간 수정", "기록된 시간도 나중에 눌러서 수정 가능"),
], "shot_record.png", [(0.3, 0.35), (0.7, 0.62), (0.5, 0.85)], label="06 · FEATURES")
progress(s, 8)

s = feature_slide("03", "숫자가 아니라 답을 보여준다", [
    ("상태 히어로", '"지금 마셔도 돼요" · "최적 시각까지 30분 남았어요"'),
    ("취침 시 예상", "취침 시간의 예상 농도와 수면 판정"),
    ("오늘의 타임라인", "기상·최적·마감·취침을 한 줄로"),
], "shot_verdict.png", [(0.5, 0.4), (0.5, 0.75)], label="07 · FEATURES")
progress(s, 9)

# 8번: 그래프 클라이맥스 — 전체 폭
s = prs.slides.add_slide(BLANK); bg(s, CREAM)
section_head(s, "08 · FEATURES", "밀어서 보는 내 몸의 하루")
pic = s.shapes.add_picture(f"{SHOT}/shot_graph.png", Inches(0.9), Inches(2.35), width=Inches(5.6))
ph = pic.height / EMU
# 신호색 범례 (핀 대신)
tf = tb(s, 6.9, 2.5, 5.5, 4.0)
para(tf, "신호는 두 가지 색으로", size=22, bold=True, color=ESPRESSO, space_after=14)
lg = [(AMBER, "최적 시각", "마셔도 되는 타이밍"), (MIDNIGHT, "마감 시각", "이제 그만 마실 때")]
for col, t, d in lg:
    dot(s, 6.9, 3.6 + lg.index((col, t, d)) * 1.15, 0.28, col)
    tf2 = tb(s, 7.35, 3.52 + lg.index((col, t, d)) * 1.15, 4.9, 1.0)
    para(tf2, t, size=18, bold=True, color=ESPRESSO, space_after=2)
    addp(tf2, d, size=14, color=MILK)
tf = tb(s, 6.9, 6.0, 5.5, 1.0)
para(tf, "좌우 드래그 · 점을 터치하면 그 시각의 농도", size=15, color=MILK)
progress(s, 10)

# ============ 11. 개인 맞춤화 ============
s = prs.slides.add_slide(BLANK); bg(s, CREAM)
section_head(s, "09 · PERSONAL", "같은 커피, 몸마다 다르게 남습니다")
s.shapes.add_picture(f"{SHOT}/shot_compare.png", Inches(0.9), Inches(2.35), width=Inches(5.3))
tf = tb(s, 6.9, 2.5, 5.5, 4.0)
for i, (bt, bd) in enumerate([
    ("체수분량 (Watson 1980)", "키·몸무게·성별·나이로 카페인이 퍼지는 부피 계산"),
    ("대사 속도 선택", "빠름 3.5시간 · 보통 5시간 · 느림 7시간 (CYP1A2 유전자 근거)"),
    ("비교로 증명", "같은 150mg, 취침 시 농도가 대사 속도만 바꿔도 1.0 → 1.9mg/L"),
]):
    p = addp(tf, "", size=19, space_after=3)
    r = p.add_run(); r.text = f"{i+1}  "; set_run(r, 19, True, AMBER)
    r = p.add_run(); r.text = bt; set_run(r, 19, True, ESPRESSO)
    addp(tf, "      " + bd, size=15, color=MILK, space_after=12)
progress(s, 11)

# ============ 12. 차별성 ============
s = prs.slides.add_slide(BLANK); bg(s, CREAM)
section_head(s, "10 · DIFFERENT", "카페인 계산기와 수면 계산기, 따로 있었습니다")
headers = ["", "일반 카페인 계산기", "수면 계산기", "CaffeinGO"]
rows = [
    ("알려주는 것", "취침 시 남은 양", "90분 주기 기상 시간", "취침·기상·총 수면시간 플랜"),
    ("개인화", "반감기 5시간 고정", "없음", "Watson + 대사 속도"),
    ("카페인 + 수면", "수면 계획 없음", "카페인 반영 안 함", "곡선과 주기를 하나로 연결"),
]
col_x = [0.9, 3.0, 6.0, 9.0]
col_w = [2.1, 3.0, 3.0, 3.43]
for j, h in enumerate(headers):
    fill = AMBER if j == 3 else BEIGE
    tc = ESPRESSO if j == 3 else MILK
    box = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(col_x[j]), Inches(2.5), Inches(col_w[j]), Inches(0.6))
    box.fill.solid(); box.fill.fore_color.rgb = fill; box.line.fill.background()
    tf = tb(s, col_x[j] + 0.15, 2.62, col_w[j] - 0.3, 0.4)
    para(tf, h, size=15, bold=True, color=tc, align=PP_ALIGN.CENTER)
for i, row in enumerate(rows):
    y = 3.1 + i * 1.15
    for j, cell in enumerate(row):
        fill = WHITE if i % 2 == 0 else RGBColor(0xF3, 0xE8, 0xD8)
        if j == 3:
            fill = RGBColor(0xFB, 0xEF, 0xD7)
        box = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(col_x[j]), Inches(y), Inches(col_w[j]), Inches(1.15))
        box.fill.solid(); box.fill.fore_color.rgb = fill; box.line.fill.background()
        tf = tb(s, col_x[j] + 0.15, y + 0.18, col_w[j] - 0.3, 0.85)
        para(tf, cell, size=14.5, bold=(j == 0 or j == 3), color=ESPRESSO if j in (0, 3) else MILK, align=PP_ALIGN.CENTER)
tf = tb(s, 0.9, 6.75, 11.5, 0.5)
para(tf, "차별점은 기능이 아니라, 카페인과 수면을 하나로 연결했다는 점입니다.", size=15, bold=True, color=MIDNIGHT, align=PP_ALIGN.CENTER)
progress(s, 12)

# ============ 13. 핵심 기술 ============
s = prs.slides.add_slide(BLANK); bg(s, CREAM)
section_head(s, "11 · TECH", "약동학 모델을 단순화해 브라우저에서 바로 계산")
steps = [
    ("Watson 공식", "체수분량 추정", "baseline"),
    ("45분 흡수", "선형 상승", "rise"),
    ("반감기 감쇠", "지수 하락 · 대사 속도 반영", "decay"),
    ("합산", "여러 잔 기여도", "multi"),
]
for i, (t, d, kind) in enumerate(steps):
    x = 0.75 + i * 3.15
    box = rrect(s, x, 2.5, 2.85, 3.1, WHITE, radius=0.15)
    # 미니 곡선 썸네일
    if kind == "baseline":
        ln = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x + 0.5), Inches(2.85), Inches(1.85), Pt(4))
        ln.fill.solid(); ln.fill.fore_color.rgb = CARAMEL; ln.line.fill.background()
    elif kind == "rise":
        p0 = (int((x+0.5)*EMU), int(3.55*EMU))
        p1 = (int((x+2.35)*EMU), int(2.85*EMU))
        fb = s.shapes.build_freeform(p0[0], p0[1], scale=1.0)
        fb.add_line_segments([p1], close=False)
        sh = fb.convert_to_shape(); sh.fill.background(); sh.line.color.rgb = AMBER; sh.line.width = Pt(4)
    elif kind == "decay":
        draw_curve(s, x + 0.5, 2.85, 1.85, 0.7, AMBER, lw=4, peak_x=0.05, decay=2.5, rise_end=0.05)
    else:
        draw_curve(s, x + 0.5, 2.85, 1.85, 0.7, AMBER, lw=3, peak_x=0.3, decay=3.0)
        draw_curve(s, x + 0.5, 2.85, 1.85, 0.7, MIDNIGHT, lw=3, peak_x=0.45, decay=2.0)
    tf = tb(s, x + 0.2, 3.75, 2.45, 1.7)
    para(tf, t, size=18, bold=True, color=ESPRESSO, align=PP_ALIGN.CENTER, space_after=4)
    addp(tf, d, size=14, color=MILK, align=PP_ALIGN.CENTER)
    if i < 3:
        ar = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(x + 2.85), Inches(3.9), Inches(0.3), Inches(0.22))
        ar.fill.solid(); ar.fill.fore_color.rgb = CARAMEL; ar.line.fill.background()
# 실제 계산 예시
tf = tb(s, 0.9, 6.05, 11.5, 1.0)
para(tf, "예: 16:00 아메리카노 1잔  →  최적 취침 +1일 04:05, 기상 +1일 11:35", size=16, bold=True, color=MIDNIGHT, align=PP_ALIGN.CENTER, space_after=6)
addp(tf, "대사 속도만 바꿔도 취침 시 농도가 1.0 → 1.9mg/L로 달라집니다", size=15, color=MILK, align=PP_ALIGN.CENTER)
progress(s, 13)

# ============ 11. 시스템 아키텍처 ============
s = prs.slides.add_slide(BLANK); bg(s, CREAM)
section_head(s, "12 · ARCHITECTURE", "구조는 단순하게")
layers = [
    ("프레젠테이션 · HTML / CSS", "상태 히어로 · 그래프 · 타임라인 · 기록 칩", BEIGE, ESPRESSO),
    ("비즈니스 로직 · JavaScript", "계산 엔진 · 렌더링 엔진", CARAMEL, WHITE),
    ("데이터 · localStorage", "프로필 · 섭취 기록 JSON", ESPRESSO, WHITE),
]
for i, (t, d, fill, tc) in enumerate(layers):
    y = 2.45 + i * 1.42
    # 커피잔 단면 느낌: 아래로 갈수록 살짝 좁게
    inset = i * 0.35
    box = rrect(s, 2.2 + inset, y, 7.6 - inset * 2, 1.22, fill, radius=0.12)
    tf = tb(s, 2.2 + inset + 0.4, y + 0.22, 6.8 - inset * 2, 0.85)
    para(tf, t, size=16, bold=True, color=tc, space_after=3)
    addp(tf, d, size=13, color=tc if tc == WHITE else MILK)
# 화살표 2개 (오른쪽)
tf = tb(s, 10.3, 2.6, 2.2, 3.6)
para(tf, "↓  기록 → 저장", size=14, bold=True, color=CARAMEL, align=PP_ALIGN.CENTER, space_after=24)
addp(tf, "↑  불러오기 → 그래프", size=14, bold=True, color=MIDNIGHT, align=PP_ALIGN.CENTER, space_after=24)
addp(tf, "서버 없음\n외부 라이브러리 없음\n단일 HTML 파일", size=13, color=MILK, align=PP_ALIGN.CENTER)
progress(s, 14)

# ============ 12. 개발 방법 ============
s = prs.slides.add_slide(BLANK); bg(s, CREAM)
section_head(s, "13 · PROCESS", "문서 먼저가 아니라, 동작하는 것 먼저")
# 4단계 원형 순환
cx, cy, rad = 3.6, 4.35, 1.75
cyc = [("만들고", 0), ("써보고", 1), ("고치고", 2), ("문서화", 3)]
import math as _m
for i, (t, _) in enumerate(cyc):
    ang = -_m.pi / 2 + i * _m.pi / 2
    bx = cx + _m.cos(ang) * rad - 0.85
    by = cy + _m.sin(ang) * rad - 0.42
    box = rrect(s, bx, by, 1.7, 0.84, WHITE, line=CARAMEL, radius=0.3)
    tf = box.text_frame; tf.word_wrap = True
    para(tf, t, size=17, bold=True, color=ESPRESSO, align=PP_ALIGN.CENTER)
    # 화살표 (다음 단계로)
    nang = -_m.pi / 2 + (i + 0.5) * _m.pi / 2
    ax = cx + _m.cos(nang) * (rad + 0.12)
    ay = cy + _m.sin(nang) * (rad + 0.12)
    dot(s, ax - 0.07, ay - 0.07, 0.14, CARAMEL)
dot(s, cx - 0.42, cy - 0.42, 0.84, AMBER)
tf = tb(s, cx - 0.42, cy - 0.28, 0.84, 0.6)
para(tf, "반복", size=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
# 실제로 바뀐 것들
tf = tb(s, 6.9, 2.5, 5.5, 3.8)
para(tf, "실제로 바뀐 것들", size=22, bold=True, color=ESPRESSO, space_after=14)
for t in ["대사 속도 선택 + 내 몸 기준 비교", "연구 근거 섹션 + 출처 목록", "다른 계산기와 비교표", "최적 수면 플랜 · 시간 수정 · 취침 시 예상"]:
    addp(tf, "▸  " + t, size=16, color=ESPRESSO, space_after=10)
addp(tf, "피드백을 받으면, 근거부터 다시 세우고 고친다", size=14, bold=True, color=MILK, space_after=4)
progress(s, 15)

# ============ 13. 시연 (다크) ============
s = prs.slides.add_slide(BLANK); bg(s, ESPRESSO)
# LIVE DEMO 태그
dot(s, 0.9, 0.62, 0.22, AMBER)
tf = tb(s, 1.25, 0.55, 3.0, 0.4)
para(tf, "LIVE DEMO", size=16, bold=True, color=WHITE, space_after=2)
tf = tb(s, 0.9, 1.15, 11.5, 1.0)
para(tf, "이어서 실제 사이트를 시연하겠습니다", size=34, bold=True, color=WHITE)
# 필름 스트립
shots = [("shot_record.png", "기록"), ("shot_plan.png", "수면 플랜"), ("shot_compare.png", "내 몸 기준"), ("shot_why.png", "근거")]
verbs = ["기록", "수면 플랜", "내 몸 기준", "근거"]
for i, (fn, vb) in enumerate(shots):
    x = 0.9 + i * 3.1
    # 2.7x2.3 박스 안에 비율 유지하며 맞춤, 열 중앙 정렬 → 라벨 겹침 방지
    from PIL import Image as _Img
    _iw, _ih = _Img.open(f"{SHOT}/{fn}").size
    _sc = min(2.7 / _iw, 2.3 / _ih)
    _w, _h = _iw * _sc, _ih * _sc
    s.shapes.add_picture(f"{SHOT}/{fn}", Inches(x + (2.7 - _w) / 2), Inches(2.7 + (2.3 - _h) / 2), width=Inches(_w), height=Inches(_h))
    tf = tb(s, x, 5.35, 2.7, 0.6)
    para(tf, f"0{i+1}  {vb}", size=18, bold=True, color=AMBER, align=PP_ALIGN.CENTER)

# ============ 14. 향후 계획 ============
s = prs.slides.add_slide(BLANK); bg(s, CREAM)
section_head(s, "15 · ROADMAP", "지금은 시작 단계, 이렇게 키워갑니다")
# 굵은 화살표
ar = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(0.9), Inches(3.3), Inches(11.53), Inches(0.5))
ar.fill.solid(); ar.fill.fore_color.rgb = BEIGE; ar.line.fill.background()
# 단기 (실선 카드)
tf = tb(s, 0.9, 2.35, 5.0, 0.6)
para(tf, "단기", size=18, bold=True, color=CARAMEL)
for i, (t, d) in enumerate([("날짜별 기록 분리", "다음 날에도 전날 기록이 남는 문제 해결"), ("마감 시각 알림", "마셔도 되는 마지막 시각에 알림")]):
    y = 4.15 + i * 1.45
    box = rrect(s, 0.9, y, 5.0, 1.2, WHITE, radius=0.12)
    tf = tb(s, 1.25, y + 0.18, 4.3, 0.9)
    para(tf, t, size=17, bold=True, color=ESPRESSO, space_after=2)
    addp(tf, d, size=13.5, color=MILK)
# 중기 (점선 카드)
tf = tb(s, 7.4, 2.35, 5.0, 0.6)
para(tf, "중기", size=18, bold=True, color=MIDNIGHT)
for i, (t, d) in enumerate([("대사 속도 자가진단", "몇 가지 질문으로 내 대사 속도를 추정해 자동 설정"), ("수면·카페인 리포트", "일주일 기록을 모아 패턴 분석")]):
    y = 4.15 + i * 1.45
    box = rrect(s, 7.4, y, 5.0, 1.2, CREAM, line=MIDNIGHT, radius=0.12)
    try:
        box.line.dash_style = 7  # 점선
    except Exception:
        pass
    tf = tb(s, 7.75, y + 0.18, 4.3, 0.9)
    para(tf, t, size=17, bold=True, color=ESPRESSO, space_after=2)
    addp(tf, d, size=13.5, color=MILK)
progress(s, 17)

# ============ 15. 감사합니다 (다크) ============
s = prs.slides.add_slide(BLANK); bg(s, ESPRESSO)
# 앰버 곡선이 내려가 기준선에 닿음
draw_curve(s, 1.0, 1.2, 11.3, 2.6, AMBER, lw=5, peak_x=0.12, decay=3.4, rise_end=0.06)
# 기준선
bl = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.0), Inches(3.72), Inches(11.3), Pt(2))
bl.fill.solid(); bl.fill.fore_color.rgb = CARAMEL; bl.line.fill.background()
# 닿는 지점에 미드나이트 점
dot(s, 10.9, 3.62, 0.3, MIDNIGHT)
tf = tb(s, 1.5, 4.3, 10.3, 1.4)
para(tf, "Q&A", size=96, bold=True, color=WHITE, align=PP_ALIGN.CENTER, space_after=4)
tf = tb(s, 1.5, 5.9, 10.3, 1.0)
para(tf, "감사합니다", size=26, color=CARAMEL, align=PP_ALIGN.CENTER, space_after=4)
addp(tf, "CaffeinGO · 팀 세송이", size=15, color=MILK, align=PP_ALIGN.CENTER)

prs.save(OUT)
print("saved:", OUT)
print("slides:", len(prs.slides.__iter__.__self__._sldIdLst))
