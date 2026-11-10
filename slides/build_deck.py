"""발표 자료 생성 스크립트.

실행: ..\\.venv\\Scripts\\python.exe build_deck.py  → slides/credit_presentation.pptx
숫자는 results.md, 그림은 figures/ 에서 가져온다. 장표를 고치면 이 파일을 다시 실행한다.
"""
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

ROOT = Path(__file__).resolve().parent.parent
FIG = ROOT / "figures"
OUT = ROOT / "slides" / "credit_presentation.pptx"

# 디자인 토큰
FONT = "Malgun Gothic"
INK = RGBColor(0x1F, 0x29, 0x33)      # 본문
MUTED = RGBColor(0x61, 0x6E, 0x7C)    # 보조 글씨
LINE = RGBColor(0xD9, 0xDE, 0xE3)     # 구분선
PANEL = RGBColor(0xF4, 0xF6, 0xF8)    # 카드 배경
BLUE = RGBColor(0x2F, 0x6F, 0xB3)     # 강조
RED = RGBColor(0xD6, 0x45, 0x45)      # 연체
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
NAVY = RGBColor(0x14, 0x2B, 0x45)     # 표지 배경

LABS = ["데이터 이해", "전처리", "특성·검증", "평가·비교", "자동화·앙상블", "MLP"]

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
BLANK = prs.slide_layouts[6]


# ---------- 공통 도구 ----------
def _font(run, size, bold=False, color=INK):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = FONT
    rpr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):  # 한글 글꼴 지정
        el = rpr.find(qn(tag))
        if el is None:
            el = rpr.makeelement(qn(tag), {})
            rpr.append(el)
        el.set("typeface", FONT)


def text(slide, x, y, w, h, lines, size=16, color=INK, bold=False, align=PP_ALIGN.LEFT,
         anchor=MSO_ANCHOR.TOP, spacing=1.15):
    """lines: 문자열 또는 [(문자열, {옵션})] 목록. 한 줄 안에서 [(조각, 옵션), ...]도 가능."""
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    if isinstance(lines, str):
        lines = [lines]
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = spacing
        pieces = line if isinstance(line, list) else [(line, {})]
        for piece, opt in pieces:
            r = p.add_run()
            r.text = piece
            _font(r, opt.get("size", size), opt.get("bold", bold), opt.get("color", color))
        if isinstance(line, list) and line and line[0][1].get("space_before"):
            p.space_before = Pt(line[0][1]["space_before"])
    return box


def rect(slide, x, y, w, h, fill, shape=MSO_SHAPE.RECTANGLE, line=None):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(1)
    s.shadow.inherit = False
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = 0.08
    return s


def notes(slide, script):
    slide.notes_slide.notes_text_frame.text = script.strip()


def lab_header(slide, lab_no, chapter, title):
    """Lab 장표 공통 머리: 단계 표시줄 + 장 이름 + 결론형 제목."""
    # 단계 표시줄 (오른쪽 위)
    x0 = 12.73 - (5 * 1.09 + 1.02)  # 오른쪽 끝을 본문 여백(12.73in)에 맞춤
    for i, name in enumerate(LABS, start=1):
        active = i == lab_no
        done = i < lab_no
        fill = BLUE if active else (RGBColor(0xC9, 0xD8, 0xEA) if done else PANEL)
        chip = rect(slide, x0 + (i - 1) * 1.09, 0.38, 1.02, 0.36, fill, MSO_SHAPE.ROUNDED_RECTANGLE)
        tf = chip.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = f"Lab {i}"
        _font(r, 11, bold=active, color=WHITE if active else MUTED)
    text(slide, 0.6, 0.4, 5.6, 0.35, [[(f"LAB {lab_no}", {"bold": True, "color": BLUE}),
                                       (f"   {chapter}", {"color": MUTED})]], size=13)
    text(slide, 0.6, 0.95, 12.1, 1.0, title, size=26, bold=True)
    rect(slide, 0.6, 1.95, 12.13, 0.02, LINE)


def next_question(slide, question, lab_next, tracker):
    """아래 띠: 다음 질문 + 성능 추적 점수."""
    rect(slide, 0.6, 6.42, 8.6, 0.68, BLUE, MSO_SHAPE.ROUNDED_RECTANGLE)
    text(slide, 0.85, 6.42, 8.2, 0.68,
         [[("다음 질문  ", {"bold": True, "color": WHITE}), (question, {"color": WHITE}),
           (f"   → {lab_next}", {"bold": True, "color": RGBColor(0xC9, 0xD8, 0xEA)})]],
         size=15, anchor=MSO_ANCHOR.MIDDLE)
    rect(slide, 9.4, 6.42, 3.33, 0.68, PANEL, MSO_SHAPE.ROUNDED_RECTANGLE)
    text(slide, 9.6, 6.42, 3.0, 0.68, tracker, size=12, anchor=MSO_ANCHOR.MIDDLE, spacing=1.0)


def finding(slide, x, y, w, no, head, body, accent=BLUE):
    rect(slide, x, y, 0.07, 1.12, accent)
    text(slide, x + 0.25, y - 0.02, w - 0.25, 0.4,
         [[(f"발견 {no}  ", {"bold": True, "color": accent}), (head, {"bold": True})]], size=16)
    text(slide, x + 0.25, y + 0.42, w - 0.25, 0.75, body, size=13, color=MUTED, spacing=1.2)


# ---------- 1. 표지 ----------
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, 13.333, 7.5, NAVY)
rect(s, 0.9, 2.35, 0.09, 1.9, RED)
text(s, 1.25, 2.2, 11, 0.5, "알고리즘분석 과제 발표 · Give Me Some Credit", size=16, color=RGBColor(0xC9, 0xD8, 0xEA))
text(s, 1.25, 2.75, 11, 1.6, ["Lab 1 → Lab 6으로 풀어 본", "신용위험 예측"], size=40, bold=True, color=WHITE, spacing=1.05)
text(s, 1.25, 5.6, 11, 0.5, "발표자: (이름) · (학번)", size=16, color=RGBColor(0xC9, 0xD8, 0xEA))
notes(s, """
안녕하세요. 저는 Lab 1부터 Lab 6까지 실습한 방법을 그대로 따라가면서,
신용위험 예측 문제를 단계별로 풀어 본 과정을 발표하겠습니다.
""")

# ---------- 2. 문제 정의 ----------
s = prs.slides.add_slide(BLANK)
text(s, 0.6, 0.4, 6, 0.35, [[("문제 정의", {"bold": True, "color": BLUE})]], size=13)
text(s, 0.6, 0.95, 12.1, 1.0, "15만 명의 재무 데이터로 '2년 안에 연체할 고객'을 미리 찾는다", size=26, bold=True)
rect(s, 0.6, 1.95, 12.13, 0.02, LINE)

cards = [
    ("데이터", ["Give Me Some Credit", "(Kaggle, 2011)", "150,000명 × 9열", "결측치 처리 완료본"]),
    ("목표", ["2년 내 90일 이상 연체 여부", "(SeriousDlqin2yrs)", "이진 분류 · 연체 확률 예측"]),
    ("조건", ["Lab 1–6 방법을 따라 진행", "MLP 반드시 사용", "15만 건을 어떻게 학습할지"]),
]
for i, (head, body) in enumerate(cards):
    x = 0.6 + i * 4.1
    rect(s, x, 2.3, 3.93, 1.95, PANEL, MSO_SHAPE.ROUNDED_RECTANGLE)
    text(s, x + 0.3, 2.45, 3.4, 0.4, head, size=15, bold=True, color=BLUE)
    text(s, x + 0.3, 2.9, 3.4, 1.3, body, size=14, spacing=1.15)

text(s, 0.6, 4.55, 6, 0.4, "입력 변수 8개", size=15, bold=True)
feats = ["리볼빙 사용 비율", "나이", "30–59일 연체 횟수", "부채 비율",
         "월 소득", "부양가족 수", "연체 이력 (통합)", "대출 보유 (통합)"]
for i, f in enumerate(feats):
    x = 0.6 + (i % 4) * 3.07
    y = 5.0 + (i // 4) * 0.55
    rect(s, x, y, 2.93, 0.43, WHITE, MSO_SHAPE.ROUNDED_RECTANGLE, line=LINE)
    text(s, x, y, 2.93, 0.43, f, size=13, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text(s, 0.6, 6.2, 12.1, 0.3, "변수 설명 출처: Han et al., NOTE, Scientific Reports 14 (2024) 26070", size=11, color=MUTED)

# 진행 순서
text(s, 0.6, 6.6, 12.1, 0.45,
     [[("진행 순서   ", {"bold": True, "color": BLUE})] +
      sum([[(f"Lab {i} {n}", {}), ("  ›  " if i < 6 else "", {"color": MUTED})] for i, n in enumerate(LABS, 1)], [])],
     size=13)
notes(s, """
사용한 데이터는 Kaggle의 Give Me Some Credit으로, 15만 명의 재무 정보가 들어 있습니다.
목표는 이 고객이 2년 안에 90일 이상 연체할지를 예측하는 이진 분류입니다.
입력 변수는 사용 비율, 나이, 연체 횟수, 부채 비율, 소득 등 8개입니다.
과제 조건에 맞춰 Lab 1부터 6까지 순서대로 진행했고, 각 단계에서 발견한 문제가 다음 단계의 질문이 되도록 구성했습니다.
""")

# ---------- 3. Lab 1 ----------
s = prs.slides.add_slide(BLANK)
lab_header(s, 1, "데이터 로딩 · 기술통계", "연체 고객은 6.7%뿐, 그리고 데이터 곳곳에 극단값이 숨어 있다")
s.shapes.add_picture(str(FIG / "lab1_class_balance.png"), Inches(0.6), Inches(2.25), width=Inches(5.9))
finding(s, 6.95, 2.35, 5.8, 1, "극단적 불균형 (1 : 14)",
        ["연체 10,026명 vs 정상 139,974명.", "모두 '정상'이라 찍어도 정확도 93.3%"], accent=RED)
finding(s, 6.95, 3.68, 5.8, 2, "비율인데 비율이 아닌 값",
        ["사용 비율 최댓값 50,708 · 부채 비율 1 초과 35,137행", "소득 결측을 6638·6746으로 채운 흔적과 겹침"])
finding(s, 6.95, 5.01, 5.8, 3, "극단값이 관계를 가린다",
        ["사용 비율과 연체의 상관계수", "Pearson -0.002  →  순위(Spearman) 0.24"])
next_question(s, "이 값들을 그대로 써도 될까?", "Lab 2",
              [[("기준선 ", {"bold": True})], [("정확도 93.3% · AUC 0.50", {"color": MUTED})]])
notes(s, """
Lab 1에서 배운 대로 데이터를 불러와 기술통계부터 봤습니다.
먼저 파일에 헤더 줄이 있어서, Lab 코드를 그대로 쓰면 첫 줄이 데이터로 섞이는 문제가 있었고 이를 고쳤습니다.
가장 큰 발견은 클래스 불균형입니다. 연체 고객은 6.7%뿐이라, 모두 정상이라고 찍어도 정확도가 93%가 나옵니다.
두 번째로, 비율이어야 할 사용 비율이 최대 5만이 넘고, 부채 비율도 1을 넘는 행이 3만 5천 개였습니다.
이 행들은 소득이 6638, 6746으로 채워진 행과 겹쳐서, 빈 소득을 채운 흔적으로 보입니다.
세 번째로, 사용 비율은 일반 상관계수로는 연체와 무관해 보이지만 순위로 보면 0.24로 꽤 관련이 있습니다.
극단값이 관계를 가리고 있다는 뜻입니다. 그래서 다음 질문은, 이 값들을 그대로 써도 되는가입니다.
""")

prs.save(OUT)
print("saved", OUT)
