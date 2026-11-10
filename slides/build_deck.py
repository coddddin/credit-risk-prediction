"""발표 자료 생성 스크립트 (디자인: A 미니멀 + B 다크 포인트).

1) ..\\.venv\\Scripts\\python.exe build_deck.py   → slides/_draft.pptx (장표 내용)
2) finalize.ps1                                  → slides/credit_presentation.pptx
   (모핑 전환 + 이징 진입 애니메이션 + 글꼴 포함 저장)

숫자는 results.md, Lab 원본 캡처는 figures/lab_captures (lab_capture.py) 에서 가져온다.
도형 이름 규칙
  "!!이름"   : 장표 사이에서 같은 이름끼리 모핑으로 이어짐
  "anim{n}_" : 진입 애니메이션 순서 n (같은 n은 함께, 조금씩 시차를 두고 나타남)
"""
import glob
import os
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd
from matplotlib import font_manager
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt
from pygments.lexers import PythonLexer
from pygments.token import Comment, Keyword, Name, Number, Operator, Punctuation, String

ROOT = Path(__file__).resolve().parent.parent
FIG = ROOT / "figures"
CAP = FIG / "lab_captures"
CHART = FIG / "slide_charts"
CHART.mkdir(parents=True, exist_ok=True)
DRAFT = ROOT / "slides" / "_draft.pptx"
DATA = ROOT / "data" / "givemesomecredit.csv"

# ---------------- 디자인 토큰 ----------------
C = dict(
    bg="#FFFFFF", ink="#191F28", sub="#4E5968", mute="#8B95A1", line="#E5E8EB", card="#F2F4F6",
    blue="#3182F6", blue_pale="#E8F3FF", red="#F04452", red_pale="#FFEEF0",
    violet="#7C5CFF", violet_pale="#F1EEFF",                       # B에서 가져온 보조 강조색
    dark="#0B0E14", dark2="#141922", dark_line="#262D3A", dark_ink="#E8ECF2", dark_mute="#8A94A6",
)
CODE = {"kw": "#C792EA", "fn": "#82AAFF", "str": "#C3E88D", "com": "#6B7689", "num": "#F78C6C",
        "op": "#89DDFF", "txt": "#E8ECF2"}
WEIGHT = {"light": "Pretendard Light", "regular": "Pretendard", "medium": "Pretendard Medium",
          "semibold": "Pretendard SemiBold", "bold": "Pretendard", "extrabold": "Pretendard ExtraBold"}
MONO = "Consolas"
LABS = ["데이터 이해", "전처리", "특성·검증", "평가·비교", "자동화·앙상블", "MLP"]
# Lecture 6의 Predictive Analytics 프로젝트 6단계와 Lab 대응
STAGES = [("문제 정의", "Define Problem"), ("데이터 분석", "Analyse Data"), ("데이터 준비", "Prepare Data"),
          ("알고리즘 평가", "Evaluate Algorithms"), ("성능 개선", "Improve Results"), ("결과 제시", "Present Results")]
LAB_STAGES = {1: [2], 2: [2, 3], 3: [3, 4], 4: [4], 5: [5], 6: [4, 5]}

for p in glob.glob(os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\Windows\Fonts\Pretendard-*.ttf")):
    font_manager.fontManager.addfont(p)
plt.rcParams["font.family"] = "Pretendard"
plt.rcParams["axes.unicode_minus"] = False


def rgb(h):
    h = h.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


# ---------------- 기본 도구 ----------------
prs = Presentation()
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)


MORPH_XML = """
<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <mc:Choice xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" Requires="p159">
    <p:transition xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" spd="slow" p14:dur="900">
      <p159:morph option="byObject"/>
    </p:transition>
  </mc:Choice>
  <mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback>
</mc:AlternateContent>"""


def add_morph(s):
    """모핑 전환 (같은 '!!이름' 도형끼리 이징으로 이어짐). 구버전 PowerPoint에서는 페이드로 대체."""
    from lxml import etree
    sld = s._element
    anchor = sld.find(qn("p:clrMapOvr"))
    if anchor is None:
        anchor = sld.find(qn("p:cSld"))
    anchor.addnext(etree.fromstring(MORPH_XML))


def new_slide(bg=C["bg"]):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = rgb(bg)
    if len(prs.slides) > 1:
        add_morph(s)
    return s


def set_font(run, size, weight="regular", color=C["ink"], latin=None, spacing=None, italic=False):
    run.font.size = Pt(size)
    run.font.bold = weight == "bold"
    run.font.italic = italic
    run.font.color.rgb = rgb(color)
    face = WEIGHT[weight]
    run.font.name = latin or face
    rpr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        el = rpr.find(qn(tag))
        if el is None:
            el = rpr.makeelement(qn(tag), {})
            rpr.append(el)
        el.set("typeface", face)
    if spacing is not None:
        rpr.set("spc", str(int(spacing * 100)))


def text(s, x, y, w, h, lines, size=16, color=C["ink"], weight="regular", align="l", anchor="t",
         line=1.2, spacing=None, name=None, latin=None):
    """lines: 문자열 | [줄] ; 줄 = 문자열 | [(조각, {size,color,weight,latin})]"""
    box = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        box.name = name
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[anchor]
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, ln in enumerate([lines] if isinstance(lines, str) else lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[align]
        p.line_spacing = line
        for piece, o in (ln if isinstance(ln, list) else [(ln, {})]):
            r = p.add_run()
            r.text = piece
            set_font(r, o.get("size", size), o.get("weight", weight), o.get("color", color),
                     o.get("latin", latin), o.get("spacing", spacing))
    return box


def rect(s, x, y, w, h, fill, radius=0.0, border=None, border_w=1.0, name=None):
    r = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE,
                           Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        r.name = name
    if fill is None:
        r.fill.background()
    else:
        r.fill.solid()
        r.fill.fore_color.rgb = rgb(fill)
    if border:
        r.line.color.rgb = rgb(border)
        r.line.width = Pt(border_w)
    else:
        r.line.fill.background()
    r.shadow.inherit = False
    if radius:
        r.adjustments[0] = min(0.5, radius / min(w, h))
    return r


def oval(s, x, y, d, fill, name=None):
    o = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(d), Inches(d))
    o.fill.solid()
    o.fill.fore_color.rgb = rgb(fill)
    o.line.fill.background()
    o.shadow.inherit = False
    if name:
        o.name = name
    return o


def group(s, shapes, name):
    g = s.shapes.add_group_shape(shapes)
    g.name = name
    return g


def notes(s, script):
    s.notes_slide.notes_text_frame.text = " ".join(script.split())


# ---------------- B 스타일 부품: 코드 창 / 실행 결과 창 ----------------
def _token_color(tt):
    if tt in Comment or tt in String.Doc:
        return CODE["com"]
    if tt in String:
        return CODE["str"]
    if tt in Number:
        return CODE["num"]
    if tt in Keyword:
        return CODE["kw"]
    if tt in Name.Builtin or tt in Name.Function:
        return CODE["fn"]
    if tt in Operator or tt in Punctuation:
        return CODE["op"]
    return CODE["txt"]


def code_window(s, x, y, w, h, filename, code, size=12.5, highlight=(), anim=None, lh=18):
    """어두운 에디터 창. highlight: 강조할 줄 번호(0부터)."""
    parts = []
    parts.append(rect(s, x, y, w, h, C["dark"], radius=0.16))
    parts.append(rect(s, x, y, w, 0.42, C["dark2"], radius=0.16))
    parts.append(rect(s, x, y + 0.26, w, 0.16, C["dark2"]))
    for i, col in enumerate(["#FF5F57", "#FEBC2E", "#28C840"]):
        parts.append(oval(s, x + 0.2 + i * 0.2, y + 0.15, 0.12, col))
    parts.append(text(s, x + 0.9, y + 0.07, w - 1.1, 0.3, filename, size=11, color=C["dark_mute"],
                      weight="medium", latin=MONO, anchor="m"))
    lines = code.strip("\n").split("\n")
    top = y + 0.62
    step = lh / 72
    for i in highlight:
        parts.append(rect(s, x + 0.12, top + i * step - 0.02, w - 0.24, step + 0.02, "#1D2A45", radius=0.05))
        parts.append(rect(s, x + 0.12, top + i * step - 0.02, 0.05, step + 0.02, C["blue"]))
    box = s.shapes.add_textbox(Inches(x + 0.32), Inches(top), Inches(w - 0.5), Inches(h - 0.75))
    tf = box.text_frame
    tf.word_wrap = False
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.line_spacing = Pt(lh)
        toks = list(PythonLexer().get_tokens(ln)) if ln.strip() else [(None, " ")]
        for tt, val in toks:
            val = val.rstrip("\n")
            if not val:
                continue
            r = p.add_run()
            r.text = val
            set_font(r, size, "regular", CODE["txt"] if tt is None else _token_color(tt), latin=MONO)
    parts.append(box)
    return group(s, parts, f"anim{anim}_code" if anim else "code")


def console(s, x, y, w, h, title, rows, size=12.5, anim=None, lh=18):
    """실행 결과 창. rows: [(문자열, 색키)] — 색키 'in'(입력), 'out', 'hi'(강조), 'red'."""
    parts = [rect(s, x, y, w, h, C["dark"], radius=0.16)]
    parts.append(text(s, x + 0.3, y + 0.18, w - 0.6, 0.3,
                      [[("●  ", {"color": "#28C840", "size": 9}), (title, {})]],
                      size=11, color=C["dark_mute"], weight="semibold"))
    color = {"in": C["violet"], "out": C["dark_ink"], "hi": "#FFD166", "red": "#FF6B7A", "dim": C["dark_mute"]}
    box = s.shapes.add_textbox(Inches(x + 0.3), Inches(y + 0.62), Inches(w - 0.5), Inches(h - 0.75))
    tf = box.text_frame
    tf.word_wrap = False
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, row in enumerate(rows):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.line_spacing = Pt(lh)
        for piece, key in (row if isinstance(row, list) else [row]):
            r = p.add_run()
            r.text = piece if piece else " "
            set_font(r, size, "regular", color[key], latin=MONO)
    parts.append(box)
    return group(s, parts, f"anim{anim}_console" if anim else "console")


def capture_card(s, x, y, w, img, label, anim=None):
    """Lab PDF 원본 캡처 카드."""
    from PIL import Image
    iw, ih = Image.open(img).size
    ph = w * 0.9 * ih / iw
    h = ph + 0.64
    parts = [rect(s, x, y, w, h, C["card"], radius=0.18),
             text(s, x + 0.25, y + 0.14, w - 0.5, 0.3,
                  [[("Lab 자료  ", {"color": C["blue"], "weight": "bold"}), (label, {"color": C["mute"]})]],
                  size=12, weight="medium")]
    pic = s.shapes.add_picture(str(img), Inches(x + w * 0.05), Inches(y + 0.5), width=Inches(w * 0.9))
    parts.append(pic)
    group(s, parts, f"anim{anim}_capture" if anim else "capture")
    return h


# ---------------- 공통 머리 / 하단 ----------------
def stage_track(s, active, x_right=12.53, seg=0.98, gap=0.06):
    """오른쪽 위 AI Modeling 6단계 표시줄. active: 강조할 단계 번호 목록."""
    x0 = x_right - (6 * seg + 5 * gap)
    done_until = min(active) - 1
    for i, (ko, _) in enumerate(STAGES, start=1):
        on = i in active
        fill = C["blue"] if on else ("#B9D5FC" if i <= done_until else C["line"])
        x = x0 + (i - 1) * (seg + gap)
        rect(s, x, 0.42, seg, 0.08, fill, radius=0.04, name=f"!!stage{i}")
        text(s, x, 0.56, seg, 0.25, f"{i} {ko}", size=10, align="c", name=f"!!stagelabel{i}",
             color=C["blue"] if on else C["mute"], weight="bold" if on else "medium")


def lab_header(s, lab, part, chapter, title):
    """part: 1=코드, 2=결과, 3=발견·결론"""
    stage_track(s, LAB_STAGES[lab])
    text(s, 0.8, 0.55, 5.4, 0.35,
         [[(f"Lab {lab}", {"color": C["blue"], "weight": "bold"}),
           (f"  {chapter}", {"color": C["mute"], "weight": "medium"}),
           (f"   {['코드', '결과', '발견 · 결론'][part - 1]} {part}/3", {"color": C["violet"], "weight": "semibold"})]],
         size=14, name="!!kicker")
    text(s, 0.8, 0.98, 11.8, 0.9, title, size=30, weight="extrabold", spacing=-1, name="!!title")


def next_bar(s, question, nxt, tracker_label, tracker_value, anim=None):
    parts = [rect(s, 0.8, 6.42, 8.9, 0.64, C["dark"], radius=0.32),
             text(s, 1.15, 6.42, 8.4, 0.64,
                  [[("NEXT  ", {"color": C["violet"], "weight": "bold", "spacing": 1.5}),
                    (question, {"color": C["dark_ink"], "weight": "semibold"}),
                    (f"   →  {nxt}", {"color": "#82AAFF", "weight": "bold"})]], size=16, anchor="m"),
             rect(s, 9.9, 6.42, 2.63, 0.64, C["blue_pale"], radius=0.32),
             text(s, 9.9, 6.42, 2.63, 0.64,
                  [[(f"{tracker_label}  ", {"color": C["sub"], "weight": "semibold", "size": 12}),
                    (tracker_value, {"color": C["blue"], "weight": "extrabold"})]], size=16, align="c", anchor="m")]
    group(s, parts, f"anim{anim}_next" if anim else "!!next")


def conclusion_card(s, x, y, w, h, lines, anim=None):
    """이번 Lab 결과가 '연체 예측'이라는 목표에 주는 의미."""
    parts = [rect(s, x, y, w, h, C["violet_pale"], radius=0.22),
             rect(s, x + 0.3, y + 0.32, 0.06, 0.36, C["violet"]),
             text(s, x + 0.5, y + 0.28, w - 0.8, 0.45,
                  [[("목표에 대한 결론", {}), ("   연체 예측 모델에 주는 의미", {"size": 12, "color": C["mute"], "weight": "medium"})]],
                  size=16, weight="bold", color=C["violet"]),
             text(s, x + 0.3, y + 0.85, w - 0.6, h - 1.0, lines, size=14, color=C["ink"], weight="medium", line=1.45)]
    group(s, parts, f"anim{anim}_conclusion" if anim else "conclusion")


def pill(s, x, y, w, label, fill, color, size=13, weight="semibold", h=0.42):
    return [rect(s, x, y, w, h, fill, radius=h / 2),
            text(s, x, y, w, h, label, size=size, color=color, weight=weight, align="c", anchor="m")]


# =====================================================================
# 1. 표지
# =====================================================================
s = new_slide()
text(s, 0.9, 1.75, 6.5, 0.45, "알고리즘분석 과제 발표", size=17, color=C["blue"], weight="semibold", name="!!kicker")
text(s, 0.9, 2.3, 6.8, 2.2, ["Lab 1부터 6까지,", "신용위험을 예측하다"], size=48, weight="extrabold",
     line=1.12, spacing=-1.5, name="!!title")
text(s, 0.9, 4.45, 6.5, 0.45, "Give Me Some Credit · 150,000명의 재무 데이터", size=17, color=C["sub"], weight="medium",
     name="anim1_sub")
group(s, pill(s, 0.9, 5.15, 1.45, "Lab 1 → 6", C["blue_pale"], C["blue"])
      + pill(s, 2.48, 5.15, 1.75, "150,000명", C["red_pale"], C["red"])
      + pill(s, 4.36, 5.15, 1.25, "MLP", C["violet_pale"], C["violet"]), "anim2_pills")
rect(s, 0.9, 6.35, 6.0, 0.012, C["line"])
text(s, 0.9, 6.5, 6, 0.35, "발표자  (이름) · (학번)", size=13, color=C["mute"], weight="medium")
code_window(s, 7.75, 1.9, 4.85, 3.35, "story.py", '''
# 발표의 흐름
for lab in range(1, 7):
    finding  = explore(data, lab)
    question = next_question(finding)
    print(f"Lab {lab}: {question}")

model = MLP(hidden=[12, 8])  # Lab 6
model.fit(X, y, batch_size=512)
''', size=12, highlight=(1, 2, 3), anim=3, lh=19)
notes(s, "안녕하세요. 저는 Lab 1부터 Lab 6까지 실습한 방법을 그대로 따라가면서, 신용위험 예측 문제를 단계별로 풀어 본 과정을 발표하겠습니다.")

# =====================================================================
# 2. 문제 정의
# =====================================================================
s = new_slide()
stage_track(s, [1])
text(s, 0.8, 0.55, 5.4, 0.35, "문제 정의", size=14, color=C["blue"], weight="semibold", name="!!kicker")
text(s, 0.8, 0.98, 11.8, 0.9, "2년 안에 연체할 고객을 미리 찾을 수 있을까?", size=34, weight="extrabold",
     spacing=-1, name="!!title")
stats = [("150,000", "명의 고객 데이터", C["ink"]), ("8 → 1", "입력 변수 8개로 연체 여부 1개 예측", C["ink"]),
         ("6단계", "AI Modeling 과정을 Lab 1–6으로", C["blue"])]
for i, (num, cap, col) in enumerate(stats):
    x = 0.8 + i * 3.95
    group(s, [rect(s, x, 2.1, 3.73, 1.7, C["card"], radius=0.22),
              text(s, x + 0.35, 2.32, 3.2, 0.8, num, size=40, color=col, weight="extrabold", spacing=-1),
              text(s, x + 0.35, 3.18, 3.2, 0.4, cap, size=14, color=C["sub"], weight="medium")], f"anim{i + 1}_stat{i}")
text(s, 0.8, 4.2, 11.8, 0.4, [[("AI Modeling · Predictive Analytics 6단계", {"weight": "bold"}),
                               ("   이 순서를 Lab 1–6으로 밟아 간다 (Lecture 6)", {"color": C["mute"], "weight": "medium"})]],
     size=15)
stage_labs = {1: ["이 장표"], 2: ["Lab 1", "Lab 2"], 3: ["Lab 2", "Lab 3"], 4: ["Lab 3", "Lab 4", "Lab 6"],
              5: ["Lab 5", "Lab 6"], 6: ["결론"]}
road = []
for i, (ko, en) in enumerate(STAGES, start=1):
    x = 0.8 + (i - 1) * 1.98
    on = i == 1
    road += [rect(s, x, 4.75, 1.86, 1.85, C["blue_pale"] if on else C["card"], radius=0.18),
             text(s, x + 0.2, 4.9, 1.6, 0.3, f"STEP {i}", size=10.5, color=C["blue"] if on else C["mute"], weight="bold",
                  spacing=1),
             text(s, x + 0.2, 5.17, 1.6, 0.4, ko, size=16, weight="bold"),
             text(s, x + 0.2, 5.5, 1.6, 0.3, en, size=10.5, color=C["mute"], weight="medium")]
    for j, lab_name in enumerate(stage_labs[i]):
        px = x + 0.2 + j * 0.53
        road += pill(s, px, 5.98, 0.5 if lab_name.startswith("Lab") else 0.8, lab_name.replace("Lab ", "Lab"),
                     C["dark"] if lab_name.startswith("Lab") else C["violet_pale"],
                     C["dark_ink"] if lab_name.startswith("Lab") else C["violet"], size=10.5, h=0.32)
    if i < 6:
        road.append(text(s, x + 1.86, 5.45, 0.12, 0.3, "›", size=14, color=C["mute"], align="c"))
group(s, road, "anim4_road")
text(s, 0.8, 6.75, 11.8, 0.3, "데이터: Kaggle Give Me Some Credit (2011)  ·  변수 설명: Han et al., NOTE, Scientific Reports 14 (2024)",
     size=11, color=C["mute"])
notes(s, """사용한 데이터는 Kaggle의 Give Me Some Credit으로, 15만 명의 재무 정보가 들어 있습니다.
목표는 이 고객이 2년 안에 90일 이상 연체할지를 예측하는 이진 분류입니다.
진행은 Lecture 6에서 배운 AI 모델링 6단계, 즉 문제 정의, 데이터 분석, 데이터 준비, 알고리즘 평가, 성능 개선, 결과 제시를 따릅니다. 이 단계를 Lab 1부터 6까지 순서대로 밟아 가며, Lab마다 코드, 결과, 그리고 그 결과가 연체 예측이라는 목표에 주는 의미를 정리하겠습니다.""")

# =====================================================================
# Lab 1 - (1) 코드
# =====================================================================
s = new_slide()
lab_header(s, 1, 1, "데이터 로딩 · 기술통계", "Lab 1의 순서 그대로, 15만 건을 불러와 살펴봤다")

# Lab 1 자료의 진행 순서 (스테퍼)
steps = [("불러오기", "read_csv", "header=0 추가"), ("미리보기", "head(20)", "아래 표"),
         ("크기", "shape", "(150000, 9)"), ("자료형", "dtypes", "9열 모두 숫자"),
         ("기술통계", "describe()", "revol 최대 50,708"), ("클래스 분포", "groupby", "연체 10,026명"),
         ("상관 · 왜도", "corr · skew", "왜도 최대 127")]
sw, gap = 1.57, 0.12
parts = [text(s, 0.8, 1.98, 11.7, 0.3, [[("Lab 1 자료의 진행 순서", {"weight": "bold"}),
                                         ("   → 실제 데이터에 같은 순서로 실습한 결과", {"color": C["mute"], "weight": "medium"})]],
              size=13)]
for i, (ko, fn, res) in enumerate(steps):
    x = 0.8 + i * (sw + gap)
    parts += [rect(s, x, 2.35, sw, 1.18, C["blue_pale"], radius=0.16),
              oval(s, x + 0.15, 2.47, 0.3, C["blue"]),
              text(s, x + 0.15, 2.47, 0.3, 0.3, str(i + 1), size=11, color="#FFFFFF", weight="bold", align="c", anchor="m"),
              text(s, x + 0.52, 2.47, sw - 0.6, 0.3, ko, size=13, weight="bold", anchor="m"),
              text(s, x + 0.15, 2.84, sw - 0.25, 0.25, fn, size=11, color=C["blue"], weight="medium", latin=MONO),
              rect(s, x + 0.12, 3.14, sw - 0.24, 0.3, C["bg"], radius=0.08),
              text(s, x + 0.12, 3.14, sw - 0.24, 0.3, res, size=10.5, color=C["ink"], weight="semibold", align="c",
                   anchor="m")]
group(s, parts, "anim1_steps")

# 근거가 된 Lab 자료 캡처
h1 = capture_card(s, 0.8, 3.7, 5.6, CAP / "lab1_read_csv.png", "p.4 · 1 불러오기 (Pima 데이터 예시)", anim=2)

# 실제 데이터 미리보기 (head)
head = pd.read_csv(DATA, names=["target", "revol", "age", "late30", "debt", "income", "depend", "defaulted", "loans"],
                   header=0).head(3)
ty = 3.7 + h1 + 0.12
tparts = [rect(s, 0.8, ty, 5.6, 7.05 - ty, C["card"], radius=0.18),
          text(s, 1.05, ty + 0.12, 5.1, 0.3, [[("실제 데이터  ", {"color": C["red"], "weight": "bold"}),
                                               ("data.head() · 15만 행 중 첫 3행", {"color": C["mute"]})]], size=12,
               weight="medium")]
cols = list(head.columns)
cw = [0.5, 0.58, 0.4, 0.55, 0.58, 0.6, 0.6, 0.72, 0.5]  # 열 너비 (합 5.03in)
tbl = s.shapes.add_table(len(head) + 1, len(cols), Inches(1.08), Inches(ty + 0.48), Inches(sum(cw)),
                         Inches(0.22 * (len(head) + 1))).table
for j, c in enumerate(cols):
    tbl.columns[j].width = Inches(cw[j])
for i in range(len(head) + 1):
    tbl.rows[i].height = Inches(0.22)
    for j, c in enumerate(cols):
        cell = tbl.cell(i, j)
        cell.margin_left = cell.margin_right = Inches(0.03)
        cell.margin_top = cell.margin_bottom = Inches(0.01)
        cell.fill.solid()
        cell.fill.fore_color.rgb = rgb(C["blue_pale"] if i == 0 else C["bg"])
        v = c if i == 0 else head.iloc[i - 1, j]
        val = v if i == 0 else (f"{v:.3f}" if isinstance(v, float) else f"{v:,}")
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = str(val)
        is_target = j == 0 and i > 0 and v == 1
        set_font(r, 9.5, "bold" if (i == 0 or is_target) else "regular",
                 C["blue"] if i == 0 else (C["red"] if is_target else C["ink"]), latin=MONO)
tbl_frame = s.shapes[-1]
tparts.append(tbl_frame)
group(s, tparts, "anim3_head")

# 같은 순서로 실습한 실행 결과
cmd = lambda code, note: [(">>> ", "in"), (code.ljust(38), "out"), (f"# {note}", "dim")]
console(s, 6.62, 3.7, 5.91, 3.35, "실행 결과 · 01_lab1_loading_stats.ipynb", [
    cmd("data.shape", "3 크기"), [("(150000, 9)", "hi")],
    cmd("data.dtypes.value_counts()", "4 자료형"), [("int64 7  ·  float64 2", "hi")],
    cmd("data.describe().loc['max']", "5 기술통계"), [("revol 50708  ·  debt 329664  ·  income 3008750", "hi")],
    cmd("data.groupby('target').size()", "6 클래스"), [("0  139974  ·  1  ", "hi"), ("10026", "red")],
    cmd("data.corr()['target'][1:].nlargest(2)", "7 상관"), [("defaulted 0.314  ·  late30 0.126", "hi")],
    cmd("data.skew().nlargest(3)", "7 왜도"), [("income 127.4  ·  revol 97.6  ·  debt 95.2", "hi")],
], size=11, anim=4, lh=15.5)
notes(s, """Lab 1 자료는 데이터를 불러와서 미리보기, 크기, 자료형, 기술통계, 클래스 분포, 상관관계와 왜도 순서로 살펴봅니다.
저는 이 일곱 단계를 그대로 실제 신용 데이터에 적용했습니다. 위쪽 칸마다 실제로 나온 결과를 적었습니다.
15만 행 9열, 모든 열이 숫자이고, 기술통계에서 사용 비율 최댓값이 5만이 넘는 이상한 값이 보였습니다. 연체 고객은 1만 26명입니다.
왼쪽 아래는 실제 데이터의 첫 세 행, 오른쪽은 각 단계의 실행 결과입니다.
Lab 자료는 헤더가 없는 파일을 기준으로 해서, 이 데이터에는 header=0 하나만 더했습니다.""")

# =====================================================================
# Lab 1 - (2) 결과
# =====================================================================
s = new_slide()
lab_header(s, 1, 2, "데이터 로딩 · 기술통계", "연체 고객은 6.7%뿐, 데이터 곳곳엔 극단값이 숨어 있다")
parts = [rect(s, 0.8, 2.15, 5.3, 4.9, C["card"], radius=0.25),
         text(s, 1.2, 2.45, 4.6, 0.4, "연체 고객 비율", size=16, color=C["sub"], weight="semibold"),
         text(s, 1.2, 2.8, 4.7, 1.5, [[("6.7", {}), ("%", {"size": 46})]], size=100, color=C["red"],
              weight="extrabold", spacing=-3),
         rect(s, 1.2, 4.6, 4.5, 0.3, C["line"], radius=0.15),
         rect(s, 1.2, 4.6, 4.5 * 0.0668, 0.3, C["red"], radius=0.15),
         text(s, 1.2, 5.1, 4.6, 0.4, [[("연체 ", {}), ("10,026", {"weight": "bold", "color": C["ink"]}),
                                      ("명   ·   정상 ", {}), ("139,974", {"weight": "bold", "color": C["ink"]}), ("명", {})]],
              size=14, color=C["sub"], weight="medium"),
         rect(s, 1.2, 5.75, 4.5, 0.95, C["bg"], radius=0.16),
         text(s, 1.45, 5.75, 4.1, 0.95, [[("모두 '정상'이라 찍어도 ", {}), ("정확도 93.3%", {"weight": "bold", "color": C["red"]})],
                                         "→ 정확도로는 평가할 수 없다"], size=14, color=C["sub"], weight="medium",
              anchor="m", line=1.3)]
group(s, parts, "anim1_stat")
console(s, 6.35, 2.15, 6.18, 4.9, "describe에서 이상한 값 파고들기 · 실제 데이터", [
    [("# 비율이어야 하는데 1을 넘는 값", "dim")],
    [(">>> ", "in"), ("(data.revol > 1).sum()", "out")],
    [("3321", "hi"), ("      ← 사용 비율 최대 50,708", "dim")],
    [(">>> ", "in"), ("(data.debt > 1).sum()", "out")],
    [("35137", "hi")],
    [("", "out")],
    [("# 소득에 유난히 자주 나오는 값", "dim")],
    [(">>> ", "in"), ("data.income.value_counts().head(2)", "out")],
    [("6638 16634  ·  6746 13103", "hi"), ("  ← 채운 값", "dim")],
    [(">>> ", "in"), ("(filled & (data.debt > 1)).sum()", "out")],
    [("27904", "red"), ("      ← 부채 '금액'이 들어감", "dim")],
    [("", "out")],
    [(">>> ", "in"), ("data.late30.isin([96, 98]).sum()", "out")],
    [("269", "hi"), ("        ← 불가능한 연체 횟수", "dim")],
], size=12, anim=2, lh=18)
notes(s, """결과입니다. 가장 큰 발견은 불균형입니다. 연체 고객은 1만 26명, 6.7%뿐이라 모두 정상이라고 찍어도 정확도가 93%가 나옵니다.
오른쪽은 기술통계에서 이상해 보인 값을 실제 데이터로 파고든 결과입니다. 비율이어야 할 사용 비율과 부채 비율이 1을 넘는 행이 각각 3천여 개, 3만 5천 개였습니다.
소득은 6638과 6746이 유난히 많은데, 이 행들 중 2만 7천여 개는 부채 비율도 1을 넘었습니다. 빈 소득을 채우면서 부채 비율 자리에 금액이 들어간 것으로 보입니다.
연체 횟수에는 96, 98 같은 불가능한 값도 있었습니다.""")

# =====================================================================
# Lab 1 - (3) 발견 · 결론
# =====================================================================
df = pd.read_csv(DATA)
y = "SeriousDlqin2yrs"
label = {"RevolvingUtilizationOfUnsecuredLines": "사용 비율", "age": "나이", "NumberOfTime30-59DaysPastDueNotWorse": "30–59일 연체",
         "DebtRatio": "부채 비율", "MonthlyIncome": "월 소득", "NumberOfDependents": "부양가족", "CombinedDefaulted": "연체 이력",
         "CombinedCreditLoans": "대출 보유"}
corr = pd.DataFrame({"pearson": df.corr()[y], "spearman": df.corr(method="spearman")[y]}).drop(y).rename(index=label)
corr = corr.reindex(corr["spearman"].abs().sort_values().index)
fig, ax = plt.subplots(figsize=(6.2, 3.35), facecolor=C["card"])
ax.set_facecolor(C["card"])
ys = range(len(corr))
ax.barh([v + 0.19 for v in ys], corr["pearson"], height=0.36, color="#C5CCD3", label="Pearson (원래 값)")
ax.barh([v - 0.19 for v in ys], corr["spearman"], height=0.36, color=C["blue"], label="Spearman (순위)")
ax.set_yticks(list(ys))
ax.set_yticklabels(corr.index, fontsize=12)
ax.axvline(0, color="#8B95A1", lw=0.8)
ax.tick_params(length=0, labelsize=11, colors="#4E5968")
for t in ax.get_yticklabels():
    if t.get_text() == "사용 비율":
        t.set_color(C["red"]); t.set_fontweight("bold")
ax.grid(axis="x", color="#E5E8EB", lw=0.8)
ax.set_axisbelow(True)
for sp in ax.spines.values():
    sp.set_visible(False)
ax.legend(frameon=False, loc="lower right", fontsize=11)
fig.tight_layout()
fig.savefig(CHART / "lab1_corr.png", dpi=220, facecolor=C["card"])
plt.close(fig)

s = new_slide()
lab_header(s, 1, 3, "데이터 로딩 · 기술통계", "극단값이 '사용 비율'과 연체의 관계를 가리고 있었다")
group(s, [rect(s, 0.8, 2.15, 6.0, 4.05, C["card"], radius=0.25),
          text(s, 1.1, 2.33, 5.4, 0.35, [[("연체와의 상관계수  ", {"weight": "bold", "color": C["ink"]}),
                                         ("사용 비율: -0.002 → 0.24", {"color": C["red"], "weight": "bold"})]], size=13),
          s.shapes.add_picture(str(CHART / "lab1_corr.png"), Inches(1.0), Inches(2.72), width=Inches(5.6))],
      "anim1_chart")
finds = [("01", "불균형 1 : 14", "정확도 대신 다른 지표가 필요하다", C["red"]),
         ("02", "이상치 · 채운 값 · 특수 코드", "비율 > 1, 소득 6638·6746, 연체 96·98", C["blue"]),
         ("03", "극단값이 관계를 가린다", "순위로 보면 사용 비율은 3번째로 중요", C["blue"])]
parts = []
for i, (no, head, body, col) in enumerate(finds):
    yy = 2.15 + i * 0.72
    parts += [text(s, 7.1, yy, 0.6, 0.4, no, size=17, color=col, weight="extrabold"),
              text(s, 7.65, yy, 4.9, 0.4, head, size=16, weight="bold"),
              text(s, 7.65, yy + 0.34, 4.9, 0.35, body, size=12.5, color=C["sub"], weight="medium")]
group(s, parts, "anim2_finds")
conclusion_card(s, 7.05, 4.4, 5.48, 1.8, [
    [("① ", {"color": C["violet"], "weight": "bold"}), ("평가 기준: 정확도 ✕ → 연체를 골라내는 AUC", {})],
    [("② ", {"color": C["violet"], "weight": "bold"}), ("모델에 넣기 전, 극단값부터 처리해야 한다", {})],
], anim=3)
next_bar(s, "이 값들을 그대로 써도 될까?", "Lab 2", "기준선", "AUC 0.50", anim=4)
notes(s, """마지막으로 상관관계입니다. 회색이 일반 상관계수, 파란색이 순위 상관계수인데, 사용 비율은 일반 상관계수로는 0에 가깝지만 순위로 보면 0.24로 세 번째로 강합니다.
극단값 몇 개가 관계를 가리고 있었다는 뜻입니다.
연체 예측이라는 목표로 보면 결론은 두 가지입니다. 첫째, 이 모델은 정확도가 아니라 연체 고객을 얼마나 잘 골라내는지, 즉 AUC로 평가해야 합니다. 지금 기준선은 0.5입니다. 둘째, 사용 비율 같은 핵심 신호를 모델이 쓰게 하려면 극단값부터 처리해야 합니다.
그래서 다음 질문은, 이 값들을 그대로 써도 되는가입니다.""")

prs.save(DRAFT)
print("saved", DRAFT)
