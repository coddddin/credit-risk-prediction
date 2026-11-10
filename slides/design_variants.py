"""디자인 시안 3종 (A 미니멀 / B 다크 테크 / C 벤토 그리드) 생성.

실행: ..\\.venv\\Scripts\\python.exe design_variants.py → slides/variants/deck_A|B|C.pptx
시안이 정해지면 선택된 스타일을 build_deck.py에 반영한다.
"""
import glob
import os
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "slides" / "variants"
OUT.mkdir(parents=True, exist_ok=True)

# Pretendard를 matplotlib에도 등록 (장표와 그래프 글꼴 통일)
for p in glob.glob(os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\Windows\Fonts\Pretendard-*.ttf")):
    font_manager.fontManager.addfont(p)
plt.rcParams["font.family"] = "Pretendard"
plt.rcParams["axes.unicode_minus"] = False

# Lab 1 숫자 (results.md 기준)
N_OK, N_BAD = 139_974, 10_026
LABS = ["데이터 이해", "전처리", "특성·검증", "평가·비교", "자동화·앙상블", "MLP"]
NOTES = {
    "cover": "안녕하세요. 저는 Lab 1부터 Lab 6까지 실습한 방법을 그대로 따라가면서, 신용위험 예측 문제를 단계별로 풀어 본 과정을 발표하겠습니다.",
    "problem": ("사용한 데이터는 Kaggle의 Give Me Some Credit으로, 15만 명의 재무 정보가 들어 있습니다. "
                "목표는 이 고객이 2년 안에 90일 이상 연체할지를 예측하는 이진 분류입니다. "
                "과제 조건에 맞춰 Lab 1부터 6까지 순서대로 진행했고, 각 단계에서 발견한 문제가 다음 단계의 질문이 되도록 구성했습니다."),
    "lab1": ("Lab 1에서 배운 대로 데이터를 불러와 기술통계부터 봤습니다. 파일에 헤더 줄이 있어 Lab 코드를 그대로 쓰면 첫 줄이 데이터로 섞이는 문제를 먼저 고쳤습니다. "
             "가장 큰 발견은 불균형입니다. 연체 고객은 6.7%뿐이라 모두 정상이라고 찍어도 정확도가 93%입니다. "
             "또 비율이어야 할 사용 비율이 최대 5만이 넘고, 부채 비율이 1을 넘는 행이 3만 5천 개였는데, 소득을 6638, 6746으로 채운 행과 겹쳤습니다. "
             "마지막으로 사용 비율은 일반 상관계수로는 연체와 무관해 보이지만 순위로 보면 0.24입니다. 극단값이 관계를 가리고 있다는 뜻입니다. "
             "그래서 다음 질문은, 이 값들을 그대로 써도 되는가입니다."),
}

WEIGHT = {"light": "Pretendard Light", "regular": "Pretendard", "medium": "Pretendard Medium",
          "semibold": "Pretendard SemiBold", "bold": "Pretendard", "extrabold": "Pretendard ExtraBold"}


def hexrgb(h):
    h = h.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


class Deck:
    def __init__(self):
        self.prs = Presentation()
        self.prs.slide_width, self.prs.slide_height = Inches(13.333), Inches(7.5)

    def slide(self, bg):
        s = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        s.background.fill.solid()
        s.background.fill.fore_color.rgb = hexrgb(bg)
        return s

    @staticmethod
    def _run_font(run, size, weight, color, spacing=None):
        run.font.size = Pt(size)
        run.font.bold = weight == "bold"
        run.font.color.rgb = hexrgb(color)
        face = WEIGHT[weight]
        run.font.name = face
        rpr = run._r.get_or_add_rPr()
        for tag in ("a:ea", "a:cs"):
            el = rpr.find(qn(tag))
            if el is None:
                el = rpr.makeelement(qn(tag), {})
                rpr.append(el)
            el.set("typeface", face)
        if spacing is not None:  # 자간 (1/100 pt)
            rpr.set("spc", str(int(spacing * 100)))

    def text(self, s, x, y, w, h, lines, size=16, color="#111111", weight="regular",
             align="l", anchor="t", line=1.2, spacing=None):
        """lines: 문자열 | [줄...] ; 줄 = 문자열 | [(조각, {size,color,weight})...]"""
        box = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
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
                self._run_font(r, o.get("size", size), o.get("weight", weight), o.get("color", color),
                               o.get("spacing", spacing))
        return box

    def rect(self, s, x, y, w, h, fill, radius=0.0, border=None, border_w=1.0):
        shape = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
        r = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
        if fill is None:
            r.fill.background()
        else:
            r.fill.solid()
            r.fill.fore_color.rgb = hexrgb(fill)
        if border:
            r.line.color.rgb = hexrgb(border)
            r.line.width = Pt(border_w)
        else:
            r.line.fill.background()
        r.shadow.inherit = False
        if radius:
            r.adjustments[0] = min(0.5, radius / min(w, h))
        return r

    def notes(self, s, key):
        s.notes_slide.notes_text_frame.text = NOTES[key]

    def save(self, name):
        path = OUT / name
        self.prs.save(path)
        return path


# =====================================================================
# A. 미니멀 (Toss / Apple)
# =====================================================================
def build_A():
    T = dict(bg="#FFFFFF", ink="#191F28", sub="#4E5968", mute="#8B95A1", line="#E5E8EB",
             card="#F2F4F6", blue="#3182F6", red="#F04452", pale="#E8F3FF")
    d = Deck()

    # 표지
    s = d.slide(T["bg"])
    d.text(s, 1.0, 2.15, 11, 0.5, "알고리즘분석 과제 발표", size=18, color=T["blue"], weight="semibold")
    d.text(s, 1.0, 2.75, 11.5, 2.2, ["Lab 1부터 6까지,", "신용위험을 예측하다"], size=54, color=T["ink"],
           weight="extrabold", line=1.12, spacing=-1.5)
    d.text(s, 1.0, 4.95, 11, 0.5, "Give Me Some Credit · 150,000명의 재무 데이터", size=18, color=T["sub"], weight="medium")
    d.rect(s, 1.0, 6.25, 11.33, 0.015, T["line"])
    d.text(s, 1.0, 6.45, 6, 0.4, "발표자  (이름) · (학번)", size=14, color=T["mute"], weight="medium")
    d.notes(s, "cover")

    # 문제 정의
    s = d.slide(T["bg"])
    d.text(s, 0.8, 0.6, 6, 0.4, "문제 정의", size=15, color=T["blue"], weight="semibold")
    d.text(s, 0.8, 1.05, 11.8, 1.0, "2년 안에 연체할 고객을 미리 찾을 수 있을까?", size=36, color=T["ink"],
           weight="extrabold", spacing=-1)
    stats = [("150,000", "명의 고객 데이터"), ("8", "개의 입력 변수"), ("6.7", "% 연체 고객 비율")]
    for i, (num, cap) in enumerate(stats):
        x = 0.8 + i * 3.95
        d.rect(s, x, 2.45, 3.7, 1.75, T["card"], radius=0.22)
        d.text(s, x + 0.35, 2.68, 3.2, 0.8, num, size=40, color=T["blue"] if i == 2 else T["ink"], weight="extrabold", spacing=-1)
        d.text(s, x + 0.35, 3.55, 3.2, 0.4, cap, size=15, color=T["sub"], weight="medium")
    d.text(s, 0.8, 4.6, 6, 0.4, "진행 순서", size=15, color=T["ink"], weight="semibold")
    for i, name in enumerate(LABS):
        x = 0.8 + i * 1.98
        d.rect(s, x, 5.1, 1.82, 0.06, T["blue"] if i == 0 else T["line"])
        d.text(s, x, 5.3, 1.82, 0.3, f"Lab {i + 1}", size=12, color=T["mute"], weight="semibold")
        d.text(s, x, 5.6, 1.82, 0.4, name, size=15, color=T["ink"], weight="semibold")
    d.text(s, 0.8, 6.7, 11.8, 0.3, "데이터: Kaggle Give Me Some Credit (2011) · 변수 설명: Han et al., NOTE, Sci. Rep. 14 (2024)",
           size=11, color=T["mute"])
    d.notes(s, "problem")

    # Lab 1
    s = d.slide(T["bg"])
    d.text(s, 0.8, 0.6, 6, 0.4, [[("Lab 1", {"color": T["blue"], "weight": "bold"}),
                                  ("   데이터 로딩 · 기술통계", {"color": T["mute"], "weight": "medium"})]], size=15)
    for i in range(6):  # 진행 막대
        d.rect(s, 9.03 + i * 0.6, 0.72, 0.52, 0.07, T["blue"] if i == 0 else T["line"], radius=0.03)
    d.text(s, 0.8, 1.05, 11.8, 1.0, "연체 고객은 6.7%뿐, 데이터 곳곳엔 극단값이 숨어 있다", size=32,
           color=T["ink"], weight="extrabold", spacing=-1)
    # 큰 숫자 + 막대
    d.rect(s, 0.8, 2.35, 5.4, 3.75, T["card"], radius=0.25)
    d.text(s, 1.2, 2.6, 4.8, 0.4, "연체 고객 비율", size=16, color=T["sub"], weight="semibold")
    d.text(s, 1.2, 2.95, 4.8, 1.5, [[("6.7", {}), ("%", {"size": 48})]], size=96, color=T["red"], weight="extrabold", spacing=-3)
    bar_w = 4.6
    d.rect(s, 1.2, 4.75, bar_w, 0.32, T["line"], radius=0.16)
    d.rect(s, 1.2, 4.75, bar_w * 0.067, 0.32, T["red"], radius=0.16)
    d.text(s, 1.2, 5.2, 4.8, 0.7, ["연체 10,026명  ·  정상 139,974명", "모두 '정상'이라 찍어도 정확도 93.3%"],
           size=14, color=T["sub"], weight="medium", line=1.35)
    # 발견 목록
    items = [("01", "극단적 불균형 (1 : 14)", "정확도로는 모델을 평가할 수 없다"),
             ("02", "비율인데 비율이 아닌 값", "사용 비율 최댓값 50,708 · 부채 비율 1 초과 35,137행"),
             ("03", "극단값이 관계를 가린다", "사용 비율의 상관계수  Pearson -0.002 → 순위 0.24")]
    for i, (no, head, body) in enumerate(items):
        y = 2.4 + i * 1.25
        d.text(s, 6.75, y, 0.7, 0.5, no, size=20, color=T["blue"], weight="extrabold")
        d.text(s, 7.45, y - 0.02, 5.2, 0.5, head, size=20, color=T["ink"], weight="bold")
        d.text(s, 7.45, y + 0.48, 5.2, 0.5, body, size=14, color=T["sub"], weight="medium")
        if i < 2:
            d.rect(s, 6.75, y + 1.07, 5.8, 0.012, T["line"])
    # 다음 질문
    d.rect(s, 0.8, 6.4, 8.5, 0.62, T["pale"], radius=0.31)
    d.text(s, 1.15, 6.4, 8.0, 0.62, [[("다음 질문   ", {"color": T["blue"], "weight": "bold"}),
                                      ("이 값들을 그대로 써도 될까?  →  Lab 2", {"color": T["ink"], "weight": "semibold"})]],
           size=16, anchor="m")
    d.text(s, 9.55, 6.4, 3.0, 0.62, [[("기준선  ", {"color": T["mute"], "weight": "semibold"}),
                                      ("AUC 0.50", {"color": T["ink"], "weight": "bold"})]], size=15, anchor="m", align="r")
    d.notes(s, "lab1")
    return d.save("deck_A_minimal.pptx")


# =====================================================================
# B. 다크 테크 (Linear / Vercel)
# =====================================================================
def dark_chart(path, T):
    fig, ax = plt.subplots(figsize=(5.2, 3.15), facecolor=T["card"])
    ax.set_facecolor(T["card"])
    bars = ax.bar(["정상", "연체"], [N_OK, N_BAD], color=[T["muted_bar"], T["red"]], width=0.5)
    for b, n in zip(bars, [N_OK, N_BAD]):
        ax.text(b.get_x() + b.get_width() / 2, n + 3000, f"{n:,}", ha="center", va="bottom",
                color=T["ink"], fontsize=13, fontweight="bold")
    ax.set_ylim(0, 165000)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"{v / 1000:.0f}K"))
    ax.tick_params(colors=T["mute"], labelsize=11, length=0)
    ax.grid(axis="y", color=T["border"], lw=0.8)
    ax.set_axisbelow(True)
    for sp in ax.spines.values():
        sp.set_visible(False)
    fig.tight_layout()
    fig.savefig(path, dpi=220, facecolor=T["card"])
    plt.close(fig)


def build_B():
    T = dict(bg="#0B0E14", card="#141922", border="#262D3A", ink="#E8ECF2", mute="#8A94A6",
             violet="#8B7CFF", cyan="#3DD6D0", red="#FF5C6C", muted_bar="#3A4356")
    d = Deck()

    # 표지
    s = d.slide(T["bg"])
    for i in range(1, 12):  # 은은한 격자선
        d.rect(s, i * 1.11, 0, 0.006, 7.5, "#121722")
    d.text(s, 1.0, 1.9, 6, 0.4, "ALGORITHM ANALYSIS · TERM PROJECT", size=13, color=T["violet"], weight="semibold", spacing=2)
    d.text(s, 1.0, 2.5, 11.5, 2.2, ["Lab 1 → Lab 6", "신용위험 예측 파이프라인"], size=52, color=T["ink"],
           weight="extrabold", line=1.1, spacing=-1.5)
    d.text(s, 1.0, 4.65, 11, 0.5, "Give Me Some Credit · 150,000 rows · binary classification", size=17, color=T["mute"], weight="medium")
    d.rect(s, 1.0, 6.2, 3.2, 0.5, T["card"], radius=0.25, border=T["border"])
    d.text(s, 1.0, 6.2, 3.2, 0.5, "발표자  (이름) · (학번)", size=13, color=T["ink"], weight="medium", align="c", anchor="m")
    d.notes(s, "cover")

    # 문제 정의
    s = d.slide(T["bg"])
    d.text(s, 0.8, 0.6, 6, 0.4, "PROBLEM", size=13, color=T["violet"], weight="semibold", spacing=2)
    d.text(s, 0.8, 1.0, 11.8, 1.0, "2년 안에 연체할 고객을 미리 찾는다", size=34, color=T["ink"], weight="extrabold", spacing=-1)
    cards = [("DATA", "Give Me Some Credit", "Kaggle 2011 · 150,000 × 9 · 결측치 처리 완료"),
             ("TARGET", "SeriousDlqin2yrs", "2년 내 90일 이상 연체 · 0 / 1"),
             ("CONSTRAINT", "Lab 1–6 · MLP 필수", "15만 건을 어떻게 학습할지 고민")]
    for i, (tag, head, body) in enumerate(cards):
        x = 0.8 + i * 3.95
        d.rect(s, x, 2.25, 3.75, 2.0, T["card"], radius=0.14, border=T["border"])
        d.text(s, x + 0.3, 2.45, 3.2, 0.3, tag, size=11, color=T["cyan"], weight="semibold", spacing=1.5)
        d.text(s, x + 0.3, 2.85, 3.2, 0.5, head, size=19, color=T["ink"], weight="bold")
        d.text(s, x + 0.3, 3.4, 3.2, 0.8, body, size=13, color=T["mute"], weight="medium", line=1.3)
    for i, name in enumerate(LABS):
        x = 0.8 + i * 1.98
        active = i == 0
        d.rect(s, x, 4.8, 1.84, 1.05, T["card"], radius=0.12, border=T["violet"] if active else T["border"],
               border_w=1.5 if active else 1.0)
        d.text(s, x + 0.2, 4.95, 1.5, 0.3, f"0{i + 1}", size=12, color=T["violet"] if active else T["mute"], weight="semibold")
        d.text(s, x + 0.2, 5.3, 1.6, 0.4, name, size=14, color=T["ink"], weight="semibold")
    d.text(s, 0.8, 6.7, 11.8, 0.3, "Source: Kaggle Give Me Some Credit (2011) · Features: Han et al., NOTE, Sci. Rep. 14 (2024)",
           size=11, color=T["mute"])
    d.notes(s, "problem")

    # Lab 1
    s = d.slide(T["bg"])
    for i in range(6):
        active = i == 0
        d.rect(s, 8.95 + i * 0.63, 0.55, 0.55, 0.36, T["violet"] if active else T["card"], radius=0.1,
               border=None if active else T["border"])
        d.text(s, 8.95 + i * 0.63, 0.55, 0.55, 0.36, f"0{i + 1}", size=11, color="#FFFFFF" if active else T["mute"],
               weight="semibold", align="c", anchor="m")
    d.text(s, 0.8, 0.6, 6, 0.4, "LAB 01 · LOADING & DESCRIPTIVE STATS", size=13, color=T["violet"], weight="semibold", spacing=1.5)
    d.text(s, 0.8, 1.05, 11.8, 1.0, "연체 고객은 6.7%뿐, 데이터 곳곳엔 극단값이 숨어 있다", size=30,
           color=T["ink"], weight="extrabold", spacing=-1)
    chart = OUT / "_B_class.png"
    dark_chart(chart, T)
    d.rect(s, 0.8, 2.2, 5.6, 4.0, T["card"], radius=0.14, border=T["border"])
    d.text(s, 1.1, 2.4, 5, 0.35, [[("CLASS BALANCE  ", {"color": T["cyan"], "weight": "semibold", "spacing": 1.5}),
                                  ("연체 6.7%", {"color": T["ink"], "weight": "bold"})]], size=12)
    s.shapes.add_picture(str(chart), Inches(1.0), Inches(2.85), width=Inches(5.2))
    items = [("극단적 불균형", "1 : 14", "모두 '정상'이라 찍어도 정확도 93.3%", T["red"]),
             ("비율이 아닌 비율", "50,708", "사용 비율 최댓값 · 부채 비율 1 초과 35,137행", T["violet"]),
             ("가려진 관계", "-0.002 → 0.24", "사용 비율 상관계수 Pearson → Spearman", T["cyan"])]
    for i, (head, big, body, col) in enumerate(items):
        y = 2.2 + i * 1.36
        d.rect(s, 6.65, y, 5.88, 1.2, T["card"], radius=0.14, border=T["border"])
        d.rect(s, 6.65, y + 0.25, 0.06, 0.7, col)
        d.text(s, 6.95, y + 0.18, 2.6, 0.4, head, size=14, color=T["mute"], weight="semibold")
        d.text(s, 6.95, y + 0.55, 2.8, 0.5, big, size=22, color=col, weight="extrabold")
        d.text(s, 9.55, y + 0.2, 2.8, 0.85, body, size=12.5, color=T["ink"], weight="medium", line=1.3, anchor="m")
    d.rect(s, 0.8, 6.45, 9.0, 0.6, T["violet"], radius=0.12)
    d.text(s, 1.1, 6.45, 8.6, 0.6, [[("NEXT  ", {"weight": "bold", "spacing": 1.5}),
                                     ("이 값들을 그대로 써도 될까?  →  Lab 02", {"weight": "semibold"})]],
           size=15, color="#FFFFFF", anchor="m")
    d.rect(s, 10.0, 6.45, 2.53, 0.6, T["card"], radius=0.12, border=T["border"])
    d.text(s, 10.0, 6.45, 2.53, 0.6, [[("BASELINE  ", {"color": T["mute"], "size": 11, "spacing": 1}),
                                       ("AUC 0.50", {"color": T["ink"], "weight": "bold"})]], size=14, align="c", anchor="m")
    d.notes(s, "lab1")
    return d.save("deck_B_dark.pptx")


# =====================================================================
# C. 벤토 그리드 (Notion / Figma Config)
# =====================================================================
def build_C():
    T = dict(bg="#F4F1EA", ink="#111111", sub="#55524C", mute="#8C877D", tile="#FFFFFF",
             peach="#FFD9C2", lav="#DCD6FF", mint="#CDEFD9", sky="#CFE6FF", lemon="#FFF1B8", red="#E5484D")
    d = Deck()

    # 표지
    s = d.slide(T["bg"])
    d.rect(s, 0.7, 0.7, 7.6, 6.1, T["ink"], radius=0.35)
    d.text(s, 1.2, 1.25, 6.8, 0.4, "알고리즘분석 과제 발표", size=16, color="#BFBAB0", weight="semibold")
    d.text(s, 1.2, 2.7, 6.8, 2.6, ["Lab 1 → 6", "신용위험", "예측 노트"], size=54, color="#FFFFFF", weight="extrabold",
           line=1.05, spacing=-2)
    d.text(s, 1.2, 6.0, 6.8, 0.4, "발표자  (이름) · (학번)", size=14, color="#BFBAB0", weight="medium")
    tiles = [(8.55, 0.7, 4.08, 2.95, T["peach"], "150,000", "명의 재무 데이터"),
             (8.55, 3.85, 1.94, 2.95, T["lav"], "8", "입력 변수"),
             (10.69, 3.85, 1.94, 2.95, T["mint"], "6", "단계의 Lab")]
    for x, y, w, h, c, num, cap in tiles:
        d.rect(s, x, y, w, h, c, radius=0.3)
        d.text(s, x + 0.35, y + h - 1.45, w - 0.5, 0.9, num, size=40, color=T["ink"], weight="extrabold", spacing=-1.5)
        d.text(s, x + 0.35, y + h - 0.6, w - 0.5, 0.4, cap, size=14, color=T["sub"], weight="semibold")
    d.notes(s, "cover")

    # 문제 정의
    s = d.slide(T["bg"])
    d.text(s, 0.7, 0.55, 6, 0.4, "문제 정의", size=15, color=T["mute"], weight="semibold")
    d.text(s, 0.7, 0.95, 12, 1.0, "2년 안에 연체할 고객을 미리 찾는다", size=36, color=T["ink"], weight="extrabold", spacing=-1.5)
    d.rect(s, 0.7, 2.1, 5.2, 4.75, T["ink"], radius=0.3)
    d.text(s, 1.1, 2.45, 4.5, 0.4, "데이터", size=14, color="#BFBAB0", weight="semibold")
    d.text(s, 1.1, 2.85, 4.5, 1.2, ["Give Me Some", "Credit"], size=30, color="#FFFFFF", weight="extrabold", line=1.05)
    d.text(s, 1.1, 5.4, 4.5, 1.2, ["Kaggle 2011", "150,000명 × 9열 · 결측치 처리 완료"], size=14, color="#BFBAB0",
           weight="medium", line=1.4)
    d.rect(s, 6.1, 2.1, 6.53, 2.25, T["lemon"], radius=0.3)
    d.text(s, 6.5, 2.4, 5.8, 0.4, "목표", size=14, color=T["sub"], weight="semibold")
    d.text(s, 6.5, 2.8, 5.8, 1.4, ["2년 내 90일 이상 연체 여부를", "예측하는 이진 분류"], size=24, color=T["ink"],
           weight="bold", line=1.2)
    d.rect(s, 6.1, 4.55, 3.17, 2.3, T["sky"], radius=0.3)
    d.text(s, 6.45, 4.85, 2.6, 0.4, "조건", size=14, color=T["sub"], weight="semibold")
    d.text(s, 6.45, 5.25, 2.6, 1.4, ["Lab 1–6 방법", "MLP 필수"], size=20, color=T["ink"], weight="bold", line=1.25)
    d.rect(s, 9.46, 4.55, 3.17, 2.3, T["mint"], radius=0.3)
    d.text(s, 9.8, 4.85, 2.6, 0.4, "진행", size=14, color=T["sub"], weight="semibold")
    d.text(s, 9.8, 5.25, 2.6, 1.4, ["발견 → 질문 →", "다음 Lab"], size=20, color=T["ink"], weight="bold", line=1.25)
    d.notes(s, "problem")

    # Lab 1
    s = d.slide(T["bg"])
    d.text(s, 0.7, 0.55, 8, 0.4, [[("Lab 1", {"color": T["ink"], "weight": "bold"}),
                                   ("  /  데이터 로딩 · 기술통계", {"color": T["mute"], "weight": "medium"})]], size=15)
    for i in range(6):
        d.rect(s, 10.0 + i * 0.45, 0.65, 0.3, 0.3, T["ink"] if i == 0 else "#DDD8CE", radius=0.15)
    d.text(s, 0.7, 0.95, 12, 1.0, "연체 고객은 6.7%뿐, 데이터 곳곳엔 극단값이 숨어 있다", size=31, color=T["ink"],
           weight="extrabold", spacing=-1.2)
    # 큰 타일: 불균형
    d.rect(s, 0.7, 2.05, 5.0, 4.05, T["peach"], radius=0.3)
    d.text(s, 1.1, 2.35, 4.3, 0.4, "발견 1 · 극단적 불균형", size=14, color=T["sub"], weight="semibold")
    d.text(s, 1.1, 2.8, 4.3, 1.5, [[("6.7", {}), ("%", {"size": 44})]], size=92, color=T["ink"], weight="extrabold", spacing=-3)
    d.rect(s, 1.1, 4.55, 4.2, 0.3, "#FFFFFF", radius=0.15)
    d.rect(s, 1.1, 4.55, 4.2 * 0.067, 0.3, T["red"], radius=0.15)
    d.text(s, 1.1, 5.0, 4.3, 0.9, ["연체 10,026 · 정상 139,974", "모두 '정상'이라 찍어도 정확도 93.3%"], size=13.5,
           color=T["sub"], weight="medium", line=1.35)
    # 작은 타일들
    d.rect(s, 5.9, 2.05, 3.27, 1.95, T["lav"], radius=0.3)
    d.text(s, 6.25, 2.3, 2.8, 0.4, "발견 2 · 사용 비율 최댓값", size=13, color=T["sub"], weight="semibold")
    d.text(s, 6.25, 2.75, 2.8, 0.8, "50,708", size=36, color=T["ink"], weight="extrabold", spacing=-1)
    d.text(s, 6.25, 3.5, 2.8, 0.4, "비율인데 1을 훌쩍 넘는다", size=12.5, color=T["sub"], weight="medium")
    d.rect(s, 9.37, 2.05, 3.26, 1.95, T["mint"], radius=0.3)
    d.text(s, 9.72, 2.3, 2.8, 0.4, "발견 2 · 부채 비율 > 1", size=13, color=T["sub"], weight="semibold")
    d.text(s, 9.72, 2.75, 2.8, 0.8, "35,137행", size=36, color=T["ink"], weight="extrabold", spacing=-1)
    d.text(s, 9.72, 3.5, 2.8, 0.4, "소득 6638·6746 채운 값과 겹침", size=12.5, color=T["sub"], weight="medium")
    d.rect(s, 5.9, 4.15, 6.73, 1.95, T["sky"], radius=0.3)
    d.text(s, 6.25, 4.4, 6.0, 0.4, "발견 3 · 극단값이 관계를 가린다 (사용 비율 ↔ 연체)", size=13, color=T["sub"], weight="semibold")
    d.text(s, 6.25, 4.85, 6.2, 0.8, [[("Pearson ", {"size": 18, "weight": "semibold", "color": T["sub"]}),
                                      ("-0.002", {}), ("   →   ", {"size": 22, "color": T["mute"]}),
                                      ("순위 ", {"size": 18, "weight": "semibold", "color": T["sub"]}), ("0.24", {})]],
           size=34, color=T["ink"], weight="extrabold")
    # 다음 질문
    d.rect(s, 0.7, 6.3, 11.93, 0.7, T["ink"], radius=0.35)
    d.text(s, 1.1, 6.3, 8.5, 0.7, [[("다음 질문  ", {"color": "#BFBAB0", "weight": "semibold"}),
                                    ("이 값들을 그대로 써도 될까?  →  Lab 2", {"color": "#FFFFFF", "weight": "bold"})]],
           size=16, anchor="m")
    d.text(s, 9.3, 6.3, 3.0, 0.7, [[("기준선 ", {"color": "#BFBAB0", "weight": "medium"}),
                                    ("AUC 0.50", {"color": "#FFFFFF", "weight": "bold"})]], size=15, anchor="m", align="r")
    d.notes(s, "lab1")
    return d.save("deck_C_bento.pptx")


if __name__ == "__main__":
    for f in (build_A, build_B, build_C):
        print("saved", f())
