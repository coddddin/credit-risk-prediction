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


# ---------------- 대본: 7분(420초)에 맞춘 짧은 대본 ----------------
# 장 번호: (장 이름, 목표 초, 짧은 대본). 슬라이드 노트 맨 위와 slides/script.md 에 들어간다.
# 노트 아래쪽에는 질문 대비용 '보충 설명'(notes() 에 넘긴 긴 대본)이 붙는다.
SCRIPTS = {
    1: ("표지", 10,
        "Lab 1부터 6까지의 실습 방법을 그대로 따라가며, 신용위험 예측 문제를 풀어 본 과정을 발표하겠습니다."),
    2: ("문제 정의", 25,
        "데이터는 15만 명의 재무 정보가 담긴 Give Me Some Credit이고, 목표는 2년 안에 연체할지를 예측하는 이진 분류입니다. "
        "진행은 Lecture 6의 AI 모델링 6단계를 따르고, 이 단계를 Lab 1부터 6까지 순서대로 밟아 가겠습니다."),
    3: ("Lab 1 · 진행", 15,
        "Lab 1의 일곱 단계를 그대로 실제 데이터에 적용했습니다. 15만 행 9열에 모두 숫자였고, "
        "Lab과 달리 헤더 줄이 있어 header=0만 더했습니다."),
    4: ("Lab 1 · 결과", 20,
        "연체 고객은 6.7%뿐이라, 모두 정상이라고 찍어도 정확도가 93%입니다. "
        "또 1을 넘는 비율, 빈 소득을 채운 값, 불가능한 연체 횟수 같은 이상값이 곳곳에 있었습니다."),
    5: ("Lab 1 · 발견·결론", 23,
        "사용 비율은 일반 상관계수로는 연체와 무관해 보이지만, 순위로 보면 세 번째로 중요합니다. 극단값이 관계를 가린 겁니다. "
        "그래서 평가는 정확도가 아닌 AUC로 하고, 모델에 넣기 전에 극단값부터 처리해야 합니다. 그럼 이 값들, 어떻게 바꿔야 할까요?"),
    6: ("Lab 2 · 진행", 15,
        "Lab 2 순서대로 히스토그램부터 산점도 행렬까지 그리고, 네 가지 변환을 적용했습니다. "
        "15만 개 점은 겹쳐 보이지 않아 산점도는 5천 건만 뽑았습니다."),
    7: ("Lab 2 · 결과", 20,
        "그려 보니 사용 비율, 부채 비율, 소득은 막대 하나에 몰려 있었습니다. "
        "로그 변환을 하자 부채 비율의 왜도가 95에서 1.7로 줄며 분포가 펴졌습니다. 연체 횟수는 96, 98 코드 때문에 여전히 치우쳐 있었습니다."),
    8: ("Lab 2 · 발견·결론", 23,
        "변환의 효과는 모델에 따라 정반대였습니다. 로그와 표준화는 로지스틱 회귀만 0.82로 올렸고, "
        "0과 1로 바꾸는 Binarizer는 결정트리만 0.78로 올렸습니다. 그래서 회귀와 MLP에는 로그와 표준화를 씁니다. 그럼 변수 8개가 다 필요할까요?"),
    9: ("Lab 3 · 진행", 15,
        "Lab 3의 특성 선택 네 가지와 검증 방법 네 가지를 그대로 적용했습니다. "
        "LOOCV는 15만 번 학습에 약 20시간이 걸려, 시간만 재고 천 건 표본으로 확인했습니다."),
    10: ("Lab 3 · 결과", 20,
         "방법마다 고른 변수가 달랐습니다. 통계 검정은 값이 큰 소득을, 트리는 범위가 넓은 변수를 우대했습니다. "
         "세 방법이 모두 고른 건 사용 비율과 나이였고, 변수를 줄여도 AUC가 거의 오르지 않아 8개를 모두 쓰기로 했습니다."),
    11: ("Lab 3 · 발견·결론", 23,
         "시험을 한 번만 나누면 나누는 방식에 따라 AUC가 0.816에서 0.828까지 흔들렸습니다. 층화 10-fold는 0.824 ± 0.009로 안정적이었습니다. "
         "그래서 이후 비교는 연체 비율을 맞춘 층화 10-fold AUC로 통일합니다. 그럼 어떤 알고리즘이 가장 잘 골라낼까요?"),
    12: ("Lab 4 · 진행", 15,
         "Lab 4 순서대로 평가지표 여섯 가지를 재고, 알고리즘 여섯 개를 같은 층화 10-fold로 비교했습니다. "
         "SVM은 15만 건에 한 시간이 넘어 1만 건 표본으로 비교했습니다."),
    13: ("Lab 4 · 결과", 20,
         "정확도는 93%였지만, 혼동행렬을 보니 연체자 3,309명 중 43명만 잡았습니다. 이 문제는 연체자를 놓치는 2종 오류가 더 비쌉니다. "
         "연체자에 가중치를 주자 75%를 잡았고, AUC도 0.844로 올랐습니다."),
    14: ("Lab 4 · 발견·결론", 23,
         "알고리즘 여섯 개 중 LDA가 0.832로 가장 높았고, 로지스틱 회귀와 나이브 베이즈가 뒤를 이었습니다. 트리는 과적합으로 0.6에 그쳤습니다. "
         "그래서 가중치를 기본으로 하고 선형 모델과 앙상블을 후보로 둡니다. 그런데 지금까지의 평가, 공정했을까요?"),
    15: ("Lab 5 · 진행", 15,
         "Lab 5 순서대로 파이프라인으로 누수를 막고, 앙상블 네 가지를 비교한 뒤, 튜닝한 모델을 저장했습니다. "
         "Lab 1에서 찾은 96, 98 코드와 소득 대체값 정리도 파이프라인에 넣었습니다."),
    16: ("Lab 5 · 결과", 20,
         "누수를 막아도 AUC는 소수 넷째 자리까지 같아, 지금까지의 결과는 유효했습니다. 데이터 정리 효과도 거의 없었습니다. "
         "결정트리 하나는 0.6이었지만, 여러 개를 합치자 0.82에서 0.85까지 올랐습니다."),
    17: ("Lab 5 · 발견·결론", 23,
         "랜덤 포레스트가 학습 데이터를 외우지 못하게 잎의 최소 인원을 튜닝하자 0.856으로 지금까지 최고가 됐습니다. "
         "이 모델을 저장해 다시 불러와도 같은 성능이었습니다. 그럼 신경망은 이 기준을 넘을 수 있을까요?"),
}

# ---------------- 예상 질문 (교수님은 PPT만 보고 질문) ----------------
QA = {
    2: [("왜 정확도가 아니라 AUC를 주 지표로 보나?",
         "연체가 6.7%뿐이라 전부 정상이라고 해도 정확도가 93%다. AUC는 연체자를 정상 고객보다 더 위험하게 줄 세우는 능력이라 "
         "불균형 데이터에서도 의미가 있고, Kaggle 대회의 공식 지표도 AUC다."),
        ("'연체 이력(통합)', '대출 보유(통합)'은 무엇인가?",
         "원본 Kaggle의 60–89일·90일 이상 연체 횟수, 대출·부동산 대출 수를 NOTE 논문에서 합쳐 0/1로 만든 변수다.")],
    3: [("header=0은 왜 넣었나?",
         "Lab 자료는 헤더 없는 CSV 기준이다. 이 파일은 첫 줄이 열 이름이라 그대로 읽으면 769행이 되고 모든 열이 문자열이 된다."),
        ("중복 행은 처리했나?",
         "완전히 같은 행 1,222개를 확인만 하고 지우지 않았다. 서로 다른 고객이 같은 값을 가질 수 있어서다.")],
    4: [("6638, 6746이 결측을 채운 값이라는 근거는?",
         "두 값이 합계 29,737행으로 원본 Kaggle 데이터의 소득 결측 수(29,731)와 거의 같고, 그중 27,904행은 부채 비율도 1을 넘는다."),
        ("96, 98은 무엇인가?",
         "연체 횟수로는 나올 수 없는 값(269행)이라 원본 데이터의 특수 코드로 보인다. 이후 단계에서 따로 처리한다.")],
    5: [("Pearson과 Spearman의 차이는?",
         "Pearson은 값 그대로의 직선 관계라 극단값에 크게 흔들리고, Spearman은 값을 순위로 바꿔 계산해 극단값 영향이 적다."),
        ("상관이 높으면 중요한 변수인가?",
         "함께 움직인다는 뜻일 뿐 원인은 아니다(Lecture 5 상관 vs 인과). 모델에서의 중요도는 Lab 3 특성 선택으로 확인한다.")],
    6: [("로그 변환은 Lab에 없는데 왜 했나?",
         "Lab 1 (e) '치우친 분포를 어떻게 바로잡나'의 답이고, Lab 2 본문이 여러 변환(view)을 만들어 시험해 보라고 권한다. "
         "log1p는 0도 변환할 수 있다."),
        ("왜 로지스틱 회귀로도 비교했나?",
         "Lab 2 p.8의 '변환은 회귀·신경망처럼 가중치를 쓰는 모델에 유용하다'를 확인하고, Lab 6 MLP에 넣을 데이터를 정하기 위해서다."),
        ("산점도를 5,000건만 그려도 되나?",
         "무작위 추출(random_state=7)이고 경향을 눈으로 보는 용도라 충분하다. 모델 학습과 평가는 15만 건 전체로 했다.")],
    7: [("로그 후 소득 왜도가 음수(-4.9)가 된 이유는?",
         "소득 0원인 1,634행이 로그 후에도 0으로 남아 왼쪽에 꼬리가 생겼기 때문이다."),
        ("부채 비율 그림의 두 번째 봉우리는?",
         "소득이 비어 있던 고객의 부채 '금액'이다. Lab 1에서 찾은 27,904행과 같은 집단이다.")],
    8: [("결정트리는 왜 변환에 영향을 받지 않나?",
         "트리는 '값이 기준보다 큰가'로만 나눈다. 단위를 바꾸거나 로그를 씌워도 값의 순서가 같아 결과가 같다."),
        ("결정트리 정확도가 기준선보다 낮은 이유는?",
         "가지 수를 제한하지 않아 학습 데이터를 외운 것으로 보인다(과적합). 깊이 제한과 앙상블은 Lab 4–5에서 다룬다."),
        ("Binarizer는 왜 트리에만 좋았나?",
         "값이 0/1로 단순해져 트리가 외울 세부가 줄었기 때문으로 보인다. 회귀는 크기 정보를 잃어 약간 낮아졌다(0.787 → 0.778)."),
        ("변환을 교차검증 전에 전체 데이터에 맞춘 건 데이터 누수 아닌가?",
         "Lab 2 방식 그대로 했다. 평균·표준편차만 쓰는 변환이라 영향은 작지만, Lab 5에서 Pipeline으로 학습 fold 안에서만 맞추도록 바로잡는다."),
        ("Normalizer는 왜 낮나?",
         "행마다 길이를 1로 맞추면 값이 큰 소득이 길이를 거의 다 차지해 다른 정보가 사라진다. 희소 데이터용 방법이다.")],
    9: [("LOOCV를 15만 건 전체로 하지 않은 이유는?",
         "측정해 보니 1회 학습 0.49초 × 15만 번 ≈ 20시간이고, 시험 데이터가 1명이라 AUC를 계산할 수 없다. 1,000건 표본으로 정확도만 확인했다(92.6%, 표준편차 26%)."),
        ("RFE 결과가 원본과 로그+표준화에서 다른 이유는?",
         "회귀 계수는 열의 단위에 따라 크기가 달라진다. 단위를 맞춰야 변수끼리 공정하게 비교된다.")],
    10: [("카이제곱이 원본에서 소득을 1위로 뽑은 이유는?",
          "카이제곱 점수는 값의 크기에 비례한다. 소득은 최대 300만이라 유리했고, 로그 변환 후에는 오히려 최하위다."),
         ("트리 중요도는 어떻게 계산하나?",
          "트리가 나눌 때 불순도(지니)가 줄어든 양을 변수별로 더해 100개 트리에서 평균한다. 값의 범위가 넓은 변수가 높게 나오는 경향이 있다."),
         ("변수를 줄이지 않은 이유는?",
          "상위 6개(0.826)와 8개(0.824)의 차이 0.002가 층화 10-fold 표준편차(약 0.009)보다 작아, 줄여서 얻는 이득이 우연과 구별되지 않는다."),
         ("PCA를 쓰지 않은 이유는?",
          "3개 성분으로 AUC가 0.822로 비슷했지만, 성분은 여러 변수를 섞은 값이라 '왜 연체 위험으로 봤는지' 설명하기 어렵다.")],
    11: [("층화(Stratified)가 왜 필요한가?",
          "연체가 6.68%뿐이라 일반 KFold는 fold마다 연체 비율이 6.45~7.18%로 달라진다. 층화하면 모든 fold가 6.68%로 같아진다."),
         ("표준편차가 왜 중요한가?",
          "데이터가 바뀌어도 비슷하게 맞히는지(안정성)를 보여 준다. 두 모델의 차이가 이 값보다 작으면 우연일 수 있다."),
         ("ShuffleSplit의 흔들림이 가장 작은데 왜 기준으로 삼지 않았나?",
          "시험 데이터가 33%로 커서 덜 흔들리지만, 회차끼리 같은 데이터가 겹친다. k-fold는 모든 데이터가 정확히 한 번씩 시험에 쓰인다."),
         ("정확도는 왜 비교하지 않았나?",
          "모든 방법에서 약 93.2%로 '전부 정상' 기준선과 같아 차이가 보이지 않는다.")],
    12: [("logloss가 음수로 나온 이유는?",
          "sklearn은 모든 점수를 '클수록 좋음'으로 맞추려고, 작을수록 좋은 지표에 음수를 붙인다. 실제 logloss는 0.199이고, 늘 6.68%라고 말하는 모델은 0.245다."),
         ("SVM을 15만 건 전체로 돌리지 않은 이유는?",
          "학습 시간이 데이터 2배에 약 4배(n²)로 늘었다. 2만 건 8초 → 13.5만 건 약 6분, 10-fold면 약 1시간이다. 대신 같은 1만 건 층화 표본에서 6개를 모두 돌려 공정하게 비교했다."),
         ("Lab 4는 정확도로 비교하는데 왜 AUC로 했나?",
          "정확도는 6개 모두 88~93%로 '전부 정상' 기준선(93.3%) 근처라 차이를 보여 주지 못한다.")],
    13: [("1종·2종 오류 중 무엇이 더 중요한가?",
          "2종(연체자를 정상으로 통과)이다. 원금을 잃는다. 1종(정상 고객 거절)은 이자 수익 손실이라 상대적으로 작다."),
         ("가중치를 주면 왜 AUC까지 오르나?",
          "balanced는 연체 1명을 정상 약 14명만큼 중요하게 학습한다. 선형 모델은 다수(정상) 쪽에 맞춰 기울기가 정해지는데, "
          "가중치가 연체자 쪽 경계를 더 신경 쓰게 해 순위도 좋아진 것으로 보인다. 층화 10-fold 0.824 → 0.844(±0.007)로 흔들림보다 크게 올랐다."),
         ("정상 고객 9,766명을 거절하는데 괜찮나?",
          "연체로 판정한 사람 중 20%가 진짜 연체자로 무작위(6.7%)의 3배다. 실제 기준값은 '놓친 연체자 1명 비용 vs 거절한 정상 고객 1명 비용'으로 정해야 하고, "
          "기준값은 나중에 바꿀 수 있으니 모델은 AUC로 먼저 고른다."),
         ("시험 데이터는 왜 33%로 나눴나?",
          "Lab 4의 혼동행렬·리포트 코드가 test_size=0.33, seed 7이다. 연체 비율이 같도록 층화해서 나눴다.")],
    14: [("LDA가 LR보다 높은 이유는?",
          "둘 다 직선 경계지만 LDA는 정규분포를 가정하고 평균·공분산으로 한 번에 계산한다. 로그+표준화로 분포가 정규에 가까워져 유리했던 것으로 보인다. "
          "다만 차이 0.008은 흔들림(약 0.008~0.009) 수준이라 확실한 우위는 아니다."),
         ("KNN이 낮은 이유는?",
          "기본 이웃 수 k=5라 확률이 0, 0.2, …, 1 여섯 단계뿐이어서 순위를 잘게 매기지 못한다. k를 키우면 오를 수 있어 Lab 5 튜닝 대상이다."),
         ("결정트리(CART)는 왜 0.595인가?",
          "깊이 제한이 없어 학습 데이터를 외웠다(과적합). 여러 트리를 합치는 Lab 5의 앙상블로 보완한다."),
         ("SVM은 왜 가중치 없이 0.579였나?",
          "정상이 93%라 오류를 줄이려면 거의 다 정상으로 가르는 경계가 유리하다. 그래서 점수가 연체 순위를 거의 담지 못했고, 가중치를 주면 0.810으로 회복됐다.")],
    15: [("데이터 누수란? 왜 문제인가?",
          "시험 데이터의 정보(평균·표준편차 등)가 학습 단계에 섞이는 것이다. 실제로는 새 고객 데이터를 미리 볼 수 없으므로 점수가 부풀려질 수 있다. "
          "Pipeline은 전처리를 각 fold의 학습 데이터로만 맞춘다."),
         ("Voting에서 SVM을 뺀 이유는?",
          "Lab 4에서 SVM은 15만 건 10-fold에 약 1시간이 걸렸고, 확률을 바로 내지 못해 soft voting에도 맞지 않는다. 같은 자리에 빠른 확률 모델인 나이브 베이즈를 넣었다."),
         ("데이터 정리는 구체적으로 무엇을 했나?",
          "30–59일 연체의 96·98 코드는 0으로 바꾸고 '코드였음' 표시 열을 추가했다. 소득 6638·6746(결측 대체값으로 추정)에도 표시 열을 추가했다. 행을 지우지는 않았다.")],
    16: [("누수를 막았는데 왜 AUC가 같나?",
          "로그 변환은 값마다 따로 계산해 누수가 없고, 표준화는 평균·표준편차를 쓰는데 15만 건에서는 fold마다 거의 같다. 데이터가 작으면 차이가 날 수 있다."),
         ("96/98 코드는 연체율이 54.6%인데 왜 정리 효과가 없나?",
          "같은 고객이 '연체 이력' 변수에서 이미 1로 잡혀 있어(60–89일·90일 이상 연체도 같은 코드) 모델이 이미 위험 고객으로 보고 있었다."),
         ("Bagging과 랜덤 포레스트의 차이는?",
          "둘 다 다른 표본으로 트리를 여러 개 만든다. RF는 나눌 때마다 변수 3개만 무작위로 후보로 둬서 트리끼리 덜 닮게 만든다."),
         ("AdaBoost는 왜 기본 설정에서 가장 높았나?",
          "기본 약한 학습기가 깊이 1인 얕은 트리라 외우지 못하고, 앞 트리가 틀린 고객에 가중치를 높여 다음 트리가 고치게 한다.")],
    17: [("min_samples_leaf를 키우면 왜 좋아지나?",
          "잎에 최소 인원을 두면 몇 명만을 위한 세밀한 규칙(외우기)을 만들 수 없어 새 데이터에 강해진다. 확률도 여러 명의 비율이라 더 부드러워져 순위가 좋아진다."),
         ("Grid와 Random Search 중 무엇이 나은가?",
          "Grid는 정한 값만 보고, Random은 같은 횟수로 넓은 범위를 본다. 여기서는 50 이상에서 평평해 둘 다 비슷했다. 변수가 많을수록 Random이 효율적이다."),
         ("저장한 모델을 다른 컴퓨터에서 쓰려면?",
          "같은 Python·scikit-learn 버전이 필요하고, 파이프라인 안의 정리·로그 함수도 같은 코드로 있어야 한다(Lab 5 (s)). 버전은 노트북에 기록했다."),
         ("0.856이면 좋은 성능인가?",
          "Lab 4 등급표로 0.80–0.90 = good(B)이다. 연체자를 정상 고객보다 위험하게 줄 세울 확률이 약 86%라는 뜻이다.")],
}
# 전체 장 시간 배분 (합계 420초). 아직 만들지 않은 장은 대본 파일에 '작성 예정'으로 표시된다.
FULL_PLAN = [("표지", 10), ("문제 정의", 25)] \
    + [(f"Lab {n} · {part}", sec) for n in range(1, 7) for part, sec in (("진행", 15), ("결과", 20), ("발견·결론", 23))] \
    + [("최종 결론", 35)]
RECORDS = []


def speak_seconds(script):
    """한국어 발표 속도 약 초당 5글자(공백 제외)로 예상 시간 계산."""
    return round(len(script.replace(" ", "")) / 5)


def notes(s, extra):
    """슬라이드 노트 = 짧은 대본(SCRIPTS) + 보충 설명(extra)."""
    no = len(prs.slides)
    extra = " ".join(extra.split())
    if no in SCRIPTS:
        name, sec, short = SCRIPTS[no]
        RECORDS.append((no, name, sec, short, extra))
        qa = "".join(f"\nQ. {q}\nA. {a}\n" for q, a in QA.get(no, []))
        s.notes_slide.notes_text_frame.text = (
            f"[대본 · 목표 {sec}초 / 예상 {speak_seconds(short)}초]\n{short}\n\n[보충 설명]\n{extra}"
            + (f"\n\n[예상 질문]{qa}" if qa else ""))
    else:
        s.notes_slide.notes_text_frame.text = extra


def write_script_md(path):
    PLANNED = FULL_PLAN[len(RECORDS):]
    done = sum(r[2] for r in RECORDS)
    plan = sum(p[1] for p in PLANNED)
    lines = ["# 발표 대본 (7분)", "",
             "build_deck.py 를 실행하면 자동으로 다시 만들어진다. 고칠 때는 build_deck.py 의 SCRIPTS 를 수정한다.",
             "예상 시간은 공백 제외 초당 약 5글자 기준이다. 실제로 소리 내어 읽으며 타이머로 확인할 것.", "",
             "## 시간 배분", "",
             "| 장 | 내용 | 목표 | 예상 |", "|---|---|---|---|"]
    for no, name, sec, short, _ in RECORDS:
        lines.append(f"| {no} | {name} | {sec}초 | {speak_seconds(short)}초 |")
    for i, (name, sec) in enumerate(PLANNED, start=len(RECORDS) + 1):
        lines.append(f"| {i} | {name} (작성 예정) | {sec}초 | - |")
    lines += ["", f"**합계 목표 {done + plan}초 ({(done + plan) // 60}분 {(done + plan) % 60}초)**"
                  f" · 작성된 장 {len(RECORDS)}개 = 목표 {done}초 / 예상 {sum(speak_seconds(r[3]) for r in RECORDS)}초", ""]
    lines += ["## 대본", ""]
    for no, name, sec, short, extra in RECORDS:
        lines += [f"### {no}. {name}  ·  {sec}초", "", short, "",
                  "<details><summary>보충 설명</summary>", "", extra, "", "</details>", ""]
    lines += ["## 예상 질문", "", "교수님은 PPT만 보고 질문하신다. 장표에 적힌 숫자와 용어 위주로 정리했다.", ""]
    for no, name, *_ in RECORDS:
        if QA.get(no):
            lines += [f"### {no}. {name}", ""]
            for q, a in QA[no]:
                lines += [f"**Q. {q}**", "", f"A. {a}", ""]
    Path(path).write_text("\n".join(lines), encoding="utf-8")


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


def capture_card(s, x, y, w, img, label, anim=None, img_w=None):
    """Lab PDF 원본 캡처 카드. img_w: 이미지 너비(인치, 기본은 카드의 90%)."""
    from PIL import Image
    iw, ih = Image.open(img).size
    img_w = img_w or w * 0.9
    ph = img_w * ih / iw
    h = ph + 0.64
    parts = [rect(s, x, y, w, h, C["card"], radius=0.18),
             text(s, x + 0.25, y + 0.14, w - 0.5, 0.3,
                  [[("Lab 자료  ", {"color": C["blue"], "weight": "bold"}), (label, {"color": C["mute"]})]],
                  size=12, weight="medium")]
    pic = s.shapes.add_picture(str(img), Inches(x + 0.25), Inches(y + 0.5), width=Inches(img_w))
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


def term_line(s, terms, y=1.62):
    """제목 아래 용어 풀이 한 줄. terms: [(용어, 쉬운 풀이)]"""
    pieces = [("용어  ", {"color": C["violet"], "weight": "bold"})]
    for i, (t, d) in enumerate(terms):
        if i:
            pieces.append(("   ·   ", {"color": C["line"]}))
        pieces += [(t, {"color": C["ink"], "weight": "semibold"}), (f" {d}", {"color": C["mute"]})]
    box = text(s, 0.8, y, 11.8, 0.3, [pieces], size=12, weight="medium", name="!!terms")
    # 반투명 처리: 본문보다 한 단계 뒤로 물러나 보이게 (불투명도 55%)
    for run in box.text_frame.paragraphs[0].runs:
        clr = run.font.color._color._xClr  # a:srgbClr
        alpha = clr.makeelement(qn("a:alpha"), {"val": "55000"})
        clr.append(alpha)


def lab_header(s, lab, part, chapter, title, terms=None):
    """part: 1=코드, 2=결과, 3=발견·결론. terms: 제목 아래 용어 풀이."""
    if terms:
        term_line(s, terms)
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
lab_header(s, 1, 1, "데이터 로딩 · 기술통계", "Lab 1의 순서 그대로, 15만 건을 불러와 살펴봤다",
           terms=[("기술통계", "평균·최댓값 등 요약"), ("상관", "함께 움직이는 정도"), ("왜도", "한쪽 쏠림 정도(0이면 대칭)")])

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
    cmd("data.shape", "3 행·열 개수"), [("(150000, 9)", "hi")],
    cmd("data.dtypes.value_counts()", "4 모두 숫자인지"), [("int64 7  ·  float64 2", "hi")],
    cmd("data.describe().loc['max']", "5 열별 최댓값"), [("revol 50708  ·  debt 329664  ·  income 3008750", "hi")],
    cmd("data.groupby('target').size()", "6 정상/연체 수"), [("0  139974  ·  1  ", "hi"), ("10026", "red")],
    cmd("data.corr()['target'][1:].nlargest(2)", "7 연체와 관계"), [("defaulted 0.314  ·  late30 0.126", "hi")],
    cmd("data.skew().nlargest(3)", "7 쏠림 큰 열"), [("income 127.4  ·  revol 97.6  ·  debt 95.2", "hi")],
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
lab_header(s, 1, 2, "데이터 로딩 · 기술통계", "연체 고객은 6.7%뿐, 데이터 곳곳엔 극단값이 숨어 있다",
           terms=[("기준선", "전부 '정상'이라고 찍었을 때의 성적"), ("극단값", "다른 값보다 지나치게 크거나 나올 수 없는 값")])
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
lab_header(s, 1, 3, "데이터 로딩 · 기술통계", "극단값이 '사용 비율'과 연체의 관계를 가리고 있었다",
           terms=[("Pearson", "값 그대로 본 관계"), ("Spearman", "순위로 본 관계"),
                  ("AUC", "연체자 구별 점수(0.5 찍기 ~ 1 완벽)")])
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

# =====================================================================
# Lab 2 - (1) 진행 : Lab 2 자료의 순서 + 실제 데이터 실습
# =====================================================================
s = new_slide()
lab_header(s, 2, 1, "시각화 · 데이터 준비", "Lab 2의 순서대로 그려 보고, 네 가지 방법으로 바꿔 봤다",
           terms=[("MinMax·표준화", "열마다 단위 맞추기"), ("Normalizer", "행마다 길이를 1로"),
                  ("Binarizer", "0/1로 바꾸기"), ("로그", "자릿수로 보기")])
steps = [("히스토그램", "hist()", "3개 변수 쏠림"), ("밀도", "kind='density'", "꼬리 긴 분포"),
         ("박스플롯", "kind='box'", "이상치 다수"), ("상관행렬", "matshow", "최대 0.22"),
         ("산점도 행렬", "scatter_matrix", "5,000건 샘플"), ("변환 4종", "MinMaxScaler 등", "+ 로그 변환"),
         ("원본 vs 변환", "DecisionTree", "정확도 89.4%")]
parts = [text(s, 0.8, 1.98, 11.7, 0.3, [[("Lab 2 자료의 진행 순서", {"weight": "bold"}),
                                         ("   → 실제 데이터에 같은 순서로 실습한 결과", {"color": C["mute"], "weight": "medium"})]],
              size=13)]
for i, (ko, fn, res) in enumerate(steps):
    x = 0.8 + i * (sw + gap)
    parts += [rect(s, x, 2.35, sw, 1.18, C["blue_pale"], radius=0.16),
              oval(s, x + 0.15, 2.47, 0.3, C["blue"]),
              text(s, x + 0.15, 2.47, 0.3, 0.3, str(i + 1), size=11, color="#FFFFFF", weight="bold", align="c", anchor="m"),
              text(s, x + 0.52, 2.47, sw - 0.6, 0.3, ko, size=13 if len(ko) < 6 else 12, weight="bold", anchor="m"),
              text(s, x + 0.15, 2.84, sw - 0.25, 0.25, fn, size=10.5, color=C["blue"], weight="medium", latin=MONO),
              rect(s, x + 0.12, 3.14, sw - 0.24, 0.3, C["bg"], radius=0.08),
              text(s, x + 0.12, 3.14, sw - 0.24, 0.3, res, size=10.5, color=C["ink"], weight="semibold", align="c",
                   anchor="m")]
group(s, parts, "anim1_steps")
hc = capture_card(s, 0.8, 3.7, 5.2, CAP / "lab2_challenge.png", "p.11 · 7 도전 과제 (o)", anim=2, img_w=4.4)
group(s, [rect(s, 0.8, 3.7 + hc + 0.14, 5.2, 7.05 - (3.7 + hc + 0.14), C["violet_pale"], radius=0.18),
          text(s, 1.05, 3.7 + hc + 0.14, 4.8, 7.05 - (3.7 + hc + 0.14),
               [[("Lab 2 p.8  ", {"color": C["violet"], "weight": "bold"}),
                 ("변환은 회귀·신경망처럼 가중치를 쓰는 모델에 유용하다", {"color": C["ink"], "weight": "semibold"})],
                [("→ 결정트리와 함께 로지스틱 회귀로도 확인", {"color": C["sub"], "weight": "medium"})]],
               size=13, anchor="m", line=1.35)], "anim3_quote")
cmd = lambda code, note: [(">>> ", "in"), (code.ljust(38), "out"), (f"# {note}", "dim")]
console(s, 6.2, 3.7, 6.33, 3.35, "실행 결과 · 02_lab2_preparing_data.ipynb", [
    cmd("X[0]", "6 원본 첫 행"), [("[0.766  45  2  0.803  9120  2  1  1]", "hi")],
    cmd("StandardScaler().fit_transform(X)[0]", "단위 맞춤"),
    [("[-0.02 -0.49 0.38 -0.17 0.19 1.14 1.99 0.59]", "hi")],
    cmd("Normalizer().fit_transform(X)[0]", "행 길이 1"),
    [("[0.  0.005  0.  0.  1.  0.  0.  0. ]", "red"), ("  ← 소득이 독차지", "dim")],
    cmd("log1p(X).skew()  # revol, debt", "쏠림 변화"), [("97.6 → 11.7  ·  95.2 → 1.7", "hi")],
    cmd("cross_val_score(DecisionTree, ...)", "7 정확도 비교"),
    [("원본 0.8938  ·  Normalizer 0.8940", "hi"), ("  ← 기준선 0.933보다 낮음", "dim")],
], size=11, anim=4, lh=15.5)
notes(s, """Lab 2는 단변량 시각화인 히스토그램, 밀도, 박스플롯, 다변량 시각화인 상관행렬과 산점도 행렬을 그린 뒤,
MinMax, Standard, Normalizer, Binarizer 네 가지 변환을 적용하고, 원본과 변환 데이터에 결정트리를 돌려 비교하는 순서입니다.
같은 순서로 실제 데이터에 적용했습니다. 15만 개 점으로 산점도 행렬을 그리면 점이 겹쳐 보이지 않아 5천 건을 무작위로 뽑았습니다.
Lab 1에서 찾은 치우친 분포를 바로잡기 위해 로그 변환 버전을 하나 더 만들었고, Lab 2 본문에서 변환이 회귀와 신경망에 유용하다고 한 것을 확인하려고 로지스틱 회귀로도 비교했습니다.
Normalizer는 행마다 길이를 1로 맞추는데, 소득 값이 커서 한 행의 길이를 거의 소득이 차지합니다.
도전 과제의 결정트리 정확도는 원본 89.4%, Normalizer 89.4%로 거의 같았고, 둘 다 모두 정상이라고 찍는 기준선 93.3%보다 낮았습니다.""")

# =====================================================================
# Lab 2 - (2) 결과 : 실제 데이터 시각화
# =====================================================================
s = new_slide()
lab_header(s, 2, 2, "시각화 · 데이터 준비", "그려 보니 세 변수가 한쪽에 몰려 있었고, 로그 변환으로 펴졌다",
           terms=[("왜도", "한쪽 쏠림 정도(0이면 대칭)"), ("log1p", "log(1+x) · 0도 변환 가능")])
group(s, [rect(s, 0.8, 2.15, 5.75, 4.9, C["card"], radius=0.25),
          text(s, 1.1, 2.32, 5.2, 0.35, [[("실제 데이터  ", {"color": C["red"], "weight": "bold"}),
                                         ("data.hist(bins=50) · 15만 건", {"color": C["mute"]})]], size=12, weight="medium"),
          s.shapes.add_picture(str(FIG / "lab2_hist.png"), Inches(1.0), Inches(2.75), width=Inches(5.35))],
      "anim1_hist")
group(s, [rect(s, 6.8, 2.15, 5.73, 3.25, C["card"], radius=0.25),
          text(s, 7.1, 2.32, 5.2, 0.35, [[("로그 변환 전후  ", {"weight": "bold"}),
                                         ("회색 원본 → 파랑 log1p", {"color": C["mute"]})]], size=12, weight="medium"),
          s.shapes.add_picture(str(FIG / "lab2_log_before_after.png"), Inches(7.0), Inches(2.72), width=Inches(5.35))],
      "anim2_log")
skew_items = [("사용 비율", "97.6", "11.7", C["blue"]), ("부채 비율", "95.2", "1.7", C["blue"]),
              ("소득", "127.4", "-4.9", C["blue"]), ("30–59일 연체", "22.6", "4.3", C["red"])]
parts = [text(s, 6.8, 5.55, 5.7, 0.3, [[("왜도 변화  ", {"weight": "bold"}), ("0에 가까울수록 대칭", {"color": C["mute"]})]],
              size=12, weight="medium")]
for i, (nm, a, b, col) in enumerate(skew_items):
    x = 6.8 + i * 1.45
    parts += [rect(s, x, 5.92, 1.37, 1.13, C["red_pale"] if col == C["red"] else C["blue_pale"], radius=0.16),
              text(s, x + 0.12, 6.0, 1.15, 0.3, nm, size=11, color=C["sub"], weight="semibold"),
              text(s, x + 0.12, 6.32, 1.2, 0.35, [[(a, {"color": C["mute"], "size": 12}), (" → ", {"color": C["mute"], "size": 12}),
                                                  (b, {"color": col})]], size=17, weight="extrabold"),
              text(s, x + 0.12, 6.7, 1.2, 0.3, "96·98로 아직 큼" if col == C["red"] else "펴짐", size=10,
                   color=col, weight="semibold")]
group(s, parts, "anim3_skew")
notes(s, """왼쪽은 실제 데이터의 히스토그램입니다. 사용 비율, 부채 비율, 소득은 극단값 때문에 가로축이 수만에서 수백만까지 늘어나, 거의 모든 값이 막대 하나에 몰렸습니다.
박스플롯에서도 이 세 변수는 상자가 납작하고 이상치 점만 길게 늘어섰습니다.
오른쪽은 이 세 변수에 log1p, 즉 log(1+x)를 적용한 전후입니다. 부채 비율의 왜도는 95.2에서 1.7로 줄어 거의 대칭이 됐고, 사용 비율도 97.6에서 11.7로 줄었습니다.
로그를 씌운 부채 비율 그림 오른쪽의 두 번째 봉우리는, Lab 1에서 찾은 소득을 채운 고객들의 부채 금액이 그림으로 드러난 것입니다.
30–59일 연체 횟수도 22.6에서 4.3으로 줄었지만, 96, 98이라는 특수 코드 때문에 여전히 가장 치우친 변수로 남았습니다. 이 값은 따로 처리해야 합니다.""")

# =====================================================================
# Lab 2 - (3) 발견 · 결론
# =====================================================================
s = new_slide()
lab_header(s, 2, 3, "시각화 · 데이터 준비", "로그+표준화는 회귀만, Binarizer는 트리만 성능을 올렸다",
           terms=[("AUC", "연체자 구별 점수(0.5 찍기 ~ 1 완벽)"), ("10-fold", "10등분해 10번 시험"),
                  ("과적합", "학습 데이터를 외워 새 데이터에 약함")])
group(s, [rect(s, 0.8, 2.15, 6.0, 4.05, C["card"], radius=0.25),
          text(s, 1.1, 2.33, 5.4, 0.35, [[("변환별 AUC  ", {"weight": "bold", "color": C["ink"]}),
                                         ("10-fold · 결정트리 vs 로지스틱 회귀", {"color": C["mute"]})]], size=13),
          s.shapes.add_picture(str(FIG / "lab2_auc_by_view.png"), Inches(1.0), Inches(2.8), width=Inches(5.6))],
      "anim1_chart")
finds = [("01", "결정트리: Binarizer만 0.59 → 0.78", "나머지 변환은 0.59 그대로 · 정확도 89.4% < 기준선 93.3%", C["red"]),
         ("02", "로지스틱 회귀: 0.787 → 0.824", "표준화 0.802 · 로그+표준화 0.824 · Binarizer는 0.778", C["blue"]),
         ("03", "Normalizer: 0.522로 하락 (찍기 수준)", "행 길이를 1로 맞추면 소득이 값 대부분을 차지", C["blue"])]
parts = []
for i, (no, head, body, col) in enumerate(finds):
    yy = 2.15 + i * 0.72
    parts += [text(s, 7.1, yy, 0.6, 0.4, no, size=17, color=col, weight="extrabold"),
              text(s, 7.65, yy, 4.9, 0.4, head, size=16, weight="bold"),
              text(s, 7.65, yy + 0.34, 4.9, 0.35, body, size=12.5, color=C["sub"], weight="medium")]
group(s, parts, "anim2_finds")
conclusion_card(s, 7.05, 4.4, 5.48, 1.8, [
    [("① ", {"color": C["violet"], "weight": "bold"}), ("회귀·MLP에는 로그+표준화 데이터를 쓴다", {})],
    [("② ", {"color": C["violet"], "weight": "bold"}), ("트리 계열은 변환보다 과적합을 잡아야 한다", {})],
], anim=3)
next_bar(s, "변수 8개가 다 필요할까? 평가는 믿을 만할까?", "Lab 3", "최고 AUC", "0.824", anim=4)
notes(s, """변환별로 10-fold AUC를 비교했습니다. 회색이 결정트리, 파란색이 로지스틱 회귀입니다.
결정트리는 값의 순서만 보고 나누기 때문에 MinMax, 표준화, 로그 변환을 해도 AUC 0.59로 그대로였습니다. 오히려 Binarizer는 0과 1로 단순해져 과적합이 줄면서 0.78이 나왔습니다.
정확도 89.4%는 모두 정상이라고 찍는 93.3%보다 낮아, 가지를 끝까지 뻗은 트리가 학습 데이터를 외운 과적합으로 보입니다.
로지스틱 회귀는 원본 0.787에서 표준화 0.802, 로그와 표준화 0.824로 올랐습니다. Lab 2 본문대로 가중치를 쓰는 모델에는 변환이 효과가 있습니다.
Normalizer는 0.522로 거의 무작위 수준이 됐습니다. 희소 데이터용 방법이라 이 데이터에는 맞지 않습니다.
목표 관점의 결론은, 회귀와 이후 MLP에는 로그와 표준화를 거친 데이터를 쓰고, 트리 계열은 변환보다 과적합 조절이 과제라는 것입니다.
다음 질문은, 변수 8개가 정말 다 필요한지, 그리고 지금처럼 하는 평가가 믿을 만한지입니다.""")

# =====================================================================
# Lab 3 - (1) 진행 : 특성 선택 4가지 + 리샘플링 4가지
# =====================================================================
s = new_slide()
lab_header(s, 3, 1, "특성 선택 · 리샘플링", "Lab 3의 특성 선택 4가지와 검증 방법 4가지를 모두 적용했다",
           terms=[("특성 선택", "예측에 쓸 변수 고르기"), ("RFE", "하나씩 빼 보며 고르기"),
                  ("PCA", "여러 변수를 소수 성분으로 압축"), ("LOOCV", "한 명씩 빼고 시험")])
steps3 = [("통계 검정", "SelectKBest", "로그 후 late30 1위"), ("RFE", "RFE(LR, 3)", "revol·age 등 3개"),
          ("PCA", "PCA(3)", "24·16·16%"), ("트리 중요도", "ExtraTrees", "debt·revol 상위"),
          ("한 번 분할", "train_test_split", "AUC 0.823"), ("k-fold", "KFold(10)", "0.824 ± 0.007"),
          ("LOOCV", "LeaveOneOut", "약 20시간"), ("반복 분할", "ShuffleSplit", "0.826 ± 0.004")]
sw3, gap3 = 1.38, 0.098
parts = [text(s, 0.8, 1.98, 11.7, 0.3, [[("Lab 3 자료의 진행 순서", {"weight": "bold"}),
                                         ("   → 1–4 특성 선택 · 5–8 리샘플링, 실제 데이터 결과", {"color": C["mute"], "weight": "medium"})]],
              size=13)]
for i, (ko, fn, res) in enumerate(steps3):
    x = 0.8 + i * (sw3 + gap3)
    fill = C["blue_pale"] if i < 4 else C["violet_pale"]
    dot = C["blue"] if i < 4 else C["violet"]
    parts += [rect(s, x, 2.35, sw3, 1.18, fill, radius=0.16),
              oval(s, x + 0.13, 2.47, 0.28, dot),
              text(s, x + 0.13, 2.47, 0.28, 0.28, str(i + 1), size=10.5, color="#FFFFFF", weight="bold", align="c", anchor="m"),
              text(s, x + 0.47, 2.47, sw3 - 0.5, 0.28, ko, size=11.5, weight="bold", anchor="m"),
              text(s, x + 0.13, 2.84, sw3 - 0.2, 0.25, fn, size=9.5, color=dot, weight="medium", latin=MONO),
              rect(s, x + 0.1, 3.14, sw3 - 0.2, 0.3, C["bg"], radius=0.08),
              text(s, x + 0.1, 3.14, sw3 - 0.2, 0.3, res, size=10, color=C["ink"], weight="semibold", align="c", anchor="m")]
group(s, parts, "anim1_steps")
h3 = capture_card(s, 0.8, 3.7, 5.2, CAP / "lab3_rfe.png", "p.3 · 2 RFE", anim=2, img_w=3.3)
capture_card(s, 0.8, 3.7 + h3 + 0.12, 5.2, CAP / "lab3_kfold.png", "p.6 · 6 k-fold 교차검증", anim=2, img_w=4.6)
cmd = lambda code, note: [(">>> ", "in"), (code.ljust(38), "out"), (f"# {note}", "dim")]
console(s, 6.2, 3.7, 6.33, 3.35, "실행 결과 · 03_lab3_feature_selection_resampling.ipynb", [
    cmd("SelectKBest(chi2, k=4).fit(X, Y)", "1 원본"),
    [("income · debt · late30 · defaulted", "hi"), ("  ← 값 큰 열 유리", "dim")],
    cmd("SelectKBest(chi2, k=4).fit(logX, Y)", "1 로그 후"), [("late30 · defaulted · age · revol", "hi")],
    cmd("RFE(LogisticRegression(), 3)", "2 하나씩 빼기"), [("revol · age · defaulted", "hi")],
    cmd("PCA(3).explained_variance_ratio_", "3 압축"),
    [("[0.24 0.16 0.16]", "hi"), ("   ← 원본은 소득이 97.5% 독점", "dim")],
    cmd("ExtraTrees().feature_importances_", "4 트리 중요도"),
    [("debt .26 · revol .25 · income .20 · age .12", "hi")],
    cmd("LeaveOneOut()  # 15만 회 학습", "7 시간 측정"), [("1회 0.49초 × 150,000 ≈ 20시간", "red")],
], size=10.5, anim=3, lh=15)
notes(s, """Lab 3은 특성 선택 네 가지, 즉 단변량 통계 검정, RFE, PCA, 트리 중요도와, 리샘플링 네 가지, 즉 한 번 분할, k-fold, LOOCV, 반복 무작위 분할을 다룹니다.
Lab 2의 결론대로 로그와 표준화를 한 데이터에 로지스틱 회귀를 기준 모델로 썼습니다.
카이제곱 검정은 원본 데이터에서 소득을 1위로 뽑았는데, 점수가 값의 크기에 비례하기 때문입니다. 로그 변환 후에는 30–59일 연체, 연체 이력, 나이, 사용 비율 순이었습니다.
LOOCV는 15만 번 학습해야 해서 한 번 학습 시간을 재어 보니 약 20시간이 걸려, 1,000건 표본으로만 실행했습니다.""")

# =====================================================================
# Lab 3 - (2) 결과 : 방법별로 고른 변수 + 줄였을 때 AUC
# =====================================================================
s = new_slide()
lab_header(s, 3, 2, "특성 선택 · 리샘플링", "방법마다 고른 변수가 달랐고, 줄여도 AUC는 거의 그대로였다",
           terms=[("AUC", "연체자 구별 점수(0.5 찍기 ~ 1 완벽)"), ("± 값", "10번 시험 점수의 흔들림(표준편차)")])
rows_sel = [("사용 비율", "revol", 1, 1, 1), ("나이", "age", 1, 1, 1), ("연체 이력", "defaulted", 1, 1, 0),
            ("30–59일 연체", "late30", 1, 0, 0), ("부채 비율", "debt", 0, 0, 1), ("소득", "income", 0, 0, 1),
            ("부양가족", "depend", 0, 0, 0), ("대출 보유", "loans", 0, 0, 0)]
methods = ["통계 검정\n(로그 후 상위4)", "RFE\n(상위 3)", "트리 중요도\n(상위 4)"]
tparts = [rect(s, 0.8, 2.15, 5.75, 4.9, C["card"], radius=0.25),
          text(s, 1.1, 2.32, 5.2, 0.35, [[("방법별로 고른 변수  ", {"weight": "bold"}), ("● = 선택됨", {"color": C["mute"]})]],
               size=12, weight="medium")]
tbl = s.shapes.add_table(len(rows_sel) + 1, 5, Inches(1.05), Inches(2.78), Inches(5.25), Inches(0.42 * 9)).table
widths = [1.55, 1.15, 0.95, 1.05, 0.55]
for j, w in enumerate(widths):
    tbl.columns[j].width = Inches(w)
header = ["변수", *methods, "합계"]
for i in range(len(rows_sel) + 1):
    tbl.rows[i].height = Inches(0.42 if i else 0.52)
    for j in range(5):
        cell = tbl.cell(i, j)
        cell.margin_left = cell.margin_right = Inches(0.04)
        cell.margin_top = cell.margin_bottom = Inches(0.02)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.fill.solid()
        if i == 0:
            cell.fill.fore_color.rgb = rgb(C["blue_pale"])
            val, col, wt, sz = header[j], C["blue"], "bold", 10
        else:
            ko, en, a, b, c = rows_sel[i - 1]
            total = a + b + c
            cell.fill.fore_color.rgb = rgb(C["violet_pale"] if total == 3 else C["bg"])
            if j == 0:
                val, col, wt, sz = f"{ko}", C["ink"], "semibold", 11
            elif j == 4:
                val, col, wt, sz = str(total), C["violet"] if total == 3 else C["sub"], "bold", 12
            else:
                v = (a, b, c)[j - 1]
                val, col, wt, sz = ("●" if v else "·"), C["blue"] if v else C["line"], "bold", 13
        tf = cell.text_frame
        lines_ = str(val).split("\n")
        for k, ln in enumerate(lines_):
            p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
            p.alignment = PP_ALIGN.LEFT if (j == 0 and i > 0) else PP_ALIGN.CENTER
            r = p.add_run()
            r.text = ln
            set_font(r, sz if k == 0 else 9, wt if k == 0 else "medium", col if k == 0 else C["mute"])
tparts.append(s.shapes[-1])
group(s, tparts, "anim1_table")
group(s, [rect(s, 6.8, 2.15, 5.73, 3.55, C["card"], radius=0.25),
          text(s, 7.1, 2.32, 5.2, 0.35, [[("변수 수별 AUC  ", {"weight": "bold"}),
                                         ("트리 중요도 순서로 추가 · 층화 10-fold", {"color": C["mute"]})]], size=12, weight="medium"),
          s.shapes.add_picture(str(FIG / "lab3_subset_auc.png"), Inches(7.0), Inches(2.75), width=Inches(5.35))],
      "anim2_subset")
group(s, [rect(s, 6.8, 5.85, 5.73, 1.2, C["blue_pale"], radius=0.2),
          text(s, 7.1, 5.85, 5.2, 1.2,
               [[("상위 6개 0.826  vs  8개 0.824", {"weight": "extrabold", "color": C["ink"], "size": 16})],
                [("차이 0.002 < 흔들림 ±0.009 → ", {"color": C["sub"]}), ("8개 모두 사용", {"color": C["blue"], "weight": "bold"})]],
               size=13, weight="medium", anchor="m", line=1.35)], "anim3_keep")
notes(s, """왼쪽은 세 방법이 고른 변수입니다. 통계 검정은 로그 변환 후 상위 4개, RFE는 상위 3개, 트리 중요도는 상위 4개를 표시했습니다.
세 방법 모두 고른 변수는 사용 비율과 나이, 두 방법이 고른 변수는 연체 이력입니다. 트리 중요도는 값의 범위가 넓은 부채 비율과 소득을 높게 쳤습니다.
오른쪽은 트리 중요도 순서로 변수를 하나씩 더하며 층화 10-fold AUC를 잰 결과입니다. 상위 3–4개만 쓰면 0.75–0.78로 떨어지고, 5개부터 0.82대를 회복합니다.
상위 6개가 0.826으로 8개의 0.824보다 0.002 높지만, 이 차이는 10번 시험의 흔들림 0.009보다 작아 의미 있는 차이로 보기 어렵습니다. 그래서 8개를 모두 쓰기로 했습니다.
PCA 3개 성분도 0.822로 비슷했지만, 성분은 여러 변수를 섞은 값이라 설명이 어려워 쓰지 않았습니다.""")

# =====================================================================
# Lab 3 - (3) 발견 · 결론 : 리샘플링
# =====================================================================
s = new_slide()
lab_header(s, 3, 3, "특성 선택 · 리샘플링", "시험을 한 번만 나누면 AUC가 0.816~0.828로 흔들렸다",
           terms=[("k-fold", "10등분해 10번 시험"), ("층화", "모든 시험의 연체 비율을 6.68%로 고정"),
                  ("표준편차", "점수가 흔들리는 정도")])
group(s, [rect(s, 0.8, 2.15, 6.0, 4.05, C["card"], radius=0.25),
          text(s, 1.1, 2.33, 5.4, 0.35, [[("검증 방법별 AUC  ", {"weight": "bold", "color": C["ink"]}),
                                         ("점 하나 = 시험 한 번 · 선 = 평균", {"color": C["mute"]})]], size=13),
          s.shapes.add_picture(str(FIG / "lab3_resampling.png"), Inches(1.0), Inches(2.8), width=Inches(5.6))],
      "anim1_chart")
finds = [("01", "한 번 분할: 0.816 ~ 0.828", "나누는 방식(seed)만 바꿔도 0.012 차이", C["red"]),
         ("02", "층화 10-fold: 0.824 ± 0.009", "모든 시험의 연체 비율 6.68%로 고정 (일반은 6.45~7.18%)", C["blue"]),
         ("03", "LOOCV: 약 20시간 · AUC 계산 불가", "시험 데이터가 1명이라 순위를 매길 수 없음", C["blue"])]
parts = []
for i, (no, head, body, col) in enumerate(finds):
    yy = 2.15 + i * 0.72
    parts += [text(s, 7.1, yy, 0.6, 0.4, no, size=17, color=col, weight="extrabold"),
              text(s, 7.65, yy, 4.9, 0.4, head, size=16, weight="bold"),
              text(s, 7.65, yy + 0.34, 4.9, 0.35, body, size=12.5, color=C["sub"], weight="medium")]
group(s, parts, "anim2_finds")
conclusion_card(s, 7.05, 4.4, 5.48, 1.8, [
    [("① ", {"color": C["violet"], "weight": "bold"}), ("이후 모든 비교는 층화 10-fold AUC로 통일", {})],
    [("② ", {"color": C["violet"], "weight": "bold"}), ("변수는 8개 모두 사용 (줄여도 이득 < 흔들림)", {})],
], anim=3)
next_bar(s, "어떤 알고리즘이 연체자를 가장 잘 골라낼까?", "Lab 4", "AUC", "0.824 ± 0.009", anim=4)
notes(s, """검증 방법별로 나온 AUC를 점으로 찍었습니다. 점 하나가 시험 한 번, 가로선이 평균입니다.
데이터를 한 번만 나누면, 나누는 방식만 바꿔도 AUC가 0.816에서 0.828까지 0.012나 흔들렸습니다. 이 정도면 모델 간 차이와 구별하기 어렵습니다.
10번 나눠 시험하는 10-fold는 평균 0.824로 안정적이었습니다. 그중 층화 10-fold는 모든 시험의 연체 비율을 6.68%로 맞춰 주고, 흔들림은 0.009입니다. 일반 KFold는 시험마다 연체 비율이 6.45%에서 7.18%로 달랐습니다.
LOOCV는 약 20시간이 걸리고, 시험 데이터가 한 명뿐이라 AUC를 계산할 수 없었습니다.
목표 관점의 결론은, 이후 모든 알고리즘 비교를 층화 10-fold AUC로 통일하고, 변수는 8개를 모두 쓴다는 것입니다.
다음 질문은, 그럼 어떤 알고리즘이 연체자를 가장 잘 골라내는가입니다.""")

# =====================================================================
# Lab 4 - (1) 진행 : 평가지표 6가지 + 알고리즘 비교
# =====================================================================
s = new_slide()
lab_header(s, 4, 1, "평가지표 · 알고리즘 비교", "Lab 4의 평가지표 6가지와 알고리즘 6개 비교를 적용했다",
           terms=[("logloss", "확률이 틀린 정도(0 완벽)"), ("혼동행렬", "맞힘·틀림을 4칸으로 센 표"),
                  ("recall", "실제 연체자 중 잡아낸 비율")])
steps4 = [("정확도", "accuracy", "0.932"), ("logloss", "neg_log_loss", "0.199"),
          ("AUC", "roc_auc", "0.824"), ("ROC 곡선", "roc_curve", "시험 0.827"),
          ("혼동행렬", "confusion_matrix", "연체자 43명만"), ("분류 리포트", "precision·recall", "recall 0.013"),
          ("후보 점검", "LR·KNN·SVM·CART", "SVM 1시간+"), ("6개 비교", "boxplot", "LDA 0.832")]
parts = [text(s, 0.8, 1.98, 11.7, 0.3, [[("Lab 4 자료의 진행 순서", {"weight": "bold"}),
                                         ("   → 1–6 평가지표 · 7–8 알고리즘 비교, 실제 데이터 결과", {"color": C["mute"], "weight": "medium"})]],
              size=13)]
for i, (ko, fn, res) in enumerate(steps4):
    x = 0.8 + i * (sw3 + gap3)
    fill = C["blue_pale"] if i < 6 else C["violet_pale"]
    dot = C["blue"] if i < 6 else C["violet"]
    parts += [rect(s, x, 2.35, sw3, 1.18, fill, radius=0.16),
              oval(s, x + 0.13, 2.47, 0.28, dot),
              text(s, x + 0.13, 2.47, 0.28, 0.28, str(i + 1), size=10.5, color="#FFFFFF", weight="bold", align="c", anchor="m"),
              text(s, x + 0.47, 2.47, sw3 - 0.5, 0.28, ko, size=11.5, weight="bold", anchor="m"),
              text(s, x + 0.13, 2.84, sw3 - 0.2, 0.25, fn, size=9.5, color=dot, weight="medium", latin=MONO),
              rect(s, x + 0.1, 3.14, sw3 - 0.2, 0.3, C["bg"], radius=0.08),
              text(s, x + 0.1, 3.14, sw3 - 0.2, 0.3, res, size=10, color=C["red"] if i in (4, 5) else C["ink"],
                   weight="semibold", align="c", anchor="m")]
group(s, parts, "anim1_steps")
h4 = capture_card(s, 0.8, 3.7, 5.2, CAP / "lab4_confusion.png", "p.5–6 · 1 혼동행렬", anim=2, img_w=4.6)
capture_card(s, 0.8, 3.7 + h4 + 0.12, 5.2, CAP / "lab4_compare.png", "p.10 · 3 알고리즘 6개 비교", anim=2, img_w=3.6)
console(s, 6.2, 3.7, 6.33, 3.35, "실행 결과 · 04_lab4_metrics_algorithms.ipynb", [
    cmd("cross_val_score(.., 'accuracy')", "1 정확도"),
    [("0.932", "hi"), ("   ← '전부 정상'이라고 찍어도 0.933", "dim")],
    cmd("cross_val_score(.., 'neg_log_loss')", "2 logloss"),
    [("-0.199", "hi"), ("  ← 작을수록 좋아 부호만 뒤집어 보고", "dim")],
    cmd("cross_val_score(.., 'roc_auc')", "3 AUC"),
    [("0.824 (0.009)", "hi"), ("  ← 0.8~0.9 = good(B)", "dim")],
    cmd("confusion_matrix(Y_test, predicted)", "5 혼동행렬"),
    [("[[46116    75]   [ 3266    43]]", "red")],
    cmd("SVC().fit(X[:20000], Y[:20000])", "7 시간 측정"),
    [("8초 → 13.5만 건 약 6분 × 10-fold ≈ 1시간", "red")],
], size=10.5, anim=3, lh=15)
notes(s, """Lab 4는 평가지표 여섯 가지, 즉 정확도, logloss, AUC, ROC 곡선, 혼동행렬, 분류 리포트를 차례로 보고, 알고리즘 여섯 개를 같은 시험으로 비교합니다.
Lab 3의 결론대로 모든 점수는 층화 10-fold로 쟀고, ROC 곡선과 혼동행렬은 Lab 4 코드대로 33%를 시험 데이터로 떼어 냈습니다.
정확도는 0.932로 '전부 정상'이라고 찍는 0.933과 같았습니다. logloss는 0.199로, 늘 6.68%라고 말하는 모델의 0.245보다 낫습니다.
Lab 4는 알고리즘을 정확도로 비교하지만, 정확도로는 차이가 보이지 않아 AUC로 비교했습니다.
SVM은 학습 데이터가 두 배가 되면 시간이 약 네 배로 늘었습니다. 2만 건에 8초였으니 13만 5천 건이면 약 6분, 10-fold면 약 한 시간이라, 같은 1만 건 표본에서 여섯 개를 모두 돌려 비교했습니다.""")

# =====================================================================
# Lab 4 - (2) 결과 : 혼동행렬 · 판정 기준 · 가중치
# =====================================================================
s = new_slide()
lab_header(s, 4, 2, "평가지표 · 알고리즘 비교", "정확도는 93%인데, 연체자 3,309명 중 43명(1.3%)만 잡았다",
           terms=[("2종 오류", "연체자를 통과시킴"), ("판정 기준", "이 확률을 넘으면 연체로 판정"),
                  ("가중치", "연체 1명 = 정상 14명으로 학습")])


def cm_block(s, x, y, title, sub, cm, caught_color):
    """혼동행렬 2×2. cm = (tn, fp, fn, tp)"""
    tn, fp, fn, tp = cm
    cw, chh, lw, g = 1.05, 0.72, 0.6, 0.06
    out = [text(s, x, y, 2.9, 0.3, [[(title, {"weight": "bold"}), (f"  {sub}", {"color": C["mute"], "size": 10.5})]], size=12.5)]
    y0 = y + 0.42
    for j, lab in enumerate(["예측 정상", "예측 연체"]):
        out.append(text(s, x + lw + j * (cw + g), y0, cw, 0.25, lab, size=9.5, color=C["mute"], weight="semibold", align="c"))
    y1 = y0 + 0.3
    cells = [[(tn, "맞게 통과", C["bg"], C["sub"]), (fp, "잘못 거절", C["bg"], C["sub"])],
             [(fn, "놓친 연체자", C["red_pale"], C["red"]), (tp, "잡은 연체자", C["blue_pale"], C["blue"])]]
    for i, lab in enumerate(["실제\n정상", "실제\n연체"]):
        yy = y1 + i * (chh + g)
        out.append(text(s, x, yy, lw - 0.08, chh, lab.split("\n"), size=9.5, color=C["mute"], weight="semibold",
                        align="c", anchor="m", line=1.05))
        for j, (v, lab2, fill, col) in enumerate(cells[i]):
            xx = x + lw + j * (cw + g)
            out += [rect(s, xx, yy, cw, chh, fill, radius=0.1),
                    text(s, xx, yy + 0.08, cw, 0.36, f"{v:,}", size=16, color=col if i else C["ink"], weight="extrabold", align="c"),
                    text(s, xx, yy + 0.43, cw, 0.22, lab2, size=9, color=col, weight="medium", align="c")]
    ye = y1 + 2 * chh + g + 0.14
    out.append(text(s, x, ye, lw + 2 * cw + g, 0.36,
                    [[(f"연체자 {tp:,}명 잡음 · {tp / (tp + fn):.1%}", {"color": caught_color, "weight": "extrabold"})]],
                    size=15, align="c"))
    out.append(text(s, x, ye + 0.36, lw + 2 * cw + g, 0.3, f"대신 정상 고객 {fp:,}명 거절", size=11,
                    color=C["sub"], weight="medium", align="c"))
    return out


cparts = [rect(s, 0.8, 2.15, 6.3, 4.9, C["card"], radius=0.25),
          text(s, 1.1, 2.32, 5.8, 0.35, [[("혼동행렬  ", {"weight": "bold"}),
                                         ("시험 데이터 49,500명 · 실제 연체자 3,309명", {"color": C["mute"]})]], size=12, weight="medium")]
cparts += cm_block(s, 1.05, 2.82, "기본 (기준 0.5)", "Lab 4 그대로", (46116, 75, 3266, 43), C["red"])
cparts += cm_block(s, 4.1, 2.82, "연체자 가중치", "balanced", (36425, 9766, 831, 2478), C["blue"])
group(s, cparts, "anim1_cm")
group(s, [rect(s, 1.05, 5.95, 5.8, 0.88, C["bg"], radius=0.16),
          text(s, 1.3, 5.95, 5.4, 0.88,
               [[("연체 판정 중 진짜 연체자 20%", {"weight": "bold", "color": C["ink"]}),
                 ("  무작위로 고르면 6.7% → 3배", {"color": C["mute"]})],
                [("가중치를 주자 AUC도 ", {"color": C["sub"]}), ("0.824 → 0.844", {"color": C["violet"], "weight": "extrabold"}),
                 ("  (층화 10-fold)", {"color": C["mute"]})]],
               size=12.5, weight="medium", anchor="m", line=1.3)], "anim2_tradeoff")
group(s, [rect(s, 7.3, 2.15, 5.23, 4.9, C["card"], radius=0.25),
          text(s, 7.6, 2.32, 4.8, 0.35, [[("ROC 곡선  ", {"weight": "bold"}),
                                         ("판정 기준에 따라 잡는 연체자 비율", {"color": C["mute"]})]], size=12, weight="medium"),
          s.shapes.add_picture(str(FIG / "lab4_roc.png"), Inches(7.62), Inches(2.8), width=Inches(4.6))],
      "anim3_roc")
notes(s, """왼쪽은 Lab 4 코드대로 시험 데이터 49,500명에 대한 혼동행렬입니다. 실제 연체자는 3,309명입니다.
기본 판정 기준 0.5에서는 연체자를 43명, 1.3%만 잡았습니다. 정확도 93%는 거의 모두를 정상이라고 말한 결과입니다.
Lab 4 질문 (j)대로 오류의 비용을 보면, 정상 고객을 거절하는 1종 오류는 이자 수익 손실이고, 연체자를 통과시키는 2종 오류는 원금 손실이라 2종 오류가 더 비쌉니다.
그래서 두 가지를 해 봤습니다. 판정 기준을 0.2로 낮추면 51%, 학습할 때 연체자 1명을 정상 약 14명만큼 중요하게 보는 가중치를 주면 75%를 잡습니다.
대신 정상 고객 9,766명을 거절하게 되는데, 연체로 판정한 사람 중 20%가 진짜 연체자라 무작위의 3배입니다.
그리고 가중치를 주자 판정 기준만이 아니라 순위 자체가 좋아져, 층화 10-fold AUC가 0.824에서 0.844로 올랐습니다.
오른쪽 ROC 곡선에서 빨간 점이 판정 기준별 위치이고, 보라색이 가중치를 준 모델입니다.""")

# =====================================================================
# Lab 4 - (3) 발견 · 결론 : 알고리즘 6개 비교
# =====================================================================
s = new_slide()
lab_header(s, 4, 3, "평가지표 · 알고리즘 비교", "알고리즘 6개 중 선형 계열이 가장 잘 골라냈다 (LDA 0.832)",
           terms=[("LDA", "두 집단의 평균 차이로 나누는 직선 모델"), ("NB", "변수를 서로 독립으로 보는 확률 모델"),
                  ("KNN", "가장 가까운 5명의 다수결")])
group(s, [rect(s, 0.8, 2.15, 6.0, 4.05, C["card"], radius=0.25),
          text(s, 1.1, 2.33, 5.4, 0.35, [[("알고리즘별 AUC  ", {"weight": "bold", "color": C["ink"]}),
                                         ("층화 10-fold · 상자 = 10번 시험의 범위", {"color": C["mute"]})]], size=13),
          s.shapes.add_picture(str(FIG / "lab4_compare.png"), Inches(1.0), Inches(2.78), width=Inches(5.6)),
          text(s, 1.1, 5.72, 5.5, 0.3, "* SVM은 1만 건 표본 · 나머지는 15만 건 전체 · 정확도는 모두 88~93%로 차이 없음",
               size=10, color=C["mute"], weight="medium")],
      "anim1_chart")
finds = [("01", "LDA 0.832 > LR 0.824 ≈ NB 0.822", "직선에 가까운 관계 → 선형 모델이 강함", C["blue"]),
         ("02", "KNN 0.727 · 결정트리 0.595", "KNN은 확률이 6단계뿐, 트리는 과적합", C["red"]),
         ("03", "SVM 0.579 → 가중치 0.810", "15만 건은 10-fold 약 1시간 → 1만 건 표본", C["red"])]
parts = []
for i, (no, head, body, col) in enumerate(finds):
    yy = 2.15 + i * 0.72
    parts += [text(s, 7.1, yy, 0.6, 0.4, no, size=17, color=col, weight="extrabold"),
              text(s, 7.65, yy, 4.9, 0.4, head, size=16, weight="bold"),
              text(s, 7.65, yy + 0.34, 4.9, 0.35, body, size=12.5, color=C["sub"], weight="medium")]
group(s, parts, "anim2_finds")
conclusion_card(s, 7.05, 4.4, 5.48, 1.8, [
    [("① ", {"color": C["violet"], "weight": "bold"}), ("연체자 가중치를 기본으로 (잡는 비율 1.3% → 75%)", {})],
    [("② ", {"color": C["violet"], "weight": "bold"}), ("후보는 LR·LDA, 트리는 앙상블로 다시 (Lab 5)", {})],
], anim=3)
next_bar(s, "지금까지의 평가는 공정했을까? 더 올릴 수 있을까?", "Lab 5", "최고 AUC", "0.844", anim=4)
notes(s, """Lab 4 방식대로 여섯 알고리즘을 같은 층화 10-fold, 같은 seed로 비교했습니다. 정확도는 모두 88에서 93%로 기준선 근처라 AUC로 봤습니다.
LDA가 0.832로 가장 높았고, 로지스틱 회귀 0.824, 나이브 베이즈 0.822가 뒤를 이었습니다. 직선 경계를 쓰는 모델이 강하다는 건, 변수와 연체의 관계가 비교적 단순하다는 뜻입니다.
다만 LDA와 LR의 차이 0.008은 흔들림과 비슷한 크기라 확실한 우위라고 보긴 어렵습니다.
KNN은 기본 이웃 수가 5라 확률이 여섯 단계뿐이어서 줄 세우기가 거칠고, 결정트리는 깊이 제한이 없어 학습 데이터를 외워 0.595에 그쳤습니다.
SVM은 15만 건으로는 10-fold에 약 한 시간이 걸려 1만 건 표본으로 비교했는데, 0.579로 거의 찍기 수준이었습니다. 정상이 93%라 거의 모두 정상인 경계를 그었기 때문이고, 연체자 가중치를 주면 0.810으로 회복됐습니다.
목표 관점의 결론은, 연체자 가중치를 기본으로 쓰고, 후보는 LR과 LDA, 트리는 하나로는 약하니 Lab 5의 앙상블로 다시 본다는 것입니다.
다음 질문은, 지금까지의 평가가 공정했는지, 그리고 더 올릴 수 있는지입니다. 사실 지금까지는 로그와 표준화를 시험 데이터까지 포함한 전체에 먼저 맞췄습니다.""")

# =====================================================================
# Lab 5 - (1) 진행 : 파이프라인 · 앙상블 · 튜닝 · 저장
# =====================================================================
# 노트북 05 실행 결과 (results.md 와 같은 값)
L5 = dict(grid_leaf=200, grid_auc=0.8555, rand_leaf=176, rand_auc=0.8555, final_leaf=176, final_auc=0.8556,
          final_std=0.0070, leaf1_auc=0.8244, mb=3.7, test_auc=0.8558, fold_diff=0.001, rf_noclean=0.8555)
if L5["final_auc"] is None:  # 노트북 05 결과를 채우기 전에는 Lab 4까지만 만든다
    prs.save(DRAFT)
    write_script_md(ROOT / "slides" / "script.md")
    print("saved (Lab 5 결과 대기 중)", DRAFT)
    raise SystemExit
s = new_slide()
lab_header(s, 5, 1, "파이프라인 · 앙상블 · 튜닝", "Lab 5의 파이프라인 · 앙상블 · 튜닝 · 저장을 차례로 적용했다",
           terms=[("파이프라인", "전처리와 모델을 한 묶음으로"), ("앙상블", "여러 모델을 합쳐 예측"),
                  ("튜닝", "설정값을 바꿔 가며 최고 찾기")])
steps5 = [("Pipeline", "Pipeline", "0.8436 그대로"), ("FeatureUnion", "PCA + KBest", "0.8436"),
          ("Bagging", "트리 100개", "0.819"), ("랜덤 포레스트", "RF(100, 3)", "0.824"),
          ("AdaBoost", "트리 30개", "0.847"), ("Voting", "LR+CART+NB", "0.830"),
          ("튜닝", "Grid · Random", f"RF {L5['final_auc']:.3f}"), ("모델 저장", "pickle · joblib", f"{L5['mb']:.1f} MB")]
parts = [text(s, 0.8, 1.98, 11.7, 0.3, [[("Lab 5 자료의 진행 순서", {"weight": "bold"}),
                                         ("   → 1–2 파이프라인 · 3–6 앙상블 · 7–8 튜닝·저장, 실제 데이터 결과", {"color": C["mute"], "weight": "medium"})]],
              size=13)]
for i, (ko, fn, res) in enumerate(steps5):
    x = 0.8 + i * (sw3 + gap3)
    ens_step = 2 <= i <= 5
    fill = C["violet_pale"] if ens_step else C["blue_pale"]
    dot = C["violet"] if ens_step else C["blue"]
    parts += [rect(s, x, 2.35, sw3, 1.18, fill, radius=0.16),
              oval(s, x + 0.13, 2.47, 0.28, dot),
              text(s, x + 0.13, 2.47, 0.28, 0.28, str(i + 1), size=10.5, color="#FFFFFF", weight="bold", align="c", anchor="m"),
              text(s, x + 0.47, 2.47, sw3 - 0.5, 0.28, ko, size=11.5 if len(ko) < 7 else (10.5 if len(ko) < 10 else 9.5), weight="bold", anchor="m"),
              text(s, x + 0.13, 2.84, sw3 - 0.2, 0.25, fn, size=9.5, color=dot, weight="medium", latin=MONO),
              rect(s, x + 0.1, 3.14, sw3 - 0.2, 0.3, C["bg"], radius=0.08),
              text(s, x + 0.1, 3.14, sw3 - 0.2, 0.3, res, size=10, color=C["blue"] if i == 6 else C["ink"],
                   weight="semibold", align="c", anchor="m")]
group(s, parts, "anim1_steps")
h5 = capture_card(s, 0.8, 3.7, 5.2, CAP / "lab5_pipeline.png", "p.2 · 1 Pipeline", anim=2, img_w=4.4)
capture_card(s, 0.8, 3.7 + h5 + 0.12, 5.2, CAP / "lab5_rf.png", "p.4 · 2 Random Forest", anim=2, img_w=4.4)
console(s, 6.2, 3.7, 6.33, 3.35, "실행 결과 · 05_lab5_pipeline_ensemble_tuning.ipynb", [
    cmd("Pipeline([log, standardize, lr])", "1 누수 막기"),
    [("0.8436 → 0.8436", "hi"), ("   ← 전체에 먼저 맞춘 값과 같음", "dim")],
    cmd("isin(late30, [96, 98])", "+ 데이터 정리"),
    [("269행 · 연체율 54.6%", "red"), ("   ← 연체 이력 변수에 이미 담김", "dim")],
    cmd("RandomForestClassifier(100, 3)", "4 RF"),
    [("0.824", "hi"), ("   ← 트리가 끝까지 자라 외움", "dim")],
    cmd("GridSearchCV(rf, min_samples_leaf)", "7 튜닝"),
    [(f"leaf={L5['grid_leaf']} → {L5['grid_auc']:.3f}", "hi"), (f"   ← leaf=1이면 {L5['leaf1_auc']:.3f}", "dim")],
    cmd("joblib.dump(final, '...sav')", "8 저장"),
    [(f"{L5['mb']:.1f} MB · 다시 불러와도 AUC {L5['test_auc']:.3f}", "hi")],
], size=10.5, anim=3, lh=15)
notes(s, f"""Lab 5는 파이프라인, 앙상블, 튜닝, 모델 저장을 다룹니다.
먼저 Lab 4까지의 약점을 바로잡았습니다. 지금까지는 로그와 표준화를 시험 데이터까지 포함한 전체에 먼저 맞췄는데, 파이프라인으로 묶으면 각 fold의 학습 데이터로만 맞춥니다.
여기에 Lab 1에서 찾아 두었던 문제 데이터, 즉 30–59일 연체의 96·98 코드와 소득 대체값 6638·6746을 정리하는 단계도 파이프라인에 넣었습니다.
앙상블은 Lab 5 순서대로 결정트리 100개 Bagging, 변수 3개씩 고르는 랜덤 포레스트 100개, AdaBoost 30개, 그리고 Voting을 비교했습니다. Voting의 SVM은 Lab 4에서 본 대로 15만 건에 너무 느려 나이브 베이즈로 바꿨습니다.
튜닝은 랜덤 포레스트의 잎 최소 인원을 Grid와 Random Search로 찾았고, 최종 모델을 pickle과 joblib으로 저장한 뒤 다시 불러와 시험했습니다.""")

# =====================================================================
# Lab 5 - (2) 결과 : 누수 확인 + 앙상블
# =====================================================================
s = new_slide()
lab_header(s, 5, 2, "파이프라인 · 앙상블 · 튜닝", "누수를 막아도 결과는 같았고, 트리는 합치자 0.597 → 0.847로 올랐다",
           terms=[("누수", "시험 데이터 정보가 학습에 섞임"), ("Bagging", "다른 표본으로 만든 트리의 평균"),
                  ("AdaBoost", "앞 트리의 실수를 다음 트리가 고침")])
rows5 = [("", "전체에 먼저 맞춤", "파이프라인", "+ 데이터 정리"),
         ("LR 가중치", "0.8436", "0.8436", "0.8436"),
         ("LDA", "0.8318", "0.8318", "0.8325"),
         ("RF 튜닝", "변환 불필요", f"{L5['rf_noclean']:.4f}", f"{L5['final_auc']:.4f}")]
lparts = [rect(s, 0.8, 2.15, 5.6, 4.9, C["card"], radius=0.25),
          text(s, 1.1, 2.32, 5.0, 0.35, [[("파이프라인 · 데이터 정리  ", {"weight": "bold"}),
                                         ("층화 10-fold AUC", {"color": C["mute"]})]], size=12, weight="medium")]
tbl = s.shapes.add_table(4, 4, Inches(1.05), Inches(2.8), Inches(5.1), Inches(1.9)).table
for j, w in enumerate([1.2, 1.4, 1.2, 1.3]):
    tbl.columns[j].width = Inches(w)
for i, row in enumerate(rows5):
    tbl.rows[i].height = Inches(0.46)
    for j, val in enumerate(row):
        cell = tbl.cell(i, j)
        cell.margin_left = cell.margin_right = Inches(0.06)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.fill.solid()
        cell.fill.fore_color.rgb = rgb(C["blue_pale"] if i == 0 else C["bg"])
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER
        r = p.add_run()
        r.text = val
        if i == 0:
            set_font(r, 10, "bold", C["blue"])
        else:
            set_font(r, 12 if j else 11.5, "semibold", C["ink"] if j else C["ink"], latin=None)
lparts.append(s.shapes[-1])
lparts += [rect(s, 1.05, 4.9, 5.1, 0.95, C["bg"], radius=0.16),
           text(s, 1.3, 4.9, 4.7, 0.95,
                [[("누수 영향 없음", {"weight": "bold", "color": C["blue"]}), ("  15만 건이라 fold마다", {"color": C["sub"]})],
                 [(f"평균 차이가 표준편차의 {L5['fold_diff']:.1%} 이하", {"color": C["sub"]})]],
                size=12.5, weight="medium", anchor="m", line=1.3),
           rect(s, 1.05, 5.97, 5.1, 0.88, C["bg"], radius=0.16),
           text(s, 1.3, 5.97, 4.7, 0.88,
                [[("데이터 정리 효과 ≈ 0", {"weight": "bold", "color": C["ink"]}), ("  96/98 코드는 연체율 54.6%지만", {"color": C["sub"]})],
                 [("'연체 이력' 변수에 이미 담겨 있었다", {"color": C["sub"]})]],
                size=12.5, weight="medium", anchor="m", line=1.3)]
group(s, lparts, "anim1_pipe")
group(s, [rect(s, 6.6, 2.15, 5.93, 4.9, C["card"], radius=0.25),
          text(s, 6.9, 2.32, 5.4, 0.35, [[("모델별 AUC  ", {"weight": "bold"}),
                                         ("층화 10-fold · 회색 = 트리 앙상블(기본 설정)", {"color": C["mute"]})]], size=12, weight="medium"),
          s.shapes.add_picture(str(FIG / "lab5_models.png"), Inches(6.78), Inches(2.85), width=Inches(5.6))],
      "anim2_models")
group(s, [rect(s, 6.85, 5.97, 5.43, 0.88, C["blue_pale"], radius=0.16),
          text(s, 7.1, 5.97, 5.0, 0.88,
               [[("트리 1개 0.597", {"weight": "bold", "color": C["ink"]}), ("  학습 데이터를 외움", {"color": C["sub"]})],
                [("→ 합치면(Bagging·RF) 0.82 · 이어 붙이면(AdaBoost) ", {"color": C["sub"]}),
                 ("0.847", {"color": C["blue"], "weight": "extrabold"})]],
               size=12.5, weight="medium", anchor="m", line=1.3)], "anim3_ens")
notes(s, f"""왼쪽은 파이프라인 결과입니다. 로그와 표준화를 전체 데이터에 먼저 맞췄을 때와, 파이프라인으로 각 fold 안에서만 맞췄을 때 AUC가 소수 넷째 자리까지 같았습니다.
15만 건이라 fold마다 표준화에 쓰는 평균이 거의 같기 때문입니다. 차이가 표준편차의 {L5['fold_diff']:.1%} 이하였습니다. 그래서 Lab 2에서 4까지의 결과는 그대로 유효합니다.
Lab 1에서 찾은 96, 98 코드는 269행으로 연체율이 54.6%나 되지만, 같은 고객이 '연체 이력' 변수에 이미 1로 잡혀 있어 정리해도 거의 변화가 없었습니다.
오른쪽은 앙상블입니다. 결정트리 하나는 0.597였지만, 100개를 합친 Bagging은 0.819, 랜덤 포레스트는 0.824로 올랐습니다. 서로 다른 표본으로 만든 트리의 실수가 평균으로 상쇄되기 때문입니다.
얕은 트리를 이어 붙여 앞의 실수를 고치는 AdaBoost는 0.847로 가장 높았고, Voting은 과적합된 결정트리가 평균을 끌어내려 0.830이었습니다.""")

# =====================================================================
# Lab 5 - (3) 발견 · 결론 : 튜닝 + 저장
# =====================================================================
s = new_slide()
lab_header(s, 5, 3, "파이프라인 · 앙상블 · 튜닝", f"트리가 외우지 못하게 막자 랜덤 포레스트가 {L5['final_auc']:.3f}으로 최고였다",
           terms=[("min_samples_leaf", "잎 하나에 남아야 할 최소 인원"), ("Grid", "정해 둔 값을 모두 시험"),
                  ("Random", "무작위로 골라 시험")])
group(s, [rect(s, 0.8, 2.15, 6.0, 4.05, C["card"], radius=0.25),
          text(s, 1.1, 2.33, 5.4, 0.35, [[("랜덤 포레스트 튜닝  ", {"weight": "bold", "color": C["ink"]}),
                                         ("잎 최소 인원에 따른 AUC", {"color": C["mute"]})]], size=13),
          s.shapes.add_picture(str(FIG / "lab5_tuning.png"), Inches(1.45), Inches(2.75), width=Inches(4.7))],
      "anim1_chart")
finds = [("01", f"Grid: leaf {L5['grid_leaf']} → {L5['grid_auc']:.3f}", f"leaf 1(기본)은 {L5['leaf1_auc']:.3f} · 클수록 덜 외움", C["blue"]),
         ("02", f"Random 10번: leaf {L5['rand_leaf']} → {L5['rand_auc']:.3f}", "격자 사이 값을 찾았지만 차이는 흔들림 안", C["blue"]),
         ("03", f"저장 후 다시 불러와도 AUC {L5['test_auc']:.3f}", f"정리 단계까지 한 파일 · {L5['mb']:.1f} MB", C["blue"])]
parts = []
for i, (no, head, body, col) in enumerate(finds):
    yy = 2.15 + i * 0.72
    parts += [text(s, 7.1, yy, 0.6, 0.4, no, size=17, color=col, weight="extrabold"),
              text(s, 7.65, yy, 4.9, 0.4, head, size=16, weight="bold"),
              text(s, 7.65, yy + 0.34, 4.9, 0.35, body, size=12.5, color=C["sub"], weight="medium")]
group(s, parts, "anim2_finds")
conclusion_card(s, 7.05, 4.4, 5.48, 1.8, [
    [("① ", {"color": C["violet"], "weight": "bold"}), (f"최종 후보: 정리 + RF(잎 최소 {L5['final_leaf']}명), AUC {L5['final_auc']:.3f}", {})],
    [("② ", {"color": C["violet"], "weight": "bold"}), (f"MLP는 이 기준 {L5['final_auc']:.3f}을 넘어야 의미가 있다", {})],
], anim=3)
next_bar(s, "신경망(MLP)은 15만 건에서 더 잘 골라낼까?", "Lab 6", "최고 AUC", f"{L5['final_auc']:.3f}", anim=4)
notes(s, f"""랜덤 포레스트의 트리는 기본 설정에서 잎에 한 명이 남을 때까지 자라 학습 데이터를 외웁니다. 그래서 잎 하나에 남아야 할 최소 인원을 Lab 5의 Grid Search와 Random Search로 찾았습니다.
Grid는 1, 10, 50, 100, 200 다섯 값을 모두 시험해 {L5['grid_leaf']}에서 {L5['grid_auc']:.3f}이었고, 기본값 1은 {L5['leaf1_auc']:.3f}이었습니다. Random은 1에서 300 사이를 무작위로 10번 골라 {L5['rand_leaf']}에서 {L5['rand_auc']:.3f}을 찾았습니다.
두 방법의 차이는 흔들림보다 작아, 이 문제에서는 잎 최소 인원을 수십에서 백 명대로 두면 된다는 뜻입니다.
최종 모델은 데이터 정리와 잎 최소 {L5['final_leaf']}명 랜덤 포레스트를 묶은 파이프라인으로, 층화 10-fold AUC {L5['final_auc']:.3f}입니다. Lab 5 방식대로 pickle과 joblib으로 저장하고 다시 불러와도 시험 AUC가 {L5['test_auc']:.3f}로 같았습니다.
목표 관점의 결론은, 지금까지의 최고 기준이 {L5['final_auc']:.3f}이고, 다음 Lab 6의 신경망은 이 기준을 넘어야 의미가 있다는 것입니다.""")

prs.save(DRAFT)
print("saved", DRAFT)
write_script_md(ROOT / "slides" / "script.md")
print("saved", ROOT / "slides" / "script.md")
