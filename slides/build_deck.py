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
        "신용위험 예측 모델링 발표를 시작하겠습니다."),
    2: ("목차", 15,
         '데이터와 목표는 과제에서 정해진 그대로입니다. 발표는 Lab 1부터 6까지 찾은 결과를 다음 Lab으로 넘기며 최종 모델을 만들고, 실무에서 쓸 수 있는지까지 확인하는 순서입니다.'),
    3: ("Lab 1 · 진행", 10,
         'Lab 1의 일곱 단계를 15만 명 데이터에 그대로 적용했고, 첫 줄이 열 이름이라 header=0만 더했습니다.'),
    4: ("Lab 1 · 결과", 20,
         "연체 고객은 100명 중 7명꼴이라, 전부 '정상'이라고 답해도 93%를 맞힙니다. 또 한도의 5만 배를 썼다는 값, 같은 숫자로 채운 빈 소득, 나올 수 없는 연체 횟수가 곳곳에 있었습니다."),
    5: ("Lab 1 · 발견·결론", 23,
         '사용 비율은 숫자 그대로 보면 연체와 관계가 거의 없었지만, 순위로 보면 관계가 세 번째로 강했습니다. 극단값의 영향일 수 있어, 다음 단계에서 변환 효과를 확인하겠습니다. 또 연체자가 적어 정확도만으로는 평가가 어렵습니다. 그래서 연체자에게 정상 고객보다 높은 위험 점수를 주는지 보는 AUC를 쓰겠습니다.'),
    6: ("Lab 2 · 진행", 10,
         'Lab 2 순서대로 값이 어떻게 퍼져 있는지 그려 보고, Lab 2의 변환 네 가지에 로그를 더해 적용했습니다.'),
    7: ("Lab 2 · 결과", 20,
         '그려 보니 한도 사용 비율, 빚 부담, 소득은 대부분 한 칸에 몰려 있었습니다. 숫자를 자릿수로 보는 로그 변환을 하자, 빚 부담의 쏠림이 95에서 1.7로 줄며 펴졌습니다. 연체 횟수는 96, 98 코드 때문에 여전히 쏠려 있었습니다.'),
    8: ("Lab 2 · 발견·결론", 23,
         '같은 변환도 모델에 따라 효과가 정반대였습니다. 로그와 표준화는 로지스틱 회귀만 0.82로 올렸고, 값을 0과 1로 나누는 Binarizer는 결정트리만 0.78로 올렸습니다. 그래서 회귀와 MLP에는 로그와 표준화를 씁니다. 그럼 변수 8개가 다 필요할까요?'),
    9: ("Lab 3 · 진행", 10,
         'Lab 3의 변수 고르는 방법 네 가지와 시험 나누는 방법 네 가지를 적용했고, k-fold는 연체자 비율을 맞추는 층화로 바꿨습니다.'),
    10: ("Lab 3 · 결과", 20,
         '방법마다 고른 변수가 달랐습니다. 통계 검정은 숫자가 큰 소득을, 트리는 값의 종류가 많은 변수를 높게 쳤고, PCA는 변수를 고르지 않고 3개로 압축해 따로 비교했습니다. 6개로 줄이면 로지스틱 회귀는 0.002 올랐지만, 다른 모델에서는 다를 수 있어 6개와 8개를 모두 후보로 남겼습니다.'),
    11: ("Lab 3 · 발견·결론", 23,
         '시험용 데이터를 한 번만 나누면, 나누는 방식에 따라 AUC가 0.816에서 0.828까지 달라졌습니다. 그래서 10조각으로 돌아가며 시험하되 조각마다 연체자 비율을 맞추는 층화 10-fold로 통일했습니다. 그럼 어떤 알고리즘이 가장 잘 골라낼까요?'),
    12: ("Lab 4 · 진행", 10,
         'Lab 4의 채점 지표와 알고리즘을 각각 여섯 개씩, 같은 층화 10-fold로 비교했습니다.'),
    13: ("Lab 4 · 결과", 20,
         '정확도는 93%였지만, 혼동행렬로 맞힌 것과 틀린 것을 나눠 보니 연체자 3,309명 중 고작 43명만 잡았습니다. 은행에는 연체자를 놓치는 게 더 큰 손해라, 연체자를 더 무겁게 보는 가중치를 주자 75%를 잡았고 AUC도 0.844로 올랐습니다.'),
    14: ("Lab 4 · 발견·결론", 23,
         '알고리즘 여섯 개 중 LDA가 0.832로 가장 높았고, 결정트리는 학습 데이터를 외워 0.6에 그쳤습니다. 그래서 가중치를 기본으로 쓰고, 선형 모델은 후보로 두되 트리는 여러 개를 합치는 앙상블로 다시 보겠습니다. 그런데 지금까지의 평가, 공정했을까요?'),
    15: ("Lab 5 · 진행", 10,
         'Lab 5 순서대로 파이프라인, 앙상블, 튜닝, 저장을 적용했고, Lab 1의 이상값 정리도 넣었습니다.'),
    16: ("Lab 5 · 결과", 20,
         '지금까지는 크기 맞추기에 시험 데이터까지 썼는데, 시험 문제를 엿본 셈이라 누수라고 합니다. 이를 막아도 AUC가 같아 영향은 없었습니다. 결정트리 하나는 0.6이었지만, 여러 개를 합치자 0.85까지 올랐습니다.'),
    17: ("Lab 5 · 발견·결론", 23,
         'Lab 4·5의 좋은 후보 네 개를 같은 조건으로 튜닝했습니다. 최고는 잎 최소 인원을 튜닝한 랜덤 포레스트 0.856으로, 기본 설정 1등이던 AdaBoost보다 열 조각 모두 높았습니다. Lab 3에서 남긴 변수도 8개로 확정했습니다. 그럼 신경망은 이 기준을 넘을 수 있을까요?'),
    18: ("Lab 6 · 진행", 10,
         'Lab 6 구조 그대로 MLP를 만들었는데, 10명씩 묶어 학습하면 26분이 걸려 몇 명씩 묶을지, 즉 배치 크기부터 쟀습니다.'),
    19: ("Lab 6 · 결과", 20,
         '배치 다섯 가지를 150 에폭, 즉 데이터를 150번 볼 때까지 재 보니 AUC는 0.85대로 거의 같았지만 시간은 최대 99배 차이 났습니다. 그래서 성능은 같으면서 배치 10보다 13배 빠른 128을 골랐습니다.'),
    20: ("Lab 6 · 발견·결론", 23,
         'MLP는 AUC 0.850으로, 튜닝한 랜덤 포레스트 0.856보다 조금 낮았고 학습은 10배 느렸습니다. 변수 8개짜리 표 데이터에서는 트리를 합친 모델이 더 효율적이라, Lab 범위의 최종 모델은 랜덤 포레스트로 정했습니다.'),
    21: ("최종 모델 비교", 20,
         "Lab 과정으로 만든 랜덤 포레스트와, Lab 밖 논문 방법인 새 변수와 부스팅을 더한 최적 모델을 같은 방법으로 비교했습니다. 최적 모델이 0.858로 열 조각 모두에서 높았지만 차이는 0.002였습니다. 그럼 이렇게 만든 최종 모델, 실무에서도 쓸 수 있을까요?"),
    22: ("실무 점검 · 결론", 45,
         '이렇게 만든 최종 모델을 실무에서도 쓸 수 있을지 궁금했습니다. 앞에서 보았듯이 AUC가 0.844로 좋을 때도, 연체로 판정한 사람 중 진짜 연체자는 20%뿐이었습니다. 정상 고객이 워낙 많아서, 우량 고객을 대거 잘못 판단해도 ROC-AUC는 높게 나오는 한계가 있었습니다. 그래서 정밀도를 보는 PR-AUC로 다시 확인하니, 판정 기준 0.17에서도 판정한 세 명 중 두 명은 정상 고객이었습니다. 대신 위험 상위 10%만 심사해도 연체자의 절반을 잡을 수 있어, 혼자 판정하기보다 위험 고객을 먼저 골라 주는 선별 도구로 쓸 수 있다고 판단했습니다. 데이터, 전처리, 알고리즘, 평가 지표를 바꿀 때마다 결과가 달라진 점이 가장 흥미로웠고, 신경망이라고 당연히 좋은 것은 아니었습니다.'),
}

# ---------------- 예상 질문 (교수님은 PPT만 보고 질문) ----------------
QA = {
    2: [('왜 ROC-AUC를 주 지표로 썼나?',
          "연체가 6.7%뿐이라 전부 '정상'이라고 해도 정확도가 93%다. Lab 4와 Kaggle 대회 기준에 맞춰 ROC-AUC로 비교했고, 실무에서 쓸 수 있는지는 22장에서 PR-AUC로 다시 확인했다.")],
    3: [('header=0은 왜 넣었나?',
          'Lab 자료는 헤더 없는 CSV 기준이다. 이 파일은 첫 줄이 열 이름이라 그대로 읽으면 열 이름 줄까지 데이터로 들어가 150,001행이 되고 모든 열이 문자열이 된다.')],
    4: [('6638, 6746이 결측을 채운 값이라는 근거는?',
          '두 값이 합계 29,737행으로 원본 Kaggle 데이터의 소득 결측 수(29,731)와 거의 같고, 그중 27,904행은 부채 비율도 1을 넘는다.')],
    5: [('Pearson과 Spearman의 차이는?',
          'Pearson은 값 그대로의 직선 관계라 극단값에 크게 흔들리고, Spearman은 값을 순위로 바꿔 계산해 극단값 영향이 적다.')],
    6: [('로그 변환은 Lab에 없는데 왜 했나?',
          "Lab 1 (e) '치우친 분포를 어떻게 바로잡나'의 답이고, Lab 2 본문이 여러 변환(view)을 만들어 시험해 보라고 권한다. log1p는 0도 변환할 수 있다.")],
    7: [('로그 후 소득 왜도가 음수(-4.9)가 된 이유는?',
          '소득 0원인 1,634행이 로그 후에도 0으로 남아 왼쪽에 꼬리가 생겼기 때문이다.')],
    8: [('결정트리는 왜 변환에 영향을 받지 않나?',
          "트리는 '값이 기준보다 큰가'로만 나눈다. 단위를 바꾸거나 로그를 씌워도 값의 순서가 같아 결과가 같다."),
         ('변환을 교차검증 전에 전체 데이터에 맞춘 건 데이터 누수 아닌가?',
          'Lab 2 방식 그대로 했다. 평균·표준편차만 쓰는 변환이라 영향은 작지만, Lab 5에서 Pipeline으로 학습 fold 안에서만 맞추도록 바로잡는다.')],
    9: [('LOOCV를 15만 건 전체로 하지 않은 이유는?',
          '측정해 보니 1회 학습 0.49초 × 15만 번 ≈ 20시간이고, 조각마다 시험 데이터가 1명이라 조각별 AUC를 계산할 수 없다. 그래서 전체로는 돌리지 않았다. 참고로 1,000건 표본으로는 돌려 봤다(정확도 92.6%, 표준편차 26%).')],
    10: [('6개가 더 높은데 왜 바로 줄이지 않았나?',
          '로지스틱 회귀에서는 6개가 10번 중 9번 높았다(+0.002). 하지만 뺀 두 변수를 다른 모델은 쓸 수 있어 두 후보로 남겼고, Lab 5에서 최종 후보 RF로 비교하니 8개가 10번 중 8번 높아(0.856 vs 0.855) 8개로 확정했다.')],
    11: [('층화는 Lab 3 자료에 없는데 왜 썼나?',
          'Lab 3 k-fold(KFold)의 변형(StratifiedKFold)이다. 일반 KFold는 조각마다 연체 비율이 6.45~7.18%로 달라, 연체자가 적은 이 데이터에서는 비율을 맞춰야 조각끼리 공정하게 비교된다. 흔들림은 ±0.007과 ±0.009로 비슷해 안정성 때문에 고른 것은 아니다.')],
    12: [('SVM을 15만 건 전체로 돌리지 않은 이유는?',
          '학습 시간이 데이터 2배에 약 4배(n²)로 늘었다. 2만 건 8초 → 13.5만 건 약 6분, 10-fold면 약 1시간이다. 그래서 그래프의 다른 5개는 15만 건 결과이고, SVM과 비교할 때는 같은 1만 건 층화 표본에서 6개를 모두 다시 돌렸다(LR 0.823, LDA 0.826, SVM 0.579).')],
    13: [("'연체자 가중치'는 신용점수에서 연체 이력에 주는 감점과 같은 건가?",
          "아니다. 여기서는 학습할 때 연체자를 틀리면 정상 고객보다 약 14배 큰 벌점을 주는 방법(class_weight='balanced')이다. 연체 이력 자체는 '연체 이력' 변수로 이미 입력에 들어 있다."),
         ('정상 고객 9,766명을 거절하는데 괜찮나?',
          "연체로 판정한 사람 중 20%가 진짜 연체자로 무작위(6.7%)의 3배다. 실제 기준값은 '놓친 연체자 1명 비용 vs 거절한 정상 고객 1명 비용'으로 정해야 하고, 기준값은 나중에 바꿀 수 있으니 모델은 AUC로 먼저 고른다.")],
    14: [('LDA가 LR보다 높은 이유는?',
          '둘 다 직선 경계지만 LDA는 정규분포를 가정하고 평균·공분산으로 한 번에 계산한다. 로그+표준화로 분포가 정규에 가까워져 유리했던 것으로 보인다. 다만 차이 0.008은 흔들림(약 0.008~0.009) 수준이라 확실한 우위는 아니다.'),
         ('SVM은 왜 가중치 없이 0.579였나?',
          'AUC는 0/1 판정이 아니라 SVM 점수(decision_function)로 계산했다. 가중치 없이 학습하면 정상이 93%라 연체자를 거의 무시한 경계가 생겨 점수 순위도 연체자를 잘 구별하지 못했고, 가중치를 주면 0.810으로 회복됐다.')],
    15: [('데이터 누수란? 왜 문제인가?',
          '시험 데이터의 정보(평균·표준편차 등)가 학습 단계에 섞이는 것이다. 실제로는 새 고객 데이터를 미리 볼 수 없으므로 점수가 부풀려질 수 있다. Pipeline은 전처리를 각 fold의 학습 데이터로만 맞춘다.')],
    16: [('누수를 막았는데 왜 AUC가 같나?',
          '로그 변환은 값마다 따로 계산해 누수가 없고, 표준화는 평균·표준편차를 쓰는데 15만 건에서는 fold마다 거의 같다. 데이터가 작으면 차이가 날 수 있다.')],
    17: [('왜 RF를 골랐나? AdaBoost는 기본 설정 1등이었는데?',
          '네 후보를 같은 중첩 교차검증으로 튜닝하니 RF 0.856, AdaBoost 0.852(트리 1,000개), KNN 0.848(이웃 500명), LR 0.844였다. 같은 조각끼리 RF가 AdaBoost보다 10번 모두 높았고(+0.004), 학습 시간도 같은 실행에서 약 1/18(3.9초 vs 71초, 13.5만 건)이었다. AdaBoost는 트리를 300 → 1,000개로 늘려도 +0.001이라 더 늘려도 따라잡기 어렵다고 판단했다.'),
         ('튜닝 비교가 공정한가?',
          '바깥 층화 10-fold(지금까지와 같은 조각)로 평가하고, 설정은 각 바깥 학습 구간 안의 3-fold로만 골랐다. 시험 조각은 설정 선택에 쓰지 않았다(notebooks/05d_candidate_tuning.py). 1차에서 AdaBoost·KNN이 범위 끝을 골라 범위를 넓혀 다시 했다.')],
    18: [('배치와 에폭은 무엇인가?',
          '배치는 가중치를 한 번 고칠 때 보는 고객 수, 에폭은 전체 데이터를 한 바퀴 학습하는 것이다. 배치 128이면 한 바퀴에 약 785번 고친다.')],
    19: [('이 표는 10-fold가 아니라 한 번 분할인데 괜찮나?',
          'Lab 6의 Manual Verification(67:33) 방식이다. 배치 간 상대 비교용이고, 최종 성능은 배치 128로 층화 10-fold를 다시 했다(0.850).')],
    20: [('MLP가 랜덤 포레스트보다 낮은 이유는?',
          "변수가 8개뿐이고 관계가 비교적 단순한 표 데이터라 트리 앙상블이 강하다. 차이 0.006은 흔들림(약 0.007) 안이라 '비슷한 수준이지만 넘지 못했다'가 정확하다. 참고 논문(Ichim & Issa 2025)에서도 같은 데이터에서 신경망 0.856, 트리 계열 0.87대였다.")],
    21: [('최적 모델이 Lab 모델보다 확실히 나은가?',
          '같은 10조각에서 10조각 모두 높았고(+0.0020, 차이의 흔들림 0.0012) 우연은 아니다. 다만 은행 손실로는 0.5% 차이라 실무에서 큰 차이는 아니다.'),
         ('최적 모델의 성능은 과장되지 않았나?',
          '중첩 교차검증을 썼다. 설정은 각 조각의 학습 데이터 안에서 3-fold로만 찾고, 바깥 시험 조각은 설정을 고를 때 한 번도 쓰지 않았다.')],
    22: [('PR-AUC는 무엇이고 왜 따로 봤나?',
          '정밀도(연체로 판정한 사람 중 진짜 연체자)와 재현율(전체 연체자 중 잡은 비율)로 그린 곡선의 넓이다. 찍기 수준이 0.5가 아니라 연체율 0.067이다. ROC-AUC는 정상 고객 14만 명 중 몇 %를 잘못 거절했는지로 보므로, 13장처럼 AUC 0.844인데 판정 중 진짜 연체자가 20%인 상황이 가려진다. 같은 10-fold로 재니 Lab 모델 0.344, 최적 모델 0.350(10조각 모두 우세)이었다 (notebooks/09c_pr_auc.py).'),
         ('판정 기준 0.17과 손실 31%는 어떻게 나왔나?',
          '비용을 5 × 놓친 연체자 + 거절한 정상 고객(신용평가 연구의 흔한 5:1 가정)으로 두고, 같은 RF에서 기준을 바꿔 가며 손실이 가장 작은 값을 찾았다. 0.5 대비 손실 31% 감소. 이 RF는 가중치 없이 학습해 확률이 실제와 맞아(부록 B) 이론값 1/(5+1) = 0.167과 거의 같다.')],
}

# 전체 장 시간 배분 (합계 420초). 아직 만들지 않은 장은 대본 파일에 '작성 예정'으로 표시된다.
FULL_PLAN = [("표지", 10), ("목차", 15)] \
    + [(f"Lab {n} · {part}", sec) for n in range(1, 7) for part, sec in (("진행", 10), ("결과", 20), ("발견·결론", 23))] \
    + [("최종 모델 비교", 20), ("실무 점검 · 결론", 45)]
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


# 같은 용어는 발표 전체에서 처음 나온 장에만 풀이한다 (같은 뜻의 다른 이름은 ALIAS로 묶음)
SHOWN_TERMS = set()
TERM_ALIAS = {"k-fold": "10-fold", "층화 10-fold": "층화", "표준편차": "± 값"}

# 제목: 명사형 보고서체 (장 번호 → 제목). lab_header 에 넘긴 제목보다 우선한다.
TITLES = {
    3: "Lab 1 절차에 따른 15만 건 데이터 로딩과 기술통계",
    4: "연체 비율 6.7%와 이상값 확인",
    5: "극단값에 가려진 '사용 비율'과 연체의 관계",
    6: "Lab 2 절차에 따른 시각화와 네 가지 변환 적용",
    7: "세 변수의 쏠림과 로그 변환 효과",
    8: "로그+표준화(회귀)와 Binarizer(트리)의 효과 차이",
    9: "Lab 3의 특성 선택 4종과 검증 방법 4종 적용",
    10: "방법별 선택 변수의 차이와 변수 축소 효과",
    11: "한 번 분할의 AUC 변동(0.816~0.828)과 층화 10-fold 채택",
    12: "Lab 4의 평가지표 6종과 알고리즘 6종 비교",
    13: "정확도 93% 대비 낮은 연체자 탐지율(1.3%)",
    14: "15만 건 비교에서 LDA의 최고 AUC(0.832)",
    15: "Lab 5의 파이프라인·앙상블·튜닝·저장 적용",
    16: "표준화 누수 점검과 앙상블을 통한 트리 성능 향상",
    17: "후보 4개 튜닝 비교와 RF 선택(0.856)",
    18: "Lab 6 MLP의 15만 건 적용 과정",
    19: "배치 크기와 에폭에 따른 AUC·학습 시간 비교",
    20: "MLP(0.850)와 튜닝된 RF(0.856)의 성능 비교",
    21: "Lab 기반 모델과 최적 모델의 비교",
    22: "실무 적용 가능성: ROC-AUC의 한계와 PR-AUC 점검",
}


def term_line(s, terms, y=1.62):
    """제목 아래 용어 풀이 한 줄. terms: [(용어, 쉬운 풀이)]. 앞 장에서 풀이한 용어는 뺀다."""
    terms = [(t, d) for t, d in terms if TERM_ALIAS.get(t, t) not in SHOWN_TERMS]
    SHOWN_TERMS.update(TERM_ALIAS.get(t, t) for t, _ in terms)
    if not terms:
        return
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
    text(s, 0.8, 0.98, 11.8, 0.9, TITLES.get(len(prs.slides), title), size=30, weight="extrabold", spacing=-1, name="!!title")


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


# 발견·결론 장마다 "무엇을 바꿨더니 결과가 어떻게 달라졌나" 한 줄 (장 번호 → (바꾼 것, 결과))
CHANGES = {
    5: ("상관 계산: 값 → 순위", "사용 비율 -0.002 → 0.24"),
    8: ("로그+표준화 (로지스틱 회귀)", "AUC 0.787 → 0.824"),
    11: ("검증: 층화 10-fold", "0.816~0.828 → 0.824"),
    14: ("알고리즘 (트리 → LDA)", "AUC 0.595 → 0.832"),
    17: ("RF 잎 최소 인원 1 → 176명", "AUC 0.824 → 0.856"),
    20: ("RF → 신경망(MLP)", "AUC 0.856 → 0.850"),
}


def conclusion_card(s, x, y, w, h, lines, anim=None):
    """이번 Lab 결과가 '연체 예측'이라는 목표에 주는 의미."""
    change = CHANGES.get(len(prs.slides))
    if change:  # 바꾼 것 → 결과 줄을 넣을 자리만큼 카드를 위아래로 늘린다
        y, h = y - 0.1, h + 0.2
    parts = [rect(s, x, y, w, h, C["violet_pale"], radius=0.22),
             rect(s, x + 0.3, y + 0.32, 0.06, 0.36, C["violet"]),
             text(s, x + 0.5, y + 0.28, w - 0.8, 0.45,
                  [[("목표에 대한 결론", {}), ("   연체 예측 모델에 주는 의미", {"size": 12, "color": C["mute"], "weight": "medium"})]],
                  size=16, weight="bold", color=C["violet"]),
             text(s, x + 0.3, y + (1.22 if change else 0.85), w - 0.6, h - (1.35 if change else 1.0), lines,
                  size=14, color=C["ink"], weight="medium", line=1.45)]
    if change:
        what, res = change
        parts += [rect(s, x + 0.3, y + 0.74, w - 0.6, 0.38, C["card"], radius=0.19),
                  text(s, x + 0.3, y + 0.74, w - 0.6, 0.38,
                       [[("바꾼 것  ", {"color": C["violet"], "weight": "bold", "size": 11}),
                         (what, {"color": C["ink"], "weight": "semibold"}),
                         ("  →  ", {"color": C["mute"]}),
                         (res, {"color": C["blue"], "weight": "extrabold"})]],
                       size=12.5, align="c", anchor="m")]
    group(s, parts, f"anim{anim}_conclusion" if anim else "conclusion")


def pill(s, x, y, w, label, fill, color, size=13, weight="semibold", h=0.42):
    return [rect(s, x, y, w, h, fill, radius=h / 2),
            text(s, x, y, w, h, label, size=size, color=color, weight=weight, align="c", anchor="m")]


# =====================================================================
# 1. 표지
# =====================================================================
s = new_slide()
text(s, 0.9, 1.75, 6.5, 0.45, "알고리즘분석 과제 발표", size=17, color=C["blue"], weight="semibold", name="!!kicker")
text(s, 0.9, 2.3, 6.8, 2.2, ["Lab 1~6 기반", "신용위험 예측 모델링"], size=48, weight="extrabold",
     line=1.12, spacing=-1.5, name="!!title")
text(s, 0.9, 4.45, 6.5, 0.45, "Give Me Some Credit · 150,000명의 재무 데이터", size=17, color=C["sub"], weight="medium",
     name="anim1_sub")
group(s, pill(s, 0.9, 5.15, 1.45, "Lab 1 → 6", C["blue_pale"], C["blue"])
      + pill(s, 2.48, 5.15, 1.75, "150,000명", C["red_pale"], C["red"])
      + pill(s, 4.36, 5.15, 1.25, "MLP", C["violet_pale"], C["violet"]), "anim2_pills")
rect(s, 0.9, 6.35, 6.0, 0.012, C["line"])
text(s, 0.9, 6.47, 6.5, 0.75,
     [[("강희준", {"size": 19, "weight": "extrabold", "color": C["ink"]})],
      [("수원대학교 컴퓨터SW학과", {"size": 12.5, "weight": "medium", "color": C["mute"]})]],
     size=15, line=1.25, name="presenter")
code_window(s, 7.75, 1.9, 4.85, 3.35, "story.py", '''
# 발표의 흐름
for lab in range(1, 7):
    finding  = explore(data, lab)
    question = next_question(finding)
    print(f"Lab {lab}: {question}")

model = MLP(hidden=[12, 8])  # Lab 6
model.fit(X, y, batch_size=128)
''', size=12, highlight=(1, 2, 3), anim=3, lh=19)
notes(s, "안녕하세요. 저는 Lab 1부터 Lab 6까지 실습한 방법을 그대로 따라가면서, 신용위험 예측 문제를 단계별로 풀어 본 과정을 발표하겠습니다.")

# =====================================================================
# 목차 (문제 정의 포함: 데이터·목표 한 줄 + 항목마다 AI Modeling 단계)
# =====================================================================
s = new_slide()
text(s, 0.8, 0.55, 5.4, 0.35, "목차", size=14, color=C["blue"], weight="semibold", name="!!kicker")
text(s, 0.8, 0.98, 11.8, 0.9, "발표 순서", size=34, weight="extrabold", spacing=-1, name="!!title")
text(s, 0.8, 1.66, 11.8, 0.32,
     [[("데이터 · 목표  ", {"color": C["blue"], "weight": "bold"}),
       ("과제 기준 그대로 · Give Me Some Credit 15만 명으로 2년 안에 연체할 확률 예측 (Lecture 6의 6단계를 Lab 1~6으로)", {"color": C["sub"]})]],
     size=12.5, weight="medium")
TOC = [("01", "Lab 1  데이터 로딩 · 기술통계", "데이터는 어떻게 생겼나", ["2 데이터 분석"], C["blue"]),
       ("02", "Lab 2  시각화 · 데이터 준비", "값을 어떻게 바꿀까", ["2 분석", "3 준비"], C["blue"]),
       ("03", "Lab 3  특성 선택 · 리샘플링", "변수는 다 필요할까, 평가는 믿을 만할까", ["3 준비", "4 평가"], C["blue"]),
       ("04", "Lab 4  평가지표 · 알고리즘 비교", "어떤 알고리즘이 연체자를 잘 골라낼까", ["4 알고리즘 평가"], C["blue"]),
       ("05", "Lab 5  파이프라인 · 앙상블 · 튜닝", "어떤 모델을 최종으로 고를까", ["5 성능 개선"], C["blue"]),
       ("06", "Lab 6  Keras MLP", "신경망은 이 기준을 넘을까", ["4 평가", "5 개선"], C["blue"]),
       ("07", "최종 모델 비교", "Lab 모델 vs Lab 밖 방법을 더한 최적 모델", ["6 결과 제시"], C["violet"]),
       ("08", "실무에서 쓸 수 있을까", "ROC-AUC의 한계와 PR-AUC 점검, 느낀 점", ["6 결과 제시"], C["violet"]),
       ("부록", "A 15만 건 처리 · B 예측 확률 · C 보충 실험", "발표 시간 밖 · 질문 대비", [], C["mute"])]
for k, (no, name, q, stages, col) in enumerate(TOC):
    x, y = (0.8 if k < 5 else 6.75), 2.15 + (k % 5) * 0.93
    parts = [rect(s, x, y, 5.78, 0.82, C["card"], radius=0.18),
             text(s, x + 0.25, y, 0.85, 0.82, no, size=20 if no != "부록" else 15, color=col, weight="extrabold", anchor="m"),
             text(s, x + 1.1, y + 0.11, 4.6, 0.35, name, size=15, weight="bold"),
             text(s, x + 1.1, y + 0.45, 3.2, 0.3, q, size=11.5, color=C["sub"], weight="medium")]
    cx = x + 5.62
    for st in reversed(stages):
        w = 0.17 + 0.135 * len(st)
        cx -= w
        parts += pill(s, cx, y + 0.45, w, st, C["blue_pale"] if col != C["violet"] else C["violet_pale"],
                      C["blue"] if col != C["violet"] else C["violet"], size=9.5, h=0.28)
        cx -= 0.06
    group(s, parts, f"anim{1 if k < 5 else 2}_toc{k}")
text(s, 6.75, 6.85, 5.8, 0.3,
     [[("Lab마다  ", {"color": C["blue"], "weight": "bold"}),
       ("진행 → 결과 → 발견 · 발견이 다음 Lab의 질문", {"color": C["sub"]})]],
     size=12, weight="medium")
notes(s, "목차입니다. 데이터와 목표는 과제에서 정한 그대로이고, Lecture 6의 AI 모델링 6단계를 Lab 1부터 6까지로 밟아 갑니다. 항목 오른쪽 칩이 그 Lab이 해당하는 단계입니다. Lab마다 진행, 결과, 발견 세 장이고, 각 Lab의 발견이 다음 Lab의 질문으로 이어집니다. 마지막으로 최종 모델을 비교하고, 그 모델을 실무에서 쓸 수 있는지 확인합니다. 부록 세 장은 발표 시간에는 넘기고 질문이 나오면 씁니다.")

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
    [("② ", {"color": C["violet"], "weight": "bold"}), ("극단값 영향일 수 있어 → Lab 2에서 변환 효과 확인", {})],
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
오른쪽은 이 세 변수에 log1p, 즉 log(1+x)를 적용한 전후입니다. 부채 비율의 왜도는 95.2에서 1.7로 크게 줄었고, 사용 비율도 97.6에서 11.7로 줄었습니다.
로그를 씌운 부채 비율 그림 오른쪽의 두 번째 봉우리는, Lab 1에서 찾은 소득을 채운 고객들의 부채 금액이 그림으로 드러난 것입니다.
30–59일 연체 횟수도 22.6에서 4.3으로 줄었지만, 96, 98이라는 특수 코드 때문에 여전히 치우쳐 있습니다. 이 값은 따로 처리해야 합니다.""")

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
Normalizer는 0.522로 거의 무작위 수준이 됐습니다. 행마다 크기를 1로 맞추면서 단위가 다른 변수들의 절대 크기 정보가 사라졌기 때문으로 보입니다.
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
          ("한 번 분할", "train_test_split", "AUC 0.823"), ("층화 k-fold", "KFold→Stratified", "연체 비율 → 6.68%"),
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
k-fold는 Lab 자료의 KFold로 돌려 보니 조각마다 연체 비율이 6.45%에서 7.18%로 달라서, 비율을 6.68%로 맞추는 층화 방식(StratifiedKFold)으로 바꿨습니다.
LOOCV는 15만 번 학습해야 해서 한 번 학습 시간을 재어 보니 약 20시간이 걸려, 전체로는 돌리지 않았습니다.""")

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
          text(s, 1.1, 2.32, 5.2, 0.35, [[("방법별로 고른 변수  ", {"weight": "bold"}), ("● = 선택됨 · PCA는 압축이라 오른쪽 그래프에", {"color": C["mute"]})]],
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
                                         ("트리 중요도 순서로 추가 · 순서는 전체 데이터로 정한 탐색값", {"color": C["mute"]})]], size=12, weight="medium"),
          s.shapes.add_picture(str(FIG / "lab3_subset_auc.png"), Inches(7.0), Inches(2.75), width=Inches(5.35))],
      "anim2_subset")
group(s, [rect(s, 6.8, 5.85, 5.73, 1.2, C["blue_pale"], radius=0.2),
          text(s, 7.1, 5.85, 5.2, 1.2,
               [[("회귀: 6개 0.826 > 8개 0.824 (10번 중 9번)", {"weight": "extrabold", "color": C["ink"], "size": 16})],
                [("다른 모델은 다를 수 있어 → ", {"color": C["sub"]}), ("6개·8개 두 후보로", {"color": C["blue"], "weight": "bold"})]],
               size=13, weight="medium", anchor="m", line=1.35)], "anim3_keep")
notes(s, """왼쪽은 세 방법이 고른 변수입니다. 통계 검정은 로그 변환 후 상위 4개, RFE는 상위 3개, 트리 중요도는 상위 4개를 표시했습니다.
세 방법 모두 고른 변수는 사용 비율과 나이, 두 방법이 고른 변수는 연체 이력입니다. 트리 중요도는 값의 종류가 많아 나눌 자리가 많은 부채 비율과 소득을 높게 쳤습니다.
오른쪽은 트리 중요도 순서로 변수를 하나씩 더하며 층화 10-fold AUC를 잰 결과입니다. 상위 3–4개만 쓰면 0.75–0.78로 떨어지고, 5개부터 0.82대를 회복합니다.
상위 6개가 0.826으로 8개의 0.824보다 0.002 높았고, 같은 조각끼리 비교해도 10번 중 9번 6개가 높았습니다. 다만 빠진 두 변수(부양가족, 대출 보유)를 회귀가 아닌 다른 모델은 쓸 수도 있어서, 여기서 정하지 않고 6개와 8개를 모두 후보로 남겼습니다. 답은 모델을 정한 Lab 5에서 냅니다.
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
         ("02", "층화 10-fold: 0.824 ± 0.009", "KFold + 연체 비율 맞추기 (일반 6.45~7.18%)", C["blue"]),
         ("03", "LOOCV: 약 20시간 · 조각별 AUC 불가", "조각마다 시험이 1명이라 순위를 매길 수 없음", C["blue"])]
parts = []
for i, (no, head, body, col) in enumerate(finds):
    yy = 2.15 + i * 0.72
    parts += [text(s, 7.1, yy, 0.6, 0.4, no, size=17, color=col, weight="extrabold"),
              text(s, 7.65, yy, 4.9, 0.4, head, size=16, weight="bold"),
              text(s, 7.65, yy + 0.34, 4.9, 0.35, body, size=12.5, color=C["sub"], weight="medium")]
group(s, parts, "anim2_finds")
conclusion_card(s, 7.05, 4.4, 5.48, 1.8, [
    [("① ", {"color": C["violet"], "weight": "bold"}), ("이후 모든 비교는 층화 10-fold AUC로 통일", {})],
    [("② ", {"color": C["violet"], "weight": "bold"}), ("변수는 6개·8개 두 후보로 (모델 정한 뒤 비교)", {})],
], anim=3)
next_bar(s, "어떤 알고리즘이 연체자를 가장 잘 골라낼까?", "Lab 4", "AUC", "0.824 ± 0.009", anim=4)
notes(s, """검증 방법별로 나온 AUC를 점으로 찍었습니다. 점 하나가 시험 한 번, 가로선이 평균입니다.
데이터를 한 번만 나누면, 나누는 방식만 바꿔도 AUC가 0.816에서 0.828까지 0.012나 흔들렸습니다. 이 정도면 모델 간 차이와 구별하기 어렵습니다.
10번 나눠 시험하는 10-fold는 평균 0.824였습니다. Lab 자료의 일반 KFold는 시험마다 연체 비율이 6.45%에서 7.18%로 달라서, 비율을 6.68%로 맞추는 층화 10-fold를 더했습니다. 흔들림은 0.007과 0.009로 비슷해, 안정성보다는 모든 시험을 같은 조건으로 만들려고 고른 것입니다.
LOOCV는 약 20시간이 걸리고, 조각마다 시험 데이터가 한 명뿐이라 조각별 AUC를 계산할 수 없었습니다. 예측을 모두 모아 AUC를 한 번에 낼 수는 있지만, 다른 방법과 같은 기준(조각별 평균)으로 비교할 수 없어 제외했습니다.
목표 관점의 결론은, 이후 모든 알고리즘 비교를 층화 10-fold AUC로 통일하고, 변수는 6개와 8개 두 후보를 남긴다는 것입니다.
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
                  ("학습 가중치", "연체자를 틀리면 정상 14명을 틀린 만큼 벌점(class_weight)")])


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
                                         ("층화 10-fold · 상자 = 10번 시험의 가운데 50%", {"color": C["mute"]})]], size=13),
          s.shapes.add_picture(str(FIG / "lab4_compare.png"), Inches(1.0), Inches(2.78), width=Inches(5.6)),
          text(s, 1.1, 5.72, 5.5, 0.3, "* 음영 = SVM은 15만 건에 약 1시간이라 1만 건 표본으로 따로 비교 · 정확도는 모두 88~93%",
               size=10, color=C["mute"], weight="medium")],
      "anim1_chart")
finds = [("01", "LDA 0.832 > LR 0.824 ≈ NB 0.822", "직선에 가까운 관계 → 선형 모델이 강함", C["blue"]),
         ("02", "KNN 0.727 · 결정트리 0.595", "KNN은 확률이 6단계뿐, 트리는 과적합", C["red"]),
         ("03", "SVM 0.579 → 가중치 0.810", "별도 1만 건 표본 · 같은 표본 LR 0.823, LDA 0.826", C["red"])]
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
SVM은 15만 건으로는 10-fold에 약 한 시간이 걸려 1만 건 표본으로 비교했는데, 0.579로 거의 찍기 수준이었습니다. AUC는 0/1 판정이 아니라 SVM 점수로 계산했는데, 가중치 없이 학습하면 연체자를 거의 무시한 경계가 생겨 점수 순위도 연체자를 잘 구별하지 못했습니다. 연체자 가중치를 주면 0.810으로 회복됐습니다.
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
                [[("LR·LDA: 표준화를 fold 안으로 옮겨도 같음", {"weight": "bold", "color": C["blue"]})],
                 [("15만 건이라 fold마다 ", {"color": C["sub"]}),
                  (f"평균 차이가 표준편차의 {L5['fold_diff']:.1%} 이하", {"color": C["sub"]})]],
                size=12.5, weight="medium", anchor="m", line=1.3),
           rect(s, 1.05, 5.97, 5.1, 0.88, C["bg"], radius=0.16),
           text(s, 1.3, 5.97, 4.7, 0.88,
                [[("데이터 정리 효과 ≈ 0", {"weight": "bold", "color": C["ink"]}), ("  96/98 코드는 연체율 54.6%지만", {"color": C["sub"]})],
                 [("'연체 이력' 변수에 이미 담겨 있었다 · 96/98·소득 채운 값 정리 전후 차이 미미", {"color": C["sub"]})]],
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
# 후보별 튜닝 비교 (notebooks/05d_candidate_tuning.py, 중첩 교차검증): 2차 탐색(+)이 있으면 그 값을 쓴다
_sc = pd.read_csv(ROOT / "notebooks" / "05d_candidate_scores.csv")
_pa = pd.read_csv(ROOT / "notebooks" / "05d_candidate_params.csv")
def _cand(name):
    key = name + "+" if (_sc.model == name + "+").sum() == 10 else name  # 2차 탐색이 10조각 다 끝났을 때만
    g, p = _sc[_sc.model == key].sort_values("fold"), _pa[_pa.model == key]
    base = _sc[_sc.model == name].sort_values("fold").base_auc.mean()
    return dict(base=base, tuned=g.tuned_auc.mean(), folds=g.tuned_auc.values, params=p)
CAND = {n: _cand(n) for n in ["RF", "AdaBoost", "KNN", "LR"]}
_mode = lambda p, col: p[col].mode()[0]
CAND["RF"]["pick"] = f"잎 최소 {L5['final_leaf']}명*"
CAND["AdaBoost"]["pick"] = (f"트리 {int(_mode(CAND['AdaBoost']['params'], 'n_estimators'))} · "
                            f"학습률 {_mode(CAND['AdaBoost']['params'], 'learning_rate'):g}")
CAND["KNN"]["pick"] = f"이웃 {int(_mode(CAND['KNN']['params'], 'n_neighbors'))}명"
CAND["LR"]["pick"] = f"C {_mode(CAND['LR']['params'], 'C'):g}"
_rank = sorted(CAND, key=lambda n: -CAND[n]["tuned"])
_fit = lambda n: _sc[_sc.model == (n + "+" if (_sc.model == n + "+").sum() == 10 else n)].refit_sec.mean()
_best, _second = _rank[0], _rank[1]
_diff = CAND[_best]["folds"] - CAND[_second]["folds"]
lab_header(s, 5, 3, "파이프라인 · 앙상블 · 튜닝", "후보별 튜닝과 모델 선택",
           terms=[("중첩 교차검증", "설정은 학습 데이터 안에서만 찾고 바깥 조각으로 평가")])
rows17 = [("후보", "탐색한 설정 (왜)", "고른 값", "기본", "튜닝")]
why = {"RF": ("잎 최소 인원", "잎 1명 = 외우기"), "AdaBoost": ("트리 수 · 학습률", "기본 설정 1등"),
       "KNN": ("이웃 수", "5명이면 확률 6단계"), "LR": ("C (규제 세기)", "선형 기준선")}
for n in _rank:
    rows17.append((n, why[n], CAND[n]["pick"], f"{CAND[n]['base']:.3f}", f"{CAND[n]['tuned']:.3f}"))
cw17 = [1.0, 1.75, 1.45, 0.65, 0.65]
lp = [rect(s, 0.8, 2.15, 6.0, 4.05, C["card"], radius=0.25),
      text(s, 1.1, 2.33, 5.4, 0.35, [[("후보 4개 튜닝  ", {"weight": "bold", "color": C["ink"]}),
                                     ("같은 층화 10-fold · 설정은 안쪽 3-fold로 선택", {"color": C["mute"]})]], size=12.5)]
tb = s.shapes.add_table(len(rows17), 5, Inches(1.0), Inches(2.8), Inches(sum(cw17)), Inches(0.62 * len(rows17))).table
for j, w in enumerate(cw17):
    tb.columns[j].width = Inches(w)
for i, row in enumerate(rows17):
    tb.rows[i].height = Inches(0.48 if i == 0 else 0.62)
    win = i > 0 and row[0] == _best
    for j, val in enumerate(row):
        cell = tb.cell(i, j)
        cell.margin_left = cell.margin_right = Inches(0.05)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.fill.solid()
        cell.fill.fore_color.rgb = rgb(C["blue_pale"] if i == 0 else ("#EEF0FF" if win else C["bg"]))
        tf = cell.text_frame
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT if j in (0, 1) else PP_ALIGN.CENTER
        if i == 0:
            set_font(p.add_run(), 10.5, "bold", C["blue"]); p.runs[0].text = val
            continue
        if j == 1:
            r = p.add_run(); r.text = val[0]; set_font(r, 11.5, "semibold", C["ink"])
            p2 = tf.add_paragraph(); r2 = p2.add_run(); r2.text = val[1]; set_font(r2, 9.5, "medium", C["mute"])
            continue
        r = p.add_run(); r.text = val
        if j == 4:
            set_font(r, 13, "extrabold", C["violet"] if win else C["ink"])
        elif j == 3:
            set_font(r, 12, "medium", C["mute"])
        else:
            set_font(r, 12.5 if j == 0 else 11, "bold" if j == 0 else "medium", C["violet"] if (win and j == 0) else C["ink"])
lp.append(s.shapes[-1])
lp.append(text(s, 1.0, 5.8, 5.7, 0.4, [f"* 중첩 탐색은 주로 100명을 골랐지만 50명 이상이면 AUC가 거의 같아 Lab 5의 {L5['final_leaf']}명 유지",
                                        "  기본값 0.826은 정리 단계 포함 (16장 원본 RF는 0.824)"],
                size=9.5, color=C["mute"], weight="medium", line=1.15))
group(s, lp, "anim1_table")
_jump = max(CAND, key=lambda n: CAND[n]["tuned"] - CAND[n]["base"])
finds = [("01", f"{_best} 선택: 튜닝 후 최고 {CAND[_best]['tuned']:.3f}",
          f"{_second}보다 10번 중 {(_diff > 0).sum()}번 높음 ({_diff.mean():+.3f}) · 학습 시간 약 1/{_fit(_second) / _fit(_best):.0f}", C["blue"]),
         ("02", f"튜닝 효과 최대: {_jump} {CAND[_jump]['base']:.3f} → {CAND[_jump]['tuned']:.3f}",
          f"설정 하나({why[_jump][0]})로 크게 바뀜 · LR은 그대로", C["blue"]),
         ("03", "Lab 3 후보 비교: 8개 0.856 > 6개 0.855", f"10번 중 8번 8개 우세 → 8개 확정 · 저장 후에도 {L5['test_auc']:.3f}", C["blue"])]
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
notes(s, f"""Lab 4와 5에서 성능이 좋았거나 튜닝 여지가 분명한 후보 네 개를 같은 조건으로 튜닝해 비교했습니다. 로지스틱 회귀는 규제 세기 C, KNN은 이웃 수, AdaBoost는 트리 수와 학습률, 랜덤 포레스트는 잎 최소 인원입니다.
비교는 중첩 교차검증으로 했습니다. 바깥은 지금까지와 같은 층화 10-fold이고, 각 바깥 학습 구간 안에서 3-fold로 설정을 고른 뒤, 고른 설정을 바깥 시험 조각으로 평가했습니다. 설정을 고를 때 시험 조각을 보지 않았으므로 후보끼리 공정하게 비교됩니다.
결과는 랜덤 포레스트가 {CAND['RF']['tuned']:.3f}으로 가장 높았고, 기본 설정 1등이던 AdaBoost는 트리를 1,000개까지 늘려도 {CAND['AdaBoost']['tuned']:.3f}이었습니다. 같은 조각끼리 보면 랜덤 포레스트가 열 번 모두 높았고, 학습도 훨씬 빨랐습니다.
KNN은 이웃 5명이면 확률이 여섯 단계뿐이라 순위를 잘게 못 매겼는데, 이웃을 늘리자 {CAND['KNN']['base']:.3f}에서 {CAND['KNN']['tuned']:.3f}로 가장 크게 올랐습니다. 로지스틱 회귀는 C를 바꿔도 그대로였습니다.
랜덤 포레스트의 잎 최소 인원은 Lab 5 방식대로 Grid(1, 10, 50, 100, 200)와 Random(1~300, 10번)으로도 찾았고, 각각 {L5['grid_leaf']}명과 {L5['rand_leaf']}명에서 {L5['final_auc']:.3f}이었습니다. 50명 이상이면 거의 같아서, 최종 모델은 {L5['final_leaf']}명으로 두었습니다. 저장했다 다시 불러와도 시험 AUC {L5['test_auc']:.3f}로 같았습니다.
Lab 3에서 남긴 변수 후보 두 개도 이 모델로 비교했습니다. 같은 조각끼리 보면 8개가 10번 중 8번 높아(0.856 vs 0.855) 8개로 확정했습니다.
목표 관점의 결론은, Lab 범위의 최고 기준이 {L5['final_auc']:.3f}이고, 다음 Lab 6의 신경망은 이 기준을 넘어야 의미가 있다는 것입니다.""")

# =====================================================================
# Lab 6 - (1) 진행 : MLP 7단계
# =====================================================================
# 노트북 06 실행 결과 (results.md 와 같은 값). None 이면 Lab 5까지만 만든다.
L6 = dict(
    train_auc=0.8523, test_auc=0.8513, cv_auc=0.85047, cv_std=0.00752, fold_sec=67, mean_stop=43,
    wrap_auc=0.8513, wrap_std=0.0066, grid_best=0.8501, grid_worst=0.8490, grid_best_opt="adam", grid_best_init="glorot",
    n_params=221, final_sec=50, final_epochs=28,
    title="MLP는 0.850으로, 튜닝한 랜덤 포레스트(0.856)를 넘지 못했다",
    vs_rf_head="RF 0.856보다 0.006 낮음",
    vs_rf_body="15만 건 학습도 RF 5초 vs MLP 50초 (약 10배)",
    concl1="Lab 범위 최종: 정리 + RF(잎 176명), AUC 0.856",
    concl2="MLP 이점 없음: AUC −0.006 · 학습 시간 10배",
    tracker="0.856",
    notes="""MLP를 층화 10-fold로 평가한 AUC는 0.850, 흔들림은 0.008이었습니다. EarlyStopping 덕분에 fold당 평균 43 에폭에서 멈췄고, fold 하나에 평균 67초가 걸렸습니다.
Lab 6 Grid Search대로 optimizer 두 가지와 초기화 세 가지를 바꿔 봤지만 0.849에서 0.850 사이로, 차이가 흔들림보다 훨씬 작았습니다. Lab 기본값인 adam과 glorot이 가장 높았습니다.
그래서 신경망은 튜닝한 랜덤 포레스트 0.856보다 0.006 낮았습니다. 흔들림 범위 안이라 거의 같은 수준이지만, 15만 건 학습에 랜덤 포레스트는 5초, MLP는 50초가 걸렸습니다.
변수가 8개뿐이고 관계가 비교적 단순한 표 형태 데이터에서는 트리 앙상블이 신경망보다 효율적이라는 뜻입니다. 참고한 논문에서도 같은 데이터에서 신경망이 트리 계열과 비슷하거나 조금 낮았습니다.
목표 관점의 결론은, Lab 범위의 최종 모델은 데이터 정리와 잎 최소 176명 랜덤 포레스트로 정하고, MLP는 비교 결과로 남긴다는 것입니다.""")
if L6 is None:
    prs.save(DRAFT)
    write_script_md(ROOT / "slides" / "script.md")
    print("saved (Lab 6 결과 대기 중)", DRAFT)
    raise SystemExit
BE = pd.read_csv(ROOT / "notebooks" / "06a_batch_epoch_results.csv")   # 실험 A: 배치 × 에폭
be = lambda b, e, col: float(BE[(BE.batch == b) & (BE.epoch == e)][col].iloc[0])
s = new_slide()
lab_header(s, 6, 1, "Keras MLP", "Lab 6의 MLP 단계를 15만 건에 맞춰 적용했다",
           terms=[("배치", "가중치를 한 번 고칠 때 보는 고객 수"), ("에폭", "전체 데이터를 한 바퀴 학습"),
                  ("EarlyStopping", "검증 점수가 멈추면 학습 중단")])
steps6 = [("첫 MLP", "12-8-1 · adam", "Lab 구조 그대로"), ("학습 평가", "evaluate(X, Y)", f"{L6['train_auc']:.3f} vs {L6['test_auc']:.3f}"),
          ("검증 분리", "validation 33%", "배치 × 에폭 표"), ("배치 선택", "batch 10~2048", "128 선택"),
          ("k-fold 반복", "StratifiedKFold", f"AUC {L6['cv_auc']:.3f}"), ("Keras 래퍼", "cross_val_score", f"AUC {L6['wrap_auc']:.3f}"),
          ("Grid Search", "optimizer × init", f"{L6['grid_best']:.3f}"), ("가중치 저장", "weights.out", f"{L6['n_params']}개")]
parts = [text(s, 0.8, 1.98, 11.7, 0.3, [[("Lab 6 자료의 진행 순서", {"weight": "bold"}),
                                         ("   → 1–3 첫 MLP·검증 · 4 15만 건 학습 설정 · 5–8 교차검증·튜닝·저장", {"color": C["mute"], "weight": "medium"})]],
              size=13)]
for i, (ko, fn, res) in enumerate(steps6):
    x = 0.8 + i * (sw3 + gap3)
    key_step = i == 3
    fill = C["violet_pale"] if key_step else C["blue_pale"]
    dot = C["violet"] if key_step else C["blue"]
    parts += [rect(s, x, 2.35, sw3, 1.18, fill, radius=0.16),
              oval(s, x + 0.13, 2.47, 0.28, dot),
              text(s, x + 0.13, 2.47, 0.28, 0.28, str(i + 1), size=10.5, color="#FFFFFF", weight="bold", align="c", anchor="m"),
              text(s, x + 0.47, 2.47, sw3 - 0.5, 0.28, ko, size=11.5 if len(ko) < 7 else (10.5 if len(ko) < 10 else 9.5), weight="bold", anchor="m"),
              text(s, x + 0.13, 2.84, sw3 - 0.2, 0.25, fn, size=9.5, color=dot, weight="medium", latin=MONO),
              rect(s, x + 0.1, 3.14, sw3 - 0.2, 0.3, C["bg"], radius=0.08),
              text(s, x + 0.1, 3.14, sw3 - 0.2, 0.3, res, size=10, color=C["violet"] if key_step else C["ink"],
                   weight="semibold", align="c", anchor="m")]
group(s, parts, "anim1_steps")
h6 = capture_card(s, 0.8, 3.7, 5.2, CAP / "lab6_kfold.png", "p.6 · 2 k-fold 반복 (batch 10, 150 에폭)", anim=2, img_w=4.6)
capture_card(s, 0.8, 3.7 + h6 + 0.12, 5.2, CAP / "lab6_grid.png", "p.9 · 3 Grid Search (54가지 × 3-fold)", anim=2, img_w=4.6)
console(s, 6.2, 3.7, 6.33, 3.35, "실행 결과 · 06_lab6_mlp.ipynb", [
    cmd("fit(batch_size=10, epochs=150)", "3 Lab 그대로"),
    [(f"모델 1개 {be(10, 150, 'train_sec') / 60:.0f}분 · Grid Search 이틀 이상", "red")],
    cmd("fit(batch_size=128, EarlyStopping)", "4 선택"),
    [(f"{L6['fold_sec']:.0f}초 · 평균 {L6['mean_stop']:.0f} 에폭에서 멈춤", "hi")],
    cmd("for train, test in kfold.split(X, Y)", "5 직접 반복"),
    [(f"AUC {L6['cv_auc']:.4f} (± {L6['cv_std']:.4f})", "hi")],
    cmd("cross_val_score(KerasClassifier)", "6 래퍼"),
    [(f"AUC {L6['wrap_auc']:.4f} (± {L6['wrap_std']:.4f})", "hi")],
    cmd("GridSearchCV(optimizer × init)", "7 격자 6가지"),
    [(f"{L6['grid_best_opt']} · {L6['grid_best_init']} → {L6['grid_best']:.4f}", "hi")],
], size=10.5, anim=3, lh=15)
notes(s, f"""Lab 6은 Keras로 MLP를 만들고, 검증 데이터를 나누고, k-fold를 직접 반복한 뒤, KerasClassifier 래퍼와 Grid Search로 넘어갑니다.
구조는 Lab 6 그대로 은닉층 12개, 8개에 relu, 출력 sigmoid, adam, binary crossentropy를 썼고, 입력은 Lab 2의 로그와 표준화, Lab 4의 연체자 가중치를 적용했습니다.
문제는 학습 설정이었습니다. Lab 6 예제는 배치 10, 150 에폭인데, 15만 건에서는 모델 하나에 {be(10, 150, 'train_sec') / 60:.0f}분이 걸려 10-fold면 4시간이 넘고, Lab의 격자 54가지를 3-fold로 돌리면 이틀이 넘습니다.
그래서 먼저 배치 크기 다섯 가지를 150 에폭까지 직접 재서 배치 128을 골랐고, 에폭은 검증 점수가 10 에폭 동안 오르지 않으면 멈추는 EarlyStopping으로 정했습니다.
그 설정으로 층화 10-fold를 직접 반복했고, 같은 일을 KerasClassifier와 cross_val_score로도 해서 결과가 같음을 확인했습니다. Grid Search는 배치와 에폭을 빼고 optimizer와 초기화 6가지만 찾았습니다.""")

# =====================================================================
# Lab 6 - (2) 결과 : 배치 × 에폭 표
# =====================================================================
s = new_slide()
speed = be(10, 150, 'train_sec') / be(128, 150, 'train_sec')
lab_header(s, 6, 2, "Keras MLP", f"배치를 10에서 128로 키우자 시간은 {speed:.0f}배 줄고 AUC는 그대로였다",
           terms=[("검증 AUC", "학습에 안 쓴 33%로 잰 점수"), ("시간", "그 에폭까지 순수 학습 시간")])
E6, B6 = [1, 10, 50, 150], [10, 32, 128, 512, 2048]
tparts = [rect(s, 0.8, 2.15, 6.55, 4.9, C["card"], radius=0.25),
          text(s, 1.1, 2.32, 6.0, 0.35, [[("배치 크기 × 에폭  ", {"weight": "bold"}),
                                         ("칸 = 검증 AUC / 누적 학습 시간 · 15만 건 중 10만 건 학습", {"color": C["mute"]})]], size=12, weight="medium")]
tbl = s.shapes.add_table(len(B6) + 1, len(E6) + 1, Inches(1.05), Inches(2.8), Inches(6.05), Inches(3.4)).table
for j, w in enumerate([1.05] + [1.25] * len(E6)):
    tbl.columns[j].width = Inches(w)
for i in range(len(B6) + 1):
    tbl.rows[i].height = Inches(0.42 if i == 0 else 0.6)
    for j in range(len(E6) + 1):
        cell = tbl.cell(i, j)
        cell.margin_left = cell.margin_right = Inches(0.04)
        cell.margin_top = cell.margin_bottom = Inches(0.02)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.fill.solid()
        b = B6[i - 1] if i else None
        cell.fill.fore_color.rgb = rgb(C["blue_pale"] if i == 0 else (C["violet_pale"] if b == 128 else C["bg"]))
        tf = cell.text_frame
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        if i == 0:
            r = p.add_run(); r.text = "배치 \\ 에폭" if j == 0 else f"{E6[j - 1]}"
            set_font(r, 10, "bold", C["blue"])
        elif j == 0:
            r = p.add_run(); r.text = f"{b}" + ("  (Lab)" if b == 10 else "")
            set_font(r, 11.5, "bold", C["violet"] if b == 128 else C["ink"])
        else:
            e = E6[j - 1]
            auc, sec = be(b, e, 'val_auc'), be(b, e, 'train_sec')
            r = p.add_run(); r.text = f"{auc:.3f}"
            set_font(r, 12, "bold", C["ink"] if auc >= 0.84 else C["mute"])
            p2 = tf.add_paragraph(); p2.alignment = PP_ALIGN.CENTER
            r2 = p2.add_run(); r2.text = f"{sec / 60:.0f}분" if sec >= 600 else f"{sec:.0f}초" if sec >= 10 else f"{sec:.1f}초"
            set_font(r2, 9.5, "semibold", C["red"] if sec >= 600 else C["mute"])
tparts.append(s.shapes[-1])
tparts.append(text(s, 1.1, 6.38, 6.0, 0.5, [[("Lab 6 설정(배치 10 · 150 에폭)이면 ", {"color": C["sub"]}), (f"모델 1개 {be(10, 150, 'train_sec') / 60:.0f}분 → 10-fold 약 {be(10, 150, 'train_sec') * 10 * 1.35 / 3600:.0f}시간", {"color": C["red"], "weight": "bold"})]], size=12.5, weight="medium"))
group(s, tparts, "anim1_table")
group(s, [rect(s, 7.55, 2.15, 4.98, 3.55, C["card"], radius=0.25),
          text(s, 7.85, 2.32, 4.5, 0.35, [[("AUC × 학습 시간  ", {"weight": "bold"}), ("로그 눈금", {"color": C["mute"]})]],
               size=12, weight="medium"),
          s.shapes.add_picture(str(FIG / "lab6_auc_time.png"), Inches(7.79), Inches(2.76), width=Inches(4.5))],
      "anim2_curve")
group(s, [rect(s, 7.55, 5.85, 4.98, 1.2, C["blue_pale"], radius=0.2),
          text(s, 7.8, 5.85, 4.6, 1.2,
               [[("배치 128 + EarlyStopping 선택", {"weight": "extrabold", "color": C["ink"], "size": 15})],
                [(f"AUC는 0.851~0.855로 모두 비슷 · 시간은 최대 {be(10, 150, 'train_sec') / be(2048, 150, 'train_sec'):.0f}배 차이", {"color": C["sub"]})]],
               size=12, weight="medium", anchor="m", line=1.35)], "anim3_pick")
notes(s, f"""왼쪽 표는 배치 크기 다섯 가지를 각각 150 에폭까지 학습하며, 칸마다 학습에 쓰지 않은 33%로 잰 AUC와 그때까지의 학습 시간을 적은 것입니다. 추정 없이 모두 직접 쟀습니다.
150 에폭 기준 AUC는 0.851에서 0.855 사이로 거의 같았습니다. 반면 시간은 Lab 6 설정인 배치 10이 {be(10, 150, 'train_sec') / 60:.0f}분, 배치 2048이 {be(2048, 150, 'train_sec'):.0f}초로 100배 가까이 차이 났습니다.
배치가 작으면 한 바퀴에 가중치를 더 자주 고쳐서 적은 에폭으로도 빨리 배우고, 배치가 크면 한 번 계산이 빨라서 시간이 덜 듭니다. 배치 512는 50 에폭 뒤로 AUC가 더 오르지 않았는데, 이게 EarlyStopping이 필요한 이유입니다.
그래서 0.854로 최상위권이면서 배치 10보다 {speed:.0f}배 빠른 배치 128을 골랐고, 에폭은 EarlyStopping으로 정했습니다.""")

# =====================================================================
# Lab 6 - (3) 발견 · 결론 : MLP vs 기존 모델
# =====================================================================
s = new_slide()
lab_header(s, 6, 3, "Keras MLP", L6['title'],
           terms=[("층화 10-fold", "Lab 3부터 쓴 같은 시험 기준"), ("± 값", "10번 시험 점수의 흔들림")])
group(s, [rect(s, 0.8, 2.15, 6.0, 4.05, C["card"], radius=0.25),
          text(s, 1.1, 2.33, 5.4, 0.35, [[("모델별 AUC  ", {"weight": "bold", "color": C["ink"]}),
                                         ("모두 층화 10-fold · 막대 끝 = 흔들림", {"color": C["mute"]})]], size=13),
          s.shapes.add_picture(str(FIG / "lab6_compare.png"), Inches(1.15), Inches(2.8), width=Inches(5.3))],
      "anim1_chart")
finds = [("01", f"MLP {L6['cv_auc']:.3f} ± {L6['cv_std']:.3f}", f"fold당 {L6['fold_sec']:.0f}초 · 평균 {L6['mean_stop']:.0f} 에폭에서 멈춤", C["blue"]),
         ("02", f"Grid 6가지: {L6['grid_worst']:.3f} ~ {L6['grid_best']:.3f}", "optimizer·초기화 차이는 흔들림 안", C["blue"]),
         ("03", L6['vs_rf_head'], L6['vs_rf_body'], C["red"])]
parts = []
for i, (no, head, body, col) in enumerate(finds):
    yy = 2.15 + i * 0.72
    parts += [text(s, 7.1, yy, 0.6, 0.4, no, size=17, color=col, weight="extrabold"),
              text(s, 7.65, yy, 4.9, 0.4, head, size=16, weight="bold"),
              text(s, 7.65, yy + 0.34, 4.9, 0.35, body, size=12.5, color=C["sub"], weight="medium")]
group(s, parts, "anim2_finds")
conclusion_card(s, 7.05, 4.4, 5.48, 1.8, [
    [("① ", {"color": C["violet"], "weight": "bold"}), (L6['concl1'], {})],
    [("② ", {"color": C["violet"], "weight": "bold"}), (L6['concl2'], {})],
], anim=3)
next_bar(s, "Lab 밖 방법까지 더하면 더 오를까?", "최종 비교", "최고 AUC", L6['tracker'], anim=4)
notes(s, L6['notes'])


# =====================================================================
# 최종 결론 : Lab 기반 모델 vs 최적 모델 (노트북 09)
# =====================================================================
s = new_slide()
term_line(s, [("최적 모델", "Lab 밖 문헌 방법(새 변수 + 그래디언트 부스팅, 부록 C)을 더해 튜닝한 모델")])
stage_track(s, [6])
text(s, 0.8, 0.55, 6.5, 0.35, [[("최종 결론", {"color": C["violet"], "weight": "bold"}),
                                ("  결과 제시", {"color": C["mute"], "weight": "medium"})]], size=14, name="!!kicker")
text(s, 0.8, 0.98, 11.8, 0.9, TITLES[len(prs.slides)], size=30, weight="extrabold", spacing=-1, name="!!title")
cmp_rows = [("", "Lab 기반\n정리 + RF", "최적\n새 변수 + 부스팅"),
            ("AUC (층화 10-fold)", "0.856", "0.858  (10/10 우세)"),
            ("15만 건 학습 · 1,000명 예측", "5.0초 · 42ms", "4.6초 · 13ms"),
            ("파일 크기", "5.5 MB", "0.6 MB"),
            ("만들기 비용", "설정 1개 튜닝", "설정 6개 · 탐색 14분")]
lparts = [rect(s, 0.8, 2.15, 6.6, 3.55, C["card"], radius=0.25)]
tbl = s.shapes.add_table(len(cmp_rows), 3, Inches(1.0), Inches(2.3), Inches(6.2), Inches(2.6)).table
for j, w in enumerate([2.3, 1.75, 2.15]):
    tbl.columns[j].width = Inches(w)
for i, row in enumerate(cmp_rows):
    tbl.rows[i].height = Inches(0.6 if i == 0 else 0.55)
    for j, val in enumerate(row):
        cell = tbl.cell(i, j)
        cell.margin_left = cell.margin_right = Inches(0.06)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.fill.solid()
        cell.fill.fore_color.rgb = rgb(C["blue_pale"] if i == 0 else (C["violet_pale"] if j == 2 else C["bg"]))
        for k, ln in enumerate(val.split("\n")):
            pp = cell.text_frame.paragraphs[0] if k == 0 else cell.text_frame.add_paragraph()
            pp.alignment = PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER
            r = pp.add_run(); r.text = ln
            if i == 0:
                set_font(r, 11 if k == 0 else 9, "bold" if k == 0 else "medium", C["blue"] if j == 1 else (C["violet"] if j == 2 else C["mute"]))
            else:
                set_font(r, 11 if j == 0 else 12, "semibold" if j == 0 else "bold", C["sub"] if j == 0 else C["ink"])
lparts.append(s.shapes[-1])
group(s, lparts, "anim1_compare")
# AUC 여정
journey = [("찍기", "0.50"), ("Lab 2", "0.824"), ("Lab 4", "0.844"), ("Lab 5", "0.856"), ("최적", "0.858")]
jparts = [rect(s, 7.6, 2.15, 4.93, 1.16, C["card"], radius=0.22),
          text(s, 7.85, 2.27, 4.5, 0.3, [[("AUC 여정  ", {"weight": "bold"}), ("MLP 0.850 · GAN 0.817은 넘지 못함", {"color": C["mute"]})]],
               size=11.5, weight="medium")]
for k, (lab, v) in enumerate(journey):
    x = 7.85 + k * 0.93
    last = k == len(journey) - 1
    jparts += [text(s, x, 2.66, 0.85, 0.25, lab, size=9.5, color=C["mute"], weight="semibold", align="c"),
               text(s, x, 2.9, 0.85, 0.35, v, size=14, color=C["violet"] if last else C["ink"], weight="extrabold", align="c")]
    if k:
        jparts.append(text(s, x - 0.16, 2.92, 0.2, 0.3, "→", size=11, color=C["line"], align="c"))
group(s, jparts, "anim2_journey")
conclusion_card(s, 7.6, 3.45, 4.93, 2.38, [
    [("① ", {"color": C["violet"], "weight": "bold"}), ("최종: 그래디언트 부스팅 (입력 14개)", {"weight": "bold"})],
    [("    학습률 0.021 · 트리 608개 · 잎 최소 127명", {"size": 11.5, "color": C["sub"]})],
    [("② ", {"color": C["violet"], "weight": "bold"}), ("Lab 모델과 차이 +0.002 (10조각 모두 우세)", {})],
    [("③ ", {"color": C["violet"], "weight": "bold"}), ("Lab 과정만으로도 최적에 가까운 모델", {})],
], anim=3)
next_bar(s, "이 최종 모델, 실무에서도 쓸 수 있을까?", "실무 점검", "최고 AUC", "0.858", anim=4)
notes(s, """Lab 기반 모델과, Lab 밖 문헌에서 효과가 확인된 방법(새 변수 + 그래디언트 부스팅, 부록 C)을 더한 최적 모델을 같은 기준으로 비교했습니다. 최적 모델의 설정은 각 조각의 학습 데이터 안에서만 찾는 중첩 교차검증으로 평가했습니다.
최적 모델은 0.858로 열 조각 모두에서 Lab 모델보다 높았고, 예측도 세 배 빠르고 파일도 9분의 1이었습니다. 다만 차이는 0.002로, Lab 과정만으로도 최적에 가까운 모델을 만들었다는 뜻이기도 합니다.
다음 장에서는 이 모델을 실무에서 쓸 수 있는지, 지금까지 쓴 ROC-AUC 말고 PR-AUC로도 확인합니다.""")

# =====================================================================
# 최종 결론 2 : 실무에서 쓸 수 있을까 (ROC-AUC의 한계 → PR-AUC, notebooks/09c_pr_auc.py)
# =====================================================================
PR = pd.read_csv(ROOT / "notebooks" / "09c_pr_auc_summary.csv").set_index("model")
LAB, OPT = PR.loc["Lab RF"], PR.loc["최적 부스팅"]
s = new_slide()
term_line(s, [("PR-AUC", "연체로 판정한 사람 중 진짜 연체자 비율(정밀도)로 본 점수 · 찍기 = 연체율 0.067"),
              ("비용 5:1", "연체자 1명 놓침 = 정상 고객 5명 거절로 계산")])
stage_track(s, [6])
text(s, 0.8, 0.55, 6.5, 0.35, [[("최종 결론", {"color": C["violet"], "weight": "bold"}),
                                ("  실무 점검", {"color": C["mute"], "weight": "medium"})]], size=14, name="!!kicker")
text(s, 0.8, 0.98, 11.8, 0.9, TITLES[len(prs.slides)], size=30, weight="extrabold", spacing=-1, name="!!title")
# ① ROC-AUC의 한계 (앞에서 본 13장)
group(s, [rect(s, 0.8, 2.15, 4.3, 2.85, C["card"], radius=0.22),
          text(s, 1.05, 2.3, 3.9, 0.35, [[("① ROC-AUC의 한계  ", {"weight": "bold"}), ("앞에서 본 13장", {"color": C["mute"]})]],
               size=12, weight="medium"),
          text(s, 1.05, 2.75, 3.9, 0.5, [[("AUC 0.844 ", {"weight": "extrabold", "color": C["ink"], "size": 22}),
                                         ("'좋음'", {"color": C["blue"], "weight": "bold"})]], size=15),
          text(s, 1.05, 3.3, 3.9, 0.5, [[("그런데 판정 중 진짜 연체자 ", {"color": C["sub"]}),
                                         ("20%", {"weight": "extrabold", "color": C["red"], "size": 22})]], size=14, weight="medium"),
          text(s, 1.05, 3.85, 3.9, 0.35, [[("정상 고객 ", {"color": C["sub"]}), ("9,766명", {"weight": "bold", "color": C["red"]}),
                                          ("을 잘못 거절", {"color": C["sub"]})]], size=13, weight="medium"),
          text(s, 1.05, 4.3, 3.9, 0.6, ["정상 고객이 14만 명이라 오판이 비율로는", "작아 보인다 → 정밀도를 보는 PR-AUC로"],
               size=11, color=C["mute"], weight="medium", line=1.3)], "anim1_roc")
# ② PR-AUC로 다시 보기
rows22 = [("", "ROC-AUC", "PR-AUC"), ("찍기", "0.500", f"{0.0668:.3f}"),
          ("Lab 모델 (RF)", f"{LAB.roc_auc:.3f}", f"{LAB.pr_auc:.3f}"),
          ("최적 (부스팅)", f"{OPT.roc_auc:.3f}", f"{OPT.pr_auc:.3f}")]
rp = [rect(s, 5.3, 2.15, 7.23, 2.85, C["card"], radius=0.22),
      text(s, 5.55, 2.3, 6.8, 0.35, [[("② PR-AUC로 다시 보기  ", {"weight": "bold"}), ("같은 층화 10-fold", {"color": C["mute"]})]],
           size=12, weight="medium"),
      s.shapes.add_picture(str(FIG / "final_pr_curve.png"), Inches(5.4), Inches(2.68), width=Inches(3.65))]
t22 = s.shapes.add_table(len(rows22), 3, Inches(9.2), Inches(2.8), Inches(3.15), Inches(1.6)).table
for j, w in enumerate([1.35, 0.9, 0.9]):
    t22.columns[j].width = Inches(w)
for i, row in enumerate(rows22):
    t22.rows[i].height = Inches(0.38)
    for j, val in enumerate(row):
        cell = t22.cell(i, j)
        cell.margin_left = cell.margin_right = Inches(0.05)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.fill.solid()
        cell.fill.fore_color.rgb = rgb(C["blue_pale"] if i == 0 else (C["violet_pale"] if j == 2 else C["bg"]))
        r = cell.text_frame.paragraphs[0].add_run(); r.text = val
        cell.text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER
        if i == 0:
            set_font(r, 10, "bold", C["blue"] if j == 1 else C["violet"])
        else:
            set_font(r, 10.5 if j == 0 else 12, "semibold" if j == 0 else "bold",
                     C["sub"] if j == 0 else (C["violet"] if j == 2 else C["ink"]))
rp.append(s.shapes[-1])
rp.append(text(s, 9.2, 4.5, 3.2, 0.45, ["PR-AUC는 찍기의 약 5배", "ROC가 가린 오판이 보인다"],
               size=10.5, color=C["sub"], weight="medium", line=1.25))
group(s, rp, "anim2_pr")
# 실무에서 쓸 수 있나?
num = lambda k: (k, {"color": C["violet"], "weight": "bold"})
group(s, [rect(s, 0.8, 5.12, 11.73, 1.32, C["violet_pale"], radius=0.22),
          rect(s, 1.1, 5.3, 0.06, 0.36, C["violet"]),
          text(s, 1.3, 5.25, 2.4, 0.45, "실무 활용 결론", size=15, weight="bold", color=C["violet"]),
          text(s, 1.3, 5.66, 2.4, 0.7, [[("위험 고객 '선별 도구'", {"color": C["ink"], "weight": "bold", "size": 12.5})], "단독 판정이 아닌 심사 보조용"], size=11, color=C["sub"], weight="semibold", line=1.3),
          text(s, 3.75, 5.22, 8.6, 1.2,
               [[num("① "), ("기준 0.17: 연체자 ", {}), (f"{LAB.th017_recall:.0%}", {"weight": "bold"}),
                 (" 잡지만 판정 3명 중 2명은 정상 고객 (정밀도 ", {}), (f"{LAB.th017_precision:.0%}", {"weight": "bold", "color": C["red"]}), (")", {})],
                [num("② "), ("위험 상위 10%만 심사해도 연체자 ", {}), (f"{LAB.top10_recall:.0%}", {"weight": "bold"}),
                 (f" · 그 안의 연체율 {LAB.top10_precision:.0%} (평균 6.7%의 5배)", {})],
                [num("③ "), ("비용 5:1이면 같은 RF에서 기준 0.5 → 0.17로 손실 ", {}), ("31% 감소", {"weight": "bold"})]],
               size=13, weight="medium", line=1.38)], "anim3_use")
group(s, [rect(s, 0.8, 6.55, 11.73, 0.72, C["dark"], radius=0.22),
          text(s, 1.15, 6.55, 11.2, 0.72,
               [[("느낀 점  ", {"color": C["violet"], "weight": "bold", "spacing": 1}),
                 ("데이터 형식 · 전처리 · 알고리즘 · 평가 지표를 바꿀 때마다 결과가 달라졌고, ", {"color": C["dark_ink"], "weight": "semibold"}),
                 ("신경망(0.850)이라고 당연히 좋은 것은 아니었다", {"color": "#82AAFF", "weight": "bold"})]],
               size=13, anchor="m")], "anim4_lesson")
notes(s, f"""앞에서 본 13장에서, 학습 가중치로 연체자를 75% 잡았을 때 AUC는 0.844로 좋았지만 연체로 판정한 사람 중 진짜 연체자는 20%였고 정상 고객 9,766명을 거절했습니다. ROC-AUC는 정상 고객 14만 명 중 몇 %를 잘못 거절했는지로 보기 때문에, 정상이 많은 데이터에서는 오판이 작아 보입니다.
그래서 정밀도(연체 판정 중 진짜 연체자)와 재현율(전체 연체자 중 잡은 비율)로 보는 PR-AUC를 같은 층화 10-fold로 다시 쟀습니다. 찍기 수준은 연체율 0.067입니다. Lab 모델 {LAB.pr_auc:.3f}, 최적 모델 {OPT.pr_auc:.3f}로 찍기의 약 5배였고, 최적 모델이 10조각 모두에서 높아 모델 선택은 그대로입니다.
판정 기준 0.17에서는 연체자의 {LAB.th017_recall:.1%}를 잡지만 판정한 고객 중 진짜 연체자는 {LAB.th017_precision:.1%}입니다. 자동 거절에 쓰면 정상 고객을 많이 놓치므로, 위험한 순서로 심사 대상을 고르고 사람이 최종 판단하는 용도가 맞습니다. 위험 상위 10%에 연체자의 {LAB.top10_recall:.1%}가 있고 그 안의 연체율은 {LAB.top10_precision:.1%}입니다.
비용 5:1 가정에서 같은 RF의 기준을 0.5에서 0.17로 낮추면 손실이 31% 줄고, 같은 기준에서 모델을 부스팅으로 바꾼 효과는 0.5%입니다. 0.17은 가중치 없이 학습한 RF의 확률이 실제와 맞아서(부록 B) 이론값 1/6과 거의 같습니다.
느낀 점은, 데이터 형식, 전처리, 알고리즘, 평가 지표를 바꿀 때마다 결과가 달라졌다는 것입니다. 로그 변환은 회귀에만 효과가 있었고, 가중치는 연체자를 잡는 데 결정적이었으며, 신경망은 튜닝한 랜덤 포레스트를 넘지 못했습니다. 그리고 ROC-AUC만 보면 좋아 보이던 모델도 PR-AUC로 보면 쓰임새가 달라졌습니다.""")

# =====================================================================
# 부록 A. 15만 건을 어떻게 돌렸나 (발표하지 않음 · 과제 공지 'MLP·15만 건 어떻게 돌릴지' 요약)
# =====================================================================
s = new_slide()
stage_track(s, [2, 3, 4, 5, 6])
text(s, 0.8, 0.55, 6.5, 0.35, [[("부록 A", {"color": C["violet"], "weight": "bold"}),
                                ("  발표 시간 밖 · 질문 대비", {"color": C["mute"], "weight": "medium"})]], size=14, name="!!kicker")
text(s, 0.8, 0.98, 11.8, 0.9, "15만 건 처리 방법 요약", size=30, weight="extrabold", spacing=-1, name="!!title")
text(s, 0.8, 1.62, 11.8, 0.3, [[("원칙  ", {"color": C["violet"], "weight": "bold"}),
                                ("데이터는 줄이지 않고 방법을 바꿨다. 줄인 곳(SVM 1만 건, 산점도 5,000건)은 장표에 표시했다", {"color": C["sub"]})]],
     size=12.5, weight="medium")
_ab = _sc[_sc.model.isin(["AdaBoost", "AdaBoost+"])]
rowsA = [("단계", "15만 건에서 생긴 문제", "선택한 방법", "결과"),
         ("Lab 3 · LOOCV", "15만 번 학습 · 1회 0.49초 → 약 20시간", "층화 10-fold로 대체", "0.824 ± 0.009"),
         ("Lab 4 · SVM", "학습 시간 ∝ n² · 10-fold 약 1시간", "같은 1만 건 표본에서 6개 재비교", "0.579 → 가중치 0.810"),
         ("Lab 5 · 튜닝", "후보 4개 × 설정 × 10-fold", "안쪽 3-fold로 설정 선택 · 병렬 계산", "RF 0.856 선택"),
         ("Lab 6 · MLP 학습", "배치 10이면 모델 1개 26분 · 10-fold 약 6시간", "배치 128 + EarlyStopping", "fold당 67초 · 0.850"),
         ("Lab 6 · Grid Search", "54가지 × 3-fold → 이틀 이상", "배치는 따로 재고 격자는 6가지로", "0.849 ~ 0.850"),
         ("최종 모델", "15만 건 전체 학습", "정리 + RF (잎 176명)", "학습 5초 · 5.5 MB")]
cwA = [2.0, 3.85, 3.45, 2.4]
tA = s.shapes.add_table(len(rowsA), 4, Inches(0.8), Inches(2.15), Inches(sum(cwA)), Inches(0.6 * len(rowsA))).table
for j, w in enumerate(cwA):
    tA.columns[j].width = Inches(w)
for i, row in enumerate(rowsA):
    tA.rows[i].height = Inches(0.5 if i == 0 else 0.62)
    for j, val in enumerate(row):
        cell = tA.cell(i, j)
        cell.margin_left = cell.margin_right = Inches(0.1)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.fill.solid()
        cell.fill.fore_color.rgb = rgb(C["blue_pale"] if i == 0 else (C["card"] if i % 2 else C["bg"]))
        r = cell.text_frame.paragraphs[0].add_run()
        r.text = val
        if i == 0:
            set_font(r, 11.5, "bold", C["blue"])
        else:
            set_font(r, 13 if j == 0 else 12.5, "bold" if j in (0, 3) else "medium",
                     C["violet"] if j == 3 else (C["ink"] if j != 1 else C["sub"]))
notes(s, """부록입니다. 발표 시간에는 넘기고, 질문이 나오면 씁니다.
과제 공지의 'MLP는 무조건 사용, 15만 데이터를 어떻게 돌릴지 고민'에 대한 답을 한 장에 모았습니다.
원칙은 데이터를 줄이지 않고 방법을 바꾸는 것이었습니다. 시간이 너무 걸리는 방법(LOOCV, SVM, Lab 6의 배치 10과 54가지 격자)은 같은 목적을 이루는 다른 방법으로 바꿨고, 표본을 쓴 곳은 SVM 비교와 산점도 두 곳뿐입니다.""")

# =====================================================================
# 부록 B. 예측 확률이 실제 연체율과 맞는가 (과제 목적 = 연체 '확률' 예측)
# =====================================================================
CAL = pd.read_csv(ROOT / "notebooks" / "09b_calibration.csv")
s = new_slide()
stage_track(s, [6])
text(s, 0.8, 0.55, 6.5, 0.35, [[("부록 B", {"color": C["violet"], "weight": "bold"}),
                                ("  발표 시간 밖 · 질문 대비", {"color": C["mute"], "weight": "medium"})]], size=14, name="!!kicker")
text(s, 0.8, 0.98, 11.8, 0.9, "예측 확률과 실제 연체율의 일치", size=30, weight="extrabold", spacing=-1, name="!!title")
text(s, 0.8, 1.62, 11.8, 0.3, [[("과제 목적  ", {"color": C["violet"], "weight": "bold"}),
                                ("2년 안에 연체할 '확률'을 예측 → Lab 모델(정리 + RF)의 확률이 실제와 맞는지 확인", {"color": C["sub"]})]],
     size=12.5, weight="medium")
group(s, [rect(s, 0.8, 2.15, 7.2, 4.9, C["card"], radius=0.25),
          text(s, 1.1, 2.33, 6.6, 0.35, [[("위험 점수 10구간  ", {"weight": "bold", "color": C["ink"]}),
                                         ("구간마다 1만 5천 명 · 층화 10-fold 예측", {"color": C["mute"]})]], size=12.5),
          s.shapes.add_picture(str(FIG / "final_calibration.png"), Inches(1.0), Inches(2.8), width=Inches(6.8))],
      "anim1_cal")
top, bot = CAL.iloc[-1], CAL.iloc[0]
findsB = [("01", "평균 예측 6.68% = 실제 6.68%", "15만 명 전체로 보면 확률의 크기가 그대로 맞음", C["blue"]),
          ("02", f"상위 10%: 예측 {top['평균예측']:.1%} · 실제 {top['실제연체율']:.1%}",
           f"가장 안전한 10%도 예측 {bot['평균예측']:.1%} · 실제 {bot['실제연체율']:.1%}", C["blue"]),
          ("03", "점수를 그대로 확률로 사용 가능", "가중치 없이 학습한 RF라 보정이 필요 없음", C["blue"])]
parts = []
for i, (no, head, body, col) in enumerate(findsB):
    yy = 2.25 + i * 0.95
    parts += [text(s, 8.3, yy, 0.6, 0.4, no, size=17, color=col, weight="extrabold"),
              text(s, 8.85, yy, 3.8, 0.4, head, size=15, weight="bold"),
              text(s, 8.85, yy + 0.36, 3.8, 0.5, body, size=12, color=C["sub"], weight="medium")]
group(s, parts, "anim2_findsB")
group(s, [rect(s, 8.3, 5.3, 4.23, 1.75, C["violet_pale"], radius=0.22),
          text(s, 8.55, 5.3, 3.8, 1.75,
               [[("그래서", {"color": C["violet"], "weight": "bold"})],
                [("'연체 확률 34%' 같은 점수를 그대로 심사와", {})],
                [("판정 기준(22장의 0.17)에 쓸 수 있다", {})]],
               size=13, weight="medium", anchor="m", line=1.4)], "anim3_soB")
notes(s, """부록입니다. 과제 목적이 '2년 안에 연체할 확률' 예측이라, 순위(AUC)만이 아니라 확률 자체가 맞는지도 확인했습니다.
15만 명이 각자 시험 조각일 때 받은 예측 확률을 모아, 위험 순서로 10구간(구간마다 1만 5천 명)으로 나눠 평균 예측과 실제 연체율을 비교했습니다.
전체 평균은 예측과 실제가 모두 6.68%였고, 가장 위험한 10%는 예측 32%, 실제 34%로 거의 맞았습니다. Brier 점수는 0.051로, 늘 6.68%라고 답하는 경우(0.062)보다 낮습니다.
가중치 없이 학습한 랜덤 포레스트라 점수가 실제 확률에 가깝고, 그래서 22장의 비용 기준값 0.17도 이론값 1/6과 거의 같게 나왔습니다.""")

# =====================================================================
# 부록 C. 보충 실험: 문헌 방법 + GAN 오버샘플링 (Lab 밖, 발표하지 않음)
# =====================================================================
def extra_header(s, chapter, title, terms, stages):
    if terms:
        term_line(s, terms)
    stage_track(s, stages)
    text(s, 0.8, 0.55, 6.5, 0.35, [[("부록 C · 보충 실험", {"color": C["violet"], "weight": "bold"}),
                                    (f"  {chapter}", {"color": C["mute"], "weight": "medium"})]], size=14, name="!!kicker")
    text(s, 0.8, 0.98, 11.8, 0.9, TITLES.get(len(prs.slides), title), size=30, weight="extrabold", spacing=-1, name="!!title")


LIT = pd.read_csv(ROOT / "notebooks" / "07_literature_fold_scores.csv").pivot(index="fold", columns="model", values="auc")
LBASE = LIT["A 기준: 정리 + RF (Lab 5)"]
s = new_slide()
extra_header(s, "문헌 방법 · GAN 오버샘플링", "문헌 방법과 GAN 오버샘플링의 효과 검증",
             [("오버샘플링", "적은 쪽(연체자) 데이터를 늘려 균형 맞추기"), ("SMOTE", "연체자 두 명 사이에 새 점 만들기"),
              ("GAN", "생성자와 판별자가 경쟁하며 가짜 데이터 생성")], [5])
lit_rows = [("새 변수 + 부스팅", "D 새 변수 + 부스팅", "Ichim & Issa 2025, Kaggle"),
            ("정리 + 부스팅", "B 정리 + 부스팅", "Gidlow 2022"),
            ("Balanced RF (잎 20)", "G 새 변수 + Balanced RF (잎 20)", "Ichim & Issa 2025"),
            ("기준: 정리 + RF (Lab 5)", "A 기준: 정리 + RF (Lab 5)", "Lab 1~6"),
            ("쌓기 (RF·부스팅·LR)", "H 새 변수 + 쌓기 (RF·부스팅·LR → LR)", "Kaggle 2011 1위"),
            ("Balanced RF (잎 1)", "F 새 변수 + Balanced RF (잎 1)", "Ichim & Issa 2025")]
lparts = [rect(s, 0.8, 2.15, 6.1, 3.4, C["card"], radius=0.25),
          text(s, 1.1, 2.29, 5.6, 0.35, [[("① 문헌 방법  ", {"weight": "bold"}),
                                         ("같은 10조각에서 기준과 조각별 비교", {"color": C["mute"]})]], size=12, weight="medium")]
tbl = s.shapes.add_table(len(lit_rows) + 1, 4, Inches(1.05), Inches(2.68), Inches(5.6), Inches(2.5)).table
for j, w in enumerate([2.35, 1.0, 1.15, 1.1]):
    tbl.columns[j].width = Inches(w)
hdr = ["방법", "AUC", "기준 대비", "이긴 조각"]
for i in range(len(lit_rows) + 1):
    tbl.rows[i].height = Inches(0.34 if i == 0 else 0.36)
    for j in range(4):
        cell = tbl.cell(i, j)
        cell.margin_left = cell.margin_right = Inches(0.06)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.fill.solid()
        base_row = i > 0 and lit_rows[i - 1][1].startswith("A ")
        cell.fill.fore_color.rgb = rgb(C["blue_pale"] if i == 0 else (C["violet_pale"] if base_row else C["bg"]))
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER
        if i == 0:
            r = p.add_run(); r.text = hdr[j]; set_font(r, 10, "bold", C["blue"]); continue
        label, key, src = lit_rows[i - 1]
        col = LIT[key]; diff = col - LBASE
        if j == 0:
            r = p.add_run(); r.text = label; set_font(r, 11, "semibold", C["ink"])
        elif j == 1:
            r = p.add_run(); r.text = f"{col.mean():.4f}"; set_font(r, 11.5, "bold", C["ink"])
        elif j == 2:
            r = p.add_run(); r.text = "–" if base_row else f"{diff.mean():+.4f}"
            set_font(r, 11.5, "bold", C["mute"] if base_row else (C["blue"] if diff.mean() > 0 else C["red"]))
        else:
            r = p.add_run(); r.text = "–" if base_row else f"{(diff > 0).sum()}/10"
            set_font(r, 11.5, "semibold", C["sub"])
lparts.append(s.shapes[-1])
lparts.append(text(s, 1.1, 5.24, 5.6, 0.25, "새 변수 4개: 월 부채액 · 1인당 소득 · 월 잔액 · 사용 비율>1 표시 (근거: Ichim & Issa 2025 외)",
                   size=9.5, color=C["mute"], weight="medium"))
group(s, lparts, "anim1_lit")
group(s, [rect(s, 7.1, 2.15, 5.43, 3.4, C["card"], radius=0.25),
          text(s, 7.4, 2.29, 5.0, 0.35, [[("② GAN으로 연체자 늘리기  ", {"weight": "bold"}),
                                         ("NOTE 논문 방식 · 분류기는 같은 RF", {"color": C["mute"]})]], size=12, weight="medium"),
          s.shapes.add_picture(str(FIG / "lab8_gan_compare.png"), Inches(7.55), Inches(2.72), width=Inches(4.55))],
      "anim2_gan")
# 보충 실험의 결론 (발표 흐름에서 이 장이 무엇을 알려 주는지)
num = lambda k: (k, {"color": C["violet"], "weight": "bold"})
group(s, [rect(s, 0.8, 5.65, 11.73, 1.42, C["violet_pale"], radius=0.22),
          rect(s, 1.1, 5.85, 0.06, 0.36, C["violet"]),
          text(s, 1.3, 5.8, 2.2, 0.45, "보충 실험의 결론", size=16, weight="bold", color=C["violet"]),
          text(s, 1.3, 6.22, 2.3, 0.7, ["Lab 밖 방법을 더해도", "크게 달라지지 않았다"], size=11.5, color=C["sub"], weight="medium", line=1.3),
          text(s, 3.75, 5.78, 8.6, 1.2,
               [[num("① "), ("문헌 방법을 더해도 최고 ", {}), ("+0.0016", {"weight": "bold"}),
                 (" (새 변수 + 부스팅) → Lab 모델(0.856)이 이미 최적에 가까움", {})],
                [num("② "), ("데이터 늘리기는 효과 없음 ", {}), ("(GAN 0.817 < 0.856)", {"weight": "bold"}),
                 (" → 불균형은 학습 가중치나 판정 기준으로", {})],
                [num("③ "), ("나누기 전에 늘리면 ", {}), ("0.950", {"weight": "bold", "color": C["red"]}),
                 ("으로 부풀려짐 → 늘리기는 학습 데이터에서만", {})]],
               size=13.5, weight="medium", line=1.4)], "anim3_concl21")
notes(s, """Lab 밖에서, 문헌에서 이 데이터에 효과가 있었다고 보고한 방법을 하나씩 더해 봤습니다. 비교는 Lab과 같은 층화 10조각에서, 조각마다 기준 모델과의 차이를 계산했습니다.
가장 나은 것은 새 변수와 그래디언트 부스팅으로 0.8573이었고, 10조각 중 9조각에서 기준보다 높았지만 차이는 0.0016으로 매우 작았습니다. Balanced RF는 잎 크기에 따라 갈렸고, 여러 모델을 쌓는 방법은 오히려 낮아졌습니다.
오른쪽은 교수님 논문 NOTE처럼 GAN으로 연체자 데이터를 만들어 늘린 실험입니다. 늘리기는 각 조각의 학습 데이터 안에서만 하고, 시험은 진짜 고객으로만 했습니다.
결과는 늘리지 않은 0.856이 가장 높았고, 복제와 SMOTE는 비슷하거나 조금 낮았으며, 기본 GAN은 0.817로 떨어졌습니다. GAN이 만든 연체자가 부채 비율 등에서 진짜와 달라 분류기가 가짜의 특징까지 배운 것으로 보입니다. 논문의 오토인코더와 조건부 WGAN이 아닌 기본 GAN이라는 한계도 있습니다.
그리고 같은 GAN도 나누기 전에 전체를 늘리면 시험 데이터에 가짜 연체자가 섞여 0.950이 나왔습니다. 평가 순서 하나로 점수가 이만큼 달라질 수 있다는 점을 확인했습니다.""")


prs.save(DRAFT)
print("saved", DRAFT)
write_script_md(ROOT / "slides" / "script.md")
print("saved", ROOT / "slides" / "script.md")
