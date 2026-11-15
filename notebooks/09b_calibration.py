"""과제 목적은 '2년 안에 연체할 확률' 예측 → Lab 기반 모델(정리 + RF 잎 176명)의 예측 확률이 실제 연체율과 맞는지 확인.

방법: 층화 10-fold(seed 7)로 15만 명 모두 '자기가 시험 조각일 때' 받은 확률을 모은 뒤,
확률 순서로 10개 구간(각 15,000명)으로 나눠 구간별 '평균 예측 확률'과 '실제 연체율'을 비교한다.
실행: python 09b_calibration.py → 09b_calibration.csv, ../figures/final_calibration.png
"""
import glob, os
import numpy
from pandas import read_csv, DataFrame, qcut
from matplotlib import pyplot, font_manager
from sklearn.base import clone
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import roc_auc_score, brier_score_loss
from sklearn.model_selection import StratifiedKFold
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import FunctionTransformer

for p in glob.glob(os.path.expandvars(r'%LOCALAPPDATA%\Microsoft\Windows\Fonts\Pretendard-*.ttf')):
    font_manager.fontManager.addfont(p)
pyplot.rcParams['font.family'] = 'Pretendard'
pyplot.rcParams['axes.unicode_minus'] = False
BLUE, GRAY, VIOLET, INK = '#3182F6', '#B0B8C1', '#7C5CFF', '#191F28'

seed = 7
names = ['target', 'revol', 'age', 'late30', 'debt', 'income', 'depend', 'defaulted', 'loans']
df = read_csv('../data/givemesomecredit.csv', names=names, header=0)
X = df.values[:, 1:9].astype(float)
Y = df.values[:, 0].astype(int)
iR, iA, iL, iD, iI, iP, iF, iO = range(8)


def clean(A):
    A = A.copy()
    code = numpy.isin(A[:, iL], [96, 98]).astype(float)
    fill = numpy.isin(A[:, iI], [6638, 6746]).astype(float)
    A[code == 1, iL] = 0
    return numpy.column_stack([A, code, fill])


lab_model = Pipeline([('f', FunctionTransformer(clean)),
                      ('m', RandomForestClassifier(n_estimators=100, max_features=3, min_samples_leaf=176,
                                                   random_state=seed, n_jobs=-1))])
oof = numpy.zeros(len(Y))
for tr, te in StratifiedKFold(10, shuffle=True, random_state=seed).split(X, Y):
    oof[te] = clone(lab_model).fit(X[tr], Y[tr]).predict_proba(X[te])[:, 1]

d = DataFrame({'p': oof, 'y': Y})
d['구간'] = qcut(d.p.rank(method='first'), 10, labels=False) + 1
cal = d.groupby('구간').agg(인원=('y', 'size'), 평균예측=('p', 'mean'), 실제연체율=('y', 'mean')).reset_index()
cal.to_csv('09b_calibration.csv', index=False)
print(f'AUC {roc_auc_score(Y, oof):.4f} · Brier {brier_score_loss(Y, oof):.4f} (늘 6.68%라고 하면 {brier_score_loss(Y, numpy.full(len(Y), Y.mean())):.4f})')
print(f'전체 평균 예측 {oof.mean():.4f} vs 실제 {Y.mean():.4f}')
print(cal.round(4).to_string(index=False))

fig, ax = pyplot.subplots(figsize=(7.5, 3.8))
x = numpy.arange(1, 11)
ax.bar(x - 0.2, cal.평균예측 * 100, width=0.4, color=GRAY, label='평균 예측 확률')
ax.bar(x + 0.2, cal.실제연체율 * 100, width=0.4, color=BLUE, label='실제 연체율')
for k in range(10):
    ax.text(x[k] + 0.2, cal.실제연체율[k] * 100 + 1, f'{cal.실제연체율[k] * 100:.0f}', ha='center', fontsize=9, color=INK)
ax.set_xticks(x); ax.set_xticklabels([f'{k}' for k in x])
ax.set_xlabel('위험 점수 구간 (1 = 가장 안전한 10%, 10 = 가장 위험한 10%)')
ax.set_ylabel('%'); ax.legend(frameon=False, loc='upper left')
ax.spines[['top', 'right']].set_visible(False); ax.grid(axis='y', color='#E5E8EB'); ax.set_axisbelow(True)
pyplot.tight_layout(); pyplot.savefig('../figures/final_calibration.png', dpi=200)
print('saved ../figures/final_calibration.png')
