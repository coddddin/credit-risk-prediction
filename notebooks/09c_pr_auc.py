"""실무 관점 점검: ROC-AUC 대신 PR-AUC(평균 정밀도)와 판정 기준별 정밀도·재현율.

ROC-AUC는 정상 고객 14만 명 중 몇 %를 잘못 거절했는지(FPR)를 보므로 불균형 데이터에서 오탐이 작아 보인다.
PR-AUC는 '연체로 판정한 사람 중 진짜 연체자 비율(정밀도)'을 보므로 찍기 기준선이 연체율(0.067)이다.

모델 (모두 같은 층화 10-fold, seed 7 · 09 노트북과 같은 조각)
  Lab 기반: 정리 + RF(잎 176명)
  최적: 새 변수 + 부스팅, 설정은 각 바깥 학습 구간 안의 3-fold RandomizedSearch 20개로만 고름 (09와 같은 중첩 절차)
  참고: 로그 + 표준화 + 로지스틱 회귀(학습 가중치)
실행: python 09c_pr_auc.py → 09c_pr_auc_folds.csv, 09c_pr_auc_summary.csv, ../figures/final_pr_curve.png
"""
import glob, os, time
import numpy
from pandas import read_csv, DataFrame
from matplotlib import pyplot, font_manager
from scipy.stats import loguniform, randint, uniform
from sklearn.base import clone
from sklearn.ensemble import RandomForestClassifier, HistGradientBoostingClassifier
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import roc_auc_score, average_precision_score, precision_recall_curve
from sklearn.model_selection import StratifiedKFold, RandomizedSearchCV
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import FunctionTransformer, StandardScaler

for p in glob.glob(os.path.expandvars(r'%LOCALAPPDATA%\Microsoft\Windows\Fonts\Pretendard-*.ttf')):
    font_manager.fontManager.addfont(p)
pyplot.rcParams['font.family'] = 'Pretendard'
pyplot.rcParams['axes.unicode_minus'] = False
BLUE, GRAY, VIOLET, INK, RED = '#3182F6', '#B0B8C1', '#7C5CFF', '#191F28', '#F04452'

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


def engineer(A):
    B = clean(A)
    fill = B[:, 9] == 1
    income = B[:, iI]
    monthly_debt = numpy.where(fill, B[:, iD], B[:, iD] * income)
    return numpy.column_stack([B, monthly_debt, income / (B[:, iP] + 1), income - monthly_debt, (B[:, iR] > 1).astype(float)])


def log_skewed(A):
    A = A.copy()
    for i in (iR, iL, iD, iI):
        A[:, i] = numpy.log1p(A[:, i])
    return A


lab_model = Pipeline([('f', FunctionTransformer(clean)),
                      ('m', RandomForestClassifier(n_estimators=100, max_features=3, min_samples_leaf=176, random_state=seed, n_jobs=-1))])
lr_model = Pipeline([('l', FunctionTransformer(log_skewed)), ('s', StandardScaler()),
                     ('m', LogisticRegression(solver='liblinear', class_weight='balanced'))])
hgb_space = {
    'm__learning_rate': loguniform(0.01, 0.2), 'm__max_iter': randint(200, 1001),
    'm__max_leaf_nodes': [15, 31, 63], 'm__min_samples_leaf': randint(20, 301),
    'm__l2_regularization': uniform(0, 5), 'm__max_features': uniform(0.5, 0.5),
}
hgb_pipe = Pipeline([('f', FunctionTransformer(engineer)), ('m', HistGradientBoostingClassifier(random_state=seed))])

oof = {k: numpy.zeros(len(Y)) for k in ('Lab RF', '최적 부스팅', 'LR 가중치')}
rows = []
for k, (tr, te) in enumerate(StratifiedKFold(10, shuffle=True, random_state=seed).split(X, Y)):
    t0 = time.time()
    oof['Lab RF'][te] = clone(lab_model).fit(X[tr], Y[tr]).predict_proba(X[te])[:, 1]
    oof['LR 가중치'][te] = clone(lr_model).fit(X[tr], Y[tr]).predict_proba(X[te])[:, 1]
    search = RandomizedSearchCV(hgb_pipe, hgb_space, n_iter=20, scoring='roc_auc', random_state=seed + k,
                                cv=StratifiedKFold(n_splits=3, shuffle=True, random_state=seed), n_jobs=-1)
    oof['최적 부스팅'][te] = search.fit(X[tr], Y[tr]).predict_proba(X[te])[:, 1]
    for name, p in oof.items():
        rows.append({'fold': k + 1, 'model': name, 'roc_auc': roc_auc_score(Y[te], p[te]),
                     'pr_auc': average_precision_score(Y[te], p[te])})
    print(f'fold {k + 1:>2}: ' + ' · '.join(f"{r['model']} ROC {r['roc_auc']:.4f} PR {r['pr_auc']:.4f}" for r in rows[-3:])
          + f' ({time.time() - t0:.0f}초)', flush=True)

f = DataFrame(rows)
f.to_csv('09c_pr_auc_folds.csv', index=False)
base = Y.mean()
summary = []
for name, p in oof.items():
    g = f[f.model == name]
    order = numpy.argsort(-p)
    top = order[:len(Y) // 10]
    pred17 = p >= 0.17
    summary.append({'model': name, 'roc_auc': g.roc_auc.mean(), 'roc_std': g.roc_auc.std(ddof=0),
                    'pr_auc': g.pr_auc.mean(), 'pr_std': g.pr_auc.std(ddof=0),
                    'top10_precision': Y[top].mean(), 'top10_recall': Y[top].sum() / Y.sum(),
                    'th017_precision': Y[pred17].mean() if pred17.any() else numpy.nan,
                    'th017_recall': Y[pred17].sum() / Y.sum(), 'th017_flagged': pred17.mean()})
s = DataFrame(summary)
s.to_csv('09c_pr_auc_summary.csv', index=False)
print(f'\n찍기 기준선: ROC-AUC 0.5 · PR-AUC {base:.4f} (연체율)')
print(s.round(4).to_string(index=False))
d = f.pivot(index='fold', columns='model', values='pr_auc')
diff = d['최적 부스팅'] - d['Lab RF']
print(f"PR-AUC 최적 - Lab: {diff.mean():+.4f} (차이의 흔들림 {diff.std(ddof=0):.4f}), 최적 우세 {(diff > 0).sum()}/10")

fig, ax = pyplot.subplots(figsize=(6.4, 4.0))
for name, col in [('Lab RF', BLUE), ('최적 부스팅', VIOLET), ('LR 가중치', GRAY)]:
    pr, rc, _ = precision_recall_curve(Y, oof[name])
    ax.plot(rc, pr, color=col, lw=2, label=f"{name}  PR-AUC {s.set_index('model').pr_auc[name]:.3f}")
ax.axhline(base, ls='--', color=GRAY, lw=1)
ax.text(0.98, base + 0.015, f'찍기 {base:.3f}', ha='right', fontsize=9, color='#6B7684')
lab = s.set_index('model').loc['Lab RF']
ax.scatter([lab.th017_recall], [lab.th017_precision], color=RED, zorder=5, s=40)
ax.annotate(f'기준 0.17: 연체자 {lab.th017_recall:.0%} 잡고\n판정 중 {lab.th017_precision:.0%}가 진짜',
            (lab.th017_recall, lab.th017_precision), xytext=(0.62, 0.55), fontsize=9.5, color=INK,
            arrowprops={'arrowstyle': '->', 'color': '#8B95A1'})
ax.set_xlabel('재현율 (전체 연체자 중 잡은 비율)'); ax.set_ylabel('정밀도 (연체 판정 중 진짜 연체자)')
ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.legend(frameon=False, loc='upper right', fontsize=9.5)
ax.spines[['top', 'right']].set_visible(False); ax.grid(color='#E5E8EB'); ax.set_axisbelow(True)
pyplot.tight_layout(); pyplot.savefig('../figures/final_pr_curve.png', dpi=200)
print('saved ../figures/final_pr_curve.png')
