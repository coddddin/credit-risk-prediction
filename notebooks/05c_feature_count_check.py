"""Lab 3에서 남긴 변수 후보 두 개(8개 / 6개)를 Lab 5 최종 파이프라인(정리 + RF 잎 176명)으로 비교한다.

6개 = Lab 3 트리 중요도 상위 6개 (부양가족 수, 대출 보유를 뺌)
평가 = 층화 10-fold(seed 7), 같은 조각끼리 짝지어 비교
실행: python 05c_feature_count_check.py  → 05c_feature_count_scores.csv
"""
import numpy
from pandas import read_csv, DataFrame
from sklearn.base import clone
from sklearn.ensemble import RandomForestClassifier, ExtraTreesClassifier
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import roc_auc_score
from sklearn.model_selection import StratifiedKFold
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import FunctionTransformer, StandardScaler

seed = 7
names = ['target', 'revol', 'age', 'late30', 'debt', 'income', 'depend', 'defaulted', 'loans']
features = names[1:]
df = read_csv('../data/givemesomecredit.csv', names=names, header=0)
X = df.values[:, 1:9].astype(float)
Y = df.values[:, 0].astype(int)
iR, iA, iL, iD, iI, iP, iF, iO = range(8)


def clean(A):  # Lab 5 정리 단계 (05·09 노트북과 같음)
    A = A.copy()
    code = numpy.isin(A[:, iL], [96, 98]).astype(float)
    fill = numpy.isin(A[:, iI], [6638, 6746]).astype(float)
    A[code == 1, iL] = 0
    return numpy.column_stack([A, code, fill])


# Lab 3 트리 중요도 순서 (03 노트북과 같은 방법)
imp = ExtraTreesClassifier(n_estimators=100, random_state=seed, n_jobs=-1).fit(X, Y).feature_importances_
ranked = [features[i] for i in numpy.argsort(-imp)]
six = sorted(features.index(c) for c in ranked[:6])
dropped = ranked[6:]
print('중요도 순서:', ranked, '→ 6개에서 뺀 변수:', dropped)

keep6 = six + [8, 9]  # 정리 단계가 붙이는 표시 열 2개는 유지
rf = RandomForestClassifier(n_estimators=100, max_features=3, min_samples_leaf=176, random_state=seed, n_jobs=-1)
models = {
    'RF 8개': Pipeline([('f', FunctionTransformer(clean)), ('m', rf)]),
    'RF 6개': Pipeline([('f', FunctionTransformer(clean)), ('s', FunctionTransformer(lambda A: A[:, keep6])), ('m', rf)]),
}

# Lab 3 기준(로그+표준화 + 로지스틱 회귀)도 같은 조각으로 다시 잰다
Xlog = X.copy()
for c in ['revol', 'late30', 'debt', 'income']:
    Xlog[:, features.index(c)] = numpy.log1p(Xlog[:, features.index(c)])
lr8 = Pipeline([('s', StandardScaler()), ('m', LogisticRegression(solver='liblinear'))])

rows = []
for k, (tr, te) in enumerate(StratifiedKFold(10, shuffle=True, random_state=seed).split(X, Y)):
    r = {'fold': k + 1}
    for name, m in models.items():
        r[name] = roc_auc_score(Y[te], clone(m).fit(X[tr], Y[tr]).predict_proba(X[te])[:, 1])
    for name, cols in [('LR 8개', list(range(8))), ('LR 6개', six)]:
        m = clone(lr8).fit(Xlog[tr][:, cols], Y[tr])
        r[name] = roc_auc_score(Y[te], m.predict_proba(Xlog[te][:, cols])[:, 1])
    rows.append(r)
    print(r, flush=True)

s = DataFrame(rows)
s.to_csv('05c_feature_count_scores.csv', index=False)
for a, b in [('RF 8개', 'RF 6개'), ('LR 6개', 'LR 8개')]:
    d = s[a] - s[b]
    print(f'{a} {s[a].mean():.4f} vs {b} {s[b].mean():.4f} → 차이 {d.mean():+.4f} (차이의 흔들림 {d.std(ddof=0):.4f}), {a} 우세 {(d > 0).sum()}/10')
