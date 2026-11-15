"""Lab 5 후보별 튜닝 비교 (중첩 교차검증).

바깥: 층화 10-fold(seed 7, 다른 노트북과 같은 조각) → 이 점수로 후보를 비교한다.
안쪽: 각 바깥 학습 구간 안에서 층화 3-fold GridSearchCV로 설정을 고른다 (바깥 시험 조각은 설정 선택에 쓰지 않음).
후보: Lab 4·5에서 성능이 좋았거나 튜닝 여지가 분명한 4개
  LR       (정리 + 로그 + 표준화, 가중치)  C
  KNN      (정리 + 로그 + 표준화)          n_neighbors
  AdaBoost (정리, 깊이 1 트리)             n_estimators · learning_rate
  RF       (정리, 트리 100 · 변수 3)       min_samples_leaf
실행: python 05d_candidate_tuning.py [모델 이름 ...]   (SMOKE=1 이면 8,000건으로 빠르게)
결과: 05d_candidate_scores.csv (조각별 기본·튜닝 AUC), 05d_candidate_params.csv (조각별 고른 설정)
"""
import os, sys, time
import numpy
from pandas import read_csv, DataFrame
from sklearn.base import clone
from sklearn.ensemble import RandomForestClassifier, AdaBoostClassifier
from sklearn.linear_model import LogisticRegression
from sklearn.neighbors import KNeighborsClassifier
from sklearn.metrics import roc_auc_score
from sklearn.model_selection import StratifiedKFold, GridSearchCV, train_test_split
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import FunctionTransformer, StandardScaler

seed = 7
SMOKE = os.environ.get('SMOKE') == '1'
names = ['target', 'revol', 'age', 'late30', 'debt', 'income', 'depend', 'defaulted', 'loans']
df = read_csv('../data/givemesomecredit.csv', names=names, header=0)
X = df.values[:, 1:9].astype(float)
Y = df.values[:, 0].astype(int)
if SMOKE:
    X, _, Y, _ = train_test_split(X, Y, train_size=8000, random_state=seed, stratify=Y)
iR, iA, iL, iD, iI, iP, iF, iO = range(8)


def clean(A):  # Lab 5 정리 단계 (05·09 노트북과 같음)
    A = A.copy()
    code = numpy.isin(A[:, iL], [96, 98]).astype(float)
    fill = numpy.isin(A[:, iI], [6638, 6746]).astype(float)
    A[code == 1, iL] = 0
    return numpy.column_stack([A, code, fill])


def log_skewed(A):  # Lab 2 결론: 치우친 4개 열에 log1p
    A = A.copy()
    for i in (iR, iL, iD, iI):
        A[:, i] = numpy.log1p(A[:, i])
    return A


def scaled(model):
    return Pipeline([('c', FunctionTransformer(clean)), ('l', FunctionTransformer(log_skewed)),
                     ('s', StandardScaler()), ('m', model)])


def raw(model):
    return Pipeline([('c', FunctionTransformer(clean)), ('m', model)])


# (기본 설정 모델, 탐색할 설정, 왜 이 설정인지)
CANDIDATES = {
    'LR': (scaled(LogisticRegression(solver='liblinear', class_weight='balanced')),
           {'m__C': [0.001, 0.01, 0.1, 1, 10]}, '규제 세기'),
    'KNN': (scaled(KNeighborsClassifier(n_neighbors=5, n_jobs=-1)),
            {'m__n_neighbors': [5, 25, 100, 200]}, '이웃 5명이면 확률이 6단계뿐'),
    'AdaBoost': (raw(AdaBoostClassifier(n_estimators=30, random_state=seed)),
                 {'m__n_estimators': [30, 100, 300], 'm__learning_rate': [0.1, 0.5, 1.0]}, '트리 수 · 학습률'),
    'RF': (raw(RandomForestClassifier(n_estimators=100, max_features=3, random_state=seed, n_jobs=-1)),
           {'m__min_samples_leaf': [1, 10, 50, 100, 200]}, '잎 1명 = 학습 데이터 외우기'),
}
# 1차 탐색에서 모든 조각이 범위 끝(AdaBoost 300·1.0, KNN 200)을 골라 범위를 넓힌 2차 탐색 (직전 최고값 포함)
CANDIDATES['AdaBoost+'] = (CANDIDATES['AdaBoost'][0], {'m__n_estimators': [300, 600, 1000], 'm__learning_rate': [1.0]},
                           '1차 최고(300·1.0)가 범위 끝 → 트리 수 확장')
CANDIDATES['KNN+'] = (CANDIDATES['KNN'][0], {'m__n_neighbors': [200, 500, 1000]}, '1차 최고(200)가 범위 끝 → 이웃 수 확장')
wanted = sys.argv[1:] or list(CANDIDATES)
outer = list(StratifiedKFold(10, shuffle=True, random_state=seed).split(X, Y))
inner = StratifiedKFold(3, shuffle=True, random_state=seed)

score_path, param_path = '05d_candidate_scores.csv', '05d_candidate_params.csv'
if SMOKE:
    score_path, param_path = '05d_smoke_scores.csv', '05d_smoke_params.csv'
scores = read_csv(score_path).to_dict('records') if os.path.exists(score_path) and not SMOKE else []
params = read_csv(param_path).to_dict('records') if os.path.exists(param_path) and not SMOKE else []
scores = [r for r in scores if r['model'] not in wanted]
params = [r for r in params if r['model'] not in wanted]

for name in wanted:
    base, grid, why = CANDIDATES[name]
    t_model = time.time()
    for k, (tr, te) in enumerate(outer):
        t0 = time.time()
        b = clone(base).fit(X[tr], Y[tr])
        base_auc = roc_auc_score(Y[te], b.predict_proba(X[te])[:, 1])
        search = GridSearchCV(clone(base), grid, scoring='roc_auc', cv=inner, refit=True,
                              n_jobs=1 if name in ('RF', 'KNN') else -1)
        search.fit(X[tr], Y[tr])
        t1 = time.time()
        tuned_auc = roc_auc_score(Y[te], search.predict_proba(X[te])[:, 1])
        pred_sec = time.time() - t1
        scores.append({'model': name, 'fold': k + 1, 'base_auc': base_auc, 'tuned_auc': tuned_auc,
                       'refit_sec': search.refit_time_, 'pred_sec': pred_sec})
        params.append({'model': name, 'fold': k + 1, 'why': why,
                       **{p.replace('m__', ''): v for p, v in search.best_params_.items()}})
        print(f'{name} fold {k + 1:>2}: 기본 {base_auc:.4f} → 튜닝 {tuned_auc:.4f} {search.best_params_} '
              f'({time.time() - t0:.0f}초)', flush=True)
        DataFrame(scores).to_csv(score_path, index=False)
        DataFrame(params).to_csv(param_path, index=False)
    print(f'{name} 끝 ({(time.time() - t_model) / 60:.1f}분)', flush=True)

s = DataFrame(scores)
print('\n후보별 바깥 10-fold AUC (평균 ± 표준편차)')
for name, g in s.groupby('model', sort=False):
    print(f'  {name:9s} 기본 {g.base_auc.mean():.4f} → 튜닝 {g.tuned_auc.mean():.4f} (± {g.tuned_auc.std(ddof=0):.4f})')
if {'RF', 'AdaBoost'} <= set(s.model):
    rf = s[s.model == 'RF'].sort_values('fold').tuned_auc.values
    ada = s[s.model == 'AdaBoost'].sort_values('fold').tuned_auc.values
    d = rf - ada
    print(f'RF - AdaBoost: 평균 {d.mean():+.4f} (차이의 흔들림 {d.std(ddof=0):.4f}), RF 우세 {(d > 0).sum()}/10')
