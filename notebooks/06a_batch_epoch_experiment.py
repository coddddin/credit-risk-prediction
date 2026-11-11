"""Lab 6 실험 A — 배치 크기 × 에폭 수에 따른 AUC와 학습 시간 (15만 건).

Lab 6의 첫 MLP(은닉층 12-8, relu/sigmoid, adam, binary_crossentropy)를 그대로 쓰고,
Lab 6 'Manual Verification Dataset' 방식(67% 학습 / 33% 검증, seed 7)으로 잰다.
Lab 2·4의 결론대로 입력은 로그+표준화(학습 데이터로만 계산), 연체자 가중치(balanced)를 준다.

배치 크기마다 150 에폭을 한 번 학습하면서, 매 에폭이 끝날 때 검증 AUC·손실과
그때까지의 순수 학습 시간(평가 시간 제외)을 CSV에 한 줄씩 기록한다.
→ 끊겨도 끝난 배치는 남고, 다시 실행하면 끝나지 않은 배치부터 이어서 한다.

실행: ..\\.venv\\Scripts\\python.exe 06a_batch_epoch_experiment.py
결과: 06a_batch_epoch_results.csv  (노트북 06에서 표·그림으로 정리)
"""
import os
import time

os.environ["TF_CPP_MIN_LOG_LEVEL"] = "3"
import numpy
from pandas import read_csv
from sklearn.metrics import roc_auc_score, log_loss
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler
import keras
from keras import layers

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "06a_batch_epoch_results.csv")
BATCHES = [2048, 512, 128, 32, 10]      # 빠른 것부터
EPOCHS = 150                             # Lab 6 예제 값

names = ['target', 'revol', 'age', 'late30', 'debt', 'income', 'depend', 'defaulted', 'loans']
features = names[1:]
df = read_csv(os.path.join(HERE, '..', 'data', 'givemesomecredit.csv'), names=names, header=0)
X = df.values[:, 1:9].astype(float)
Y = df.values[:, 0].astype(int)
for c in ['revol', 'late30', 'debt', 'income']:
    X[:, features.index(c)] = numpy.log1p(X[:, features.index(c)])
X_train, X_test, y_train, y_test = train_test_split(X, Y, test_size=0.33, random_state=7, stratify=Y)
scaler = StandardScaler().fit(X_train)                       # 학습 데이터로만 맞춤 (누수 방지)
X_train = scaler.transform(X_train).astype('float32')
X_test = scaler.transform(X_test).astype('float32')
class_weight = {0: 1.0, 1: float((y_train == 0).sum() / (y_train == 1).sum())}


def create_model():
    model = keras.Sequential([
        layers.Input((8,)),
        layers.Dense(12, activation='relu'),
        layers.Dense(8, activation='relu'),
        layers.Dense(1, activation='sigmoid'),
    ])
    model.compile(loss='binary_crossentropy', optimizer='adam')
    return model


class Recorder(keras.callbacks.Callback):
    """에폭마다 순수 학습 시간(검증 계산 제외)을 누적하고, 검증 AUC·손실을 CSV에 기록."""

    def __init__(self, batch, f):
        super().__init__()
        self.batch, self.f, self.train_sec = batch, f, 0.0

    def on_epoch_begin(self, epoch, logs=None):
        self.t0 = time.perf_counter()

    def on_epoch_end(self, epoch, logs=None):
        self.train_sec += time.perf_counter() - self.t0
        p = self.model.predict(X_test, batch_size=8192, verbose=0).ravel()
        auc = roc_auc_score(y_test, p)
        vloss = log_loss(y_test, numpy.clip(p, 1e-7, 1 - 1e-7))
        self.f.write(f"{self.batch},{epoch + 1},{self.train_sec:.3f},{logs['loss']:.5f},{vloss:.5f},{auc:.5f}\n")
        self.f.flush()
        if (epoch + 1) in (1, 5, 10, 25, 50, 100, 150):
            print(f"batch {self.batch:>5} · epoch {epoch + 1:>3} · AUC {auc:.4f} · 학습 {self.train_sec:7.1f}초", flush=True)


done = set()
if os.path.exists(OUT):
    rows = read_csv(OUT)
    done = {b for b, g in rows.groupby('batch') if g['epoch'].max() >= EPOCHS}
    rows[rows['batch'].isin(done)].to_csv(OUT, index=False)   # 끝나지 않은 배치 기록은 지우고 다시
else:
    with open(OUT, 'w') as f:
        f.write("batch,epoch,train_sec,train_loss,val_loss,val_auc\n")

print(f"학습 {len(y_train):,}건 / 검증 {len(y_test):,}건 · 이미 끝난 배치: {sorted(done)}", flush=True)
for b in BATCHES:
    if b in done:
        continue
    keras.utils.set_random_seed(7)
    model = create_model()
    with open(OUT, 'a') as f:
        model.fit(X_train, y_train, epochs=EPOCHS, batch_size=b, verbose=0,
                  class_weight=class_weight, callbacks=[Recorder(b, f)])
print("끝", flush=True)
