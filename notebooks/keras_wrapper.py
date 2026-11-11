"""Lab 6의 keras.wrappers.scikit_learn.KerasClassifier 대체.

최신 Keras 3에서는 원래 래퍼가 삭제되었고, scikeras는 scikit-learn 1.9에서
분류기로 인식되지 않아 AUC 계산이 깨진다. 그래서 같은 역할을 하는 최소 래퍼를 직접 만든다.
cross_val_score, GridSearchCV, Pipeline 안에서 그대로 쓸 수 있다.
"""
import numpy as np
from sklearn.base import BaseEstimator, ClassifierMixin


class KerasClassifier(ClassifierMixin, BaseEstimator):
    def __init__(self, build_fn=None, epochs=10, batch_size=512, verbose=0,
                 validation_split=0.0, callbacks=None, class_weight=None, **model_params):
        self.build_fn = build_fn
        self.epochs = epochs
        self.batch_size = batch_size
        self.verbose = verbose
        self.validation_split = validation_split
        self.callbacks = callbacks
        self.class_weight = class_weight
        self.model_params = model_params

    # GridSearchCV가 build_fn 인자(optimizer, init 등)도 탐색할 수 있게 한다.
    def get_params(self, deep=True):
        params = super().get_params(deep=deep)
        params.pop("model_params", None)
        params.update(self.model_params)
        return params

    def set_params(self, **params):
        own = set(super().get_params(deep=False)) - {"model_params"}
        for key, value in params.items():
            if key in own:
                setattr(self, key, value)
            else:
                self.model_params[key] = value
        return self

    def fit(self, X, y):
        X = np.asarray(X, dtype="float32")
        self.classes_ = np.unique(y)
        self.model_ = self.build_fn(**self.model_params)
        self.history_ = self.model_.fit(
            X, np.asarray(y, dtype="float32"), epochs=self.epochs, batch_size=self.batch_size,
            verbose=self.verbose, validation_split=self.validation_split,
            # callbacks에 함수를 주면 학습할 때마다 새로 만든다 (GridSearchCV가 모델을 복사해도 안전)
            callbacks=self.callbacks() if callable(self.callbacks) else self.callbacks,
            class_weight=self.class_weight)
        return self

    def predict_proba(self, X):
        p = self.model_.predict(np.asarray(X, dtype="float32"), batch_size=4096, verbose=0).ravel()
        return np.column_stack([1 - p, p])

    def predict(self, X):
        return self.classes_[(self.predict_proba(X)[:, 1] >= 0.5).astype(int)]
