import pyxcel as pyx
import sklearn.linear_model
import sklearn.cluster
#import sklearn.ensemble
import numpy as np
import re


GLOBAL_VARS = dict()

MODEL_CATALOG = {
    'logisticregression': sklearn.linear_model.LogisticRegression,
    'linearregression': sklearn.linear_model.LinearRegression,
    'lassoregression': sklearn.linear_model.Lasso,
    'ridgeregression': sklearn.linear_model.Ridge,
    #'randomforestclassifier': sklearn.ensemble.RandomForestClassifier,
    #'extratreesclassifier': sklearn.ensemble.ExtraTreesClassifier,
    'kmeans': sklearn.cluster.KMeans,
}


def load_model(tag):
    tag = sanitize_tag(tag)
    if tag not in GLOBAL_VARS:
        raise ValueError
    return GLOBAL_VARS[tag]


def sanitize_tag(tag):
    return re.sub(r"\W+", "", tag.lower())


@pyx.export
def ML_Load(tag, model_type, params):
    """
    Loads a model

    :param tag string: tag for the new model
    :param model_type string: model type to use
    :param params: params to use
    """
    tag = sanitize_tag(tag)
    model_type = sanitize_tag(model_type)
    if model_type not in MODEL_CATALOG:
        raise ValueError("Model not available")
    kwargs = dict()
    for row in params:
        key = row[0]
        value = row[1]
        if key is not None:
            kwargs[key] = value
    model = MODEL_CATALOG[model_type](**kwargs)
    global GLOBAL_VARS
    GLOBAL_VARS[tag] = model
    return tag


@pyx.export
def ML_Fit(tag, X, y):
    """
    Loads a model

    :param tag string: tag for loaded model
    :param X: features
    :param y: labels
    """
    tag = sanitize_tag(tag)
    model = load_model(tag)
    X = np.asarray(X, dtype=np.float64)
    y = np.asarray(y, dtype=np.float64)
    if X.ndim == 1:
        X = X.reshape((-1, 1))
    if y.ndim > 1:
        y = y[:, 0]
    model.fit(X, y)
    return tag


@pyx.export
def ML_Predict(tag, X):
    """
    Predicts labels given features and a fitted model

    :param tag string: tag for loaded model
    :param X: features
    """
    tag = sanitize_tag(tag)
    model = load_model(tag)
    return model.predict(np.asarray(X, dtype=np.float64))


@pyx.export
def ML_Transform(tag, X):
    """
    transforms features given a fitted model

    :param tag string: tag for loaded model
    :param X: features
    """
    tag = sanitize_tag(tag)
    model = load_model(tag)
    return model.transform(np.asarray(X, dtype=np.float64))


@pyx.export
def ML_Coef(tag, coef):
    tag = sanitize_tag(tag)
    model = load_model(tag)
    return getattr(model, coef + "_")
