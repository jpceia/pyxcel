import pyxcel as pyx
import sklearn.linear_model
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
}


def load_model(model):
    if model not in GLOBAL_VARS:
        raise ValueError
    return GLOBAL_VARS[model]


@pyx.export
def ML_Load(name, model_type, params):
    """
    Loads a model

    :param name: First argument to be added
    :param model_type: model type to use
    :param params: params to use
    """
    model_type = re.sub(r"\W+", "", model_type.lower())
    if model_type not in MODEL_CATALOG:
        raise ValueError
    kwargs = dict()
    for row in params:
        key = row[0]
        value = row[1]
        if key is not None:
            kwargs[key] = value
    model = MODEL_CATALOG[model_type](**kwargs)
    global GLOBAL_VARS
    GLOBAL_VARS[name] = model
    return name


@pyx.export
def ML_Fit(name, X, y):
    model = load_model(name)
    X = np.array(X)
    y = np.array(y)
    if y.ndim > 1:   # really necessary?
        y = y[:, 0]  #
    model.fit(X, y)
    return name


@pyx.export
def ML_Predict(name, X):
    model = load_model(name)
    return model.predict(np.array(X))


@pyx.export
def ML_Transform(name, X):
    model = load_model(name)
    return model.transform(np.array(X))


@pyx.export
def ML_Coef(name, coef):
    model = load_model(name)
    return getattr(model, coef + "_")
