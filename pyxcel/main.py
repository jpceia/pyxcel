from pyxcel import *
from flask import Flask
import inspect
import json

import test

app = Flask(__name__)


@app.route("/query/<string:foo>/<string:args>")
def query(foo, args):
    assert foo in udf_dico
    args = json.loads(args)
    ret = json.dumps(udf_dico[foo](*args))
    return ret


@app.route("/udf_list/")
def list_udf():
    global udf_dico
    dico = dict()
    for name, foo in udf_dico.items():
        dico[name] = inspect.getargspec(foo).args
    return json.dumps(dico)


@app.route("/echo/<string:message>")
def echo(message):
    return message
