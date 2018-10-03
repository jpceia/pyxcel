import pyxcel as pyx
from flask import Flask
import inspect
import json

import test

app = Flask(__name__)


@app.route("/query/<string:foo>/<string:args>")
def query(foo, args):
    assert foo in pyx.udf_dico
    args = json.loads(args)
    ret = json.dumps(pyx.udf_dico[foo](*args))
    return ret


@app.route("/list_udf/")
def list_udf():
    dico = dict()
    for name, foo in pyx.udf_dico.items():
        detail = dict()
        detail['args'] = inspect.getargspec(foo).args
        if foo.__doc__ is not None:
            detail['doc'] = foo.__doc__
        dico[name] = detail
    return json.dumps(dico)


@app.route("/echo/<string:message>")
def echo(message):
    return message


if __name__ == "__main__":
    app.run(debug=True, use_reloader=False)