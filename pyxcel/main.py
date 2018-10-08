import pyxcel as pyx
from flask import Flask, request
import inspect
import json

import test

app = Flask(__name__)


@app.route("/eval", methods=["POST"])
def eval():
    query = request.json
    foo_name = query["foo"]
    args = query.get("args", ())
    return json.dumps(pyx.eval(foo_name, *args))


@app.route("/udf_signatures")
def udf_signatures():
    return json.dumps(pyx.udf_signatures())


@app.route("/echo/<string:message>")
def echo(message):
    return message

if __name__ == "__main__":
    app.run(debug=True, use_reloader=False)