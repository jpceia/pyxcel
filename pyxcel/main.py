import pyxcel as pyx
from flask import Flask, request, jsonify
import os, sys
try:
    from StringIO import StringIO
except ImportError:
    from io import StringIO

app = Flask(__name__)


@app.route("/import", methods=["POST"])
def import_files():
    query = request.json
    xldir = query['dir']
    file = query['file']
    try:
        pyx.import_module(xldir, file)
        return '1'
    except ValueError as e:
        print(e)
        return '0'


@app.route("/eval", methods=["POST"])
def eval():
    query = request.json
    foo_name = query["foo"]
    args = query.get("args", ())
    return jsonify(pyx.eval(foo_name, *args))


@app.route("/signatures")
def signatures():
    return jsonify(pyx.signatures())


@app.route("/echo/<string:message>")
def echo(message):
    return message


@app.route("/execute", methods=["POST"])
def execute():
    commands = request.json['command']
    return jsonify(pyx.execute(commands))


if __name__ == "__main__":
    app.run(debug=True, use_reloader=False)
