import pyxcel as pyx
import argparse
from flask import Flask, request, jsonify
import os, sys

app = Flask(__name__)


@app.route("/import", methods=["GET", "POST"])
def import_files():
    query = request.json
    xldir = query['dir']
    file = query['file']
    try:
        pyx.import_module(xldir, file)
        return "True"
    except ValueError as e:
        print(e)
        return "False"  # use NaN instead?


@app.route("/eval", methods=["GET", "POST"])
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


@app.route("/execute", methods=["GET", "POST"])
def execute():
    commands = request.json['command']
    return jsonify(pyx.execute(commands))


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument('--host',   help='Host', type=str, default="0.0.0.0")
    parser.add_argument('--port',   help='Port', type=int, default=5555)
    parser.add_argument('-m', '--modules', help='Module(s)', type=str, nargs='*', default=[])
    args = parser.parse_args()
    root_folder = os.getcwd()
    for module in args.modules:
        pyx.import_module(root_folder, module)
    app.run(host=args.host, port=args.port, use_reloader=True, debug=True)
