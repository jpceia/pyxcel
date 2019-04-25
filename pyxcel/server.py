import os, sys
import pyxcel as pyx
import argparse
from flask import Flask, request, jsonify
from functools import wraps

app = Flask(__name__)


def api_call(foo):
    @wraps(foo)
    def wrapper(*args):
        res = {}
        try:
            kwargs = request.json
            if kwargs is None:
                kwargs = {}
            res['result'] = foo(*args, **kwargs)
            res['success'] = True
        except Exception as err:
            print(err)
            res['success'] = False
            res['error'] = {
                'type': type(err).__name__,
                'code': getattr(err, 'code', 0),
                'message': getattr(err, 'description', '')
            }
        return jsonify(res)
    return wrapper


@app.route("/import", methods=["GET", "POST"])
@api_call
def import_files(**kwargs):
    return pyx.import_module(kwargs['dir'], kwargs['file'])


@app.route("/eval", methods=["GET", "POST"])
@api_call
def eval(**kwargs):
    foo_name = kwargs["foo"]
    args = kwargs.get("args", ())
    return pyx.eval(foo_name, *args)


@app.route("/signatures", methods=["GET", "POST"])
@api_call
def signatures(**kwargs):
    return pyx.signatures()


@app.route("/execute", methods=["GET", "POST"])
@api_call
def execute(**kwargs):
    return pyx.execute(kwargs["command"])


@app.route("/echo/<string:message>")
def echo(message):
    return message


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
