import pyxcel as pyx
from flask import Flask, request, jsonify
import os, sys
try:
    from StringIO import StringIO
except ImportError:
    from io import StringIO

import test

app = Flask(__name__)


class Capturing(list):
    def __enter__(self):
        self._stdout = sys.stdout
        sys.stdout = self._stringio = StringIO()
        return self
    def __exit__(self, *args):
        self.extend(self._stringio.getvalue().splitlines())
        del self._stringio    # free up some memory
        sys.stdout = self._stdout


@app.route("/import", methods=["POST"])
def import_files():
    paths = request.json
    path_list = []
    for path in paths:
        filedir, filename = os.path.split(path)
        if not os.path.isdir(filedir):
            msg = "{} is an invalid directory".format(filedir)
        if not file.endswith(".py"):
            msg = "{} is an invalid python file".format(filename)
        pyx.add_module(filedir, filename[:-3])


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
    s = request.json["command"]
    with Capturing() as output:
        exec(s)
    return jsonify(output)


if __name__ == "__main__":
    app.run(debug=True, use_reloader=False)
