from docstring_parser import parse
from collections import OrderedDict
from importlib import reload
import pandas as pd
import numpy as np
import inspect
import os, sys

UDF_DICO = dict()
VBA_NAMES = [
    "function",
    "sub",
    "next",
]


class DocString:
    def __init__(self, foo):
        doc = parse(getattr(foo, "__doc__", ""))
        self.short_desc = doc.short_description
        self.long_desc = doc.long_description
        if doc.returns:
            self.return_type = doc.returns.type_name
        else:
            self.return_type = None
        argDico = dict()
        for p in doc.params:
            arg_row = dict()
            if p.description:
                arg_row['desc'] = p.description
            if p.type_name:
                arg_row['type'] = p.type_name
            argDico[p.arg_name] = arg_row

        self.args = []
        for arg in inspect.getargspec(foo).args:
            arg_row = dict()
            arg_row['name'] = arg
            if arg in argDico:
                arg_row.update(argDico[arg])
            self.args.append(arg_row)


class Capturing(list):
    def __enter__(self):
        self._stdout = sys.stdout
        sys.stdout = self._stringio = StringIO()
        return self
    def __exit__(self, *args):
        self.extend(self._stringio.getvalue().splitlines())
        del self._stringio    # free up some memory
        sys.stdout = self._stdout


def export(foo):
    foo_name = foo.__name__
    #assert foo_name not in udf_dico
    assert foo_name not in VBA_NAMES
    UDF_DICO[foo_name] = foo
    return foo

def import_module(root, file):
    if not os.path.isabs(file):
        file = root + "\\" + file
    if not os.path.exists(file):
        raise ValueError("{} doesn't exist".format(file))
    folder, file = os.path.split(file)
    mod, e = os.path.splitext(file)
    if e != ".py":
        raise ValueError("{} has an invalid file extension".format(file))
    sys.path.insert(0, folder)
    m = __import__(mod)
    reload(m)
    sys.path.pop(0)


def signatures():
    dico = dict()
    for name, foo in UDF_DICO.items():
        doc = DocString(foo)
        detail = dict()
        if doc.short_desc:
            detail['doc'] = doc.short_desc
        # detail['category'] = None
        if doc.return_type:
            detail['type'] = doc.return_type
        detail['args'] = doc.args
        dico[name] = detail
    return dico


def eval(foo_name, *args):
    assert foo_name in UDF_DICO
    foo = UDF_DICO[foo_name]
    res = foo(*args)

    if isinstance(res, pd.Series):
        res = np.column_stack((res.index, res)).to_list()
    elif isinstance(res, pd.DataFrame):
        res = np.row_stack((
            np.concatenate(([res.index.name], res.columns)),
            np.column_stack((res.index, res)))).to_list()
    elif isinstance(res, np.ndarray):
        res = res.tolist()

    return res


def execute(commands):
    s = request.json["command"]
    with Capturing() as output:
        exec(s)
    return output