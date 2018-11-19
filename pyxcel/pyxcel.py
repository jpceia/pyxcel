from docstring_parser import parse
from collections import OrderedDict
import inspect
import os, sys

udf_dico = dict()
VBA_NAMES = [
    "function",
    "sub",
    "next",
]

def export(foo):
    foo_name = foo.__name__
    #assert foo_name not in udf_dico
    assert foo_name not in VBA_NAMES
    udf_dico[foo.__name__] = foo
    return foo


def import_module(path):
    if not os.path.exists(path):
        raise "{} doesn't exist".format(path)
    folder, file = os.path.split(path)
    mod, e = os.path.splitext(file)
    print(folder, file, mod, e)
    if e != ".py":
        raise "invalid file extension - needs to be .py"
    sys.path.insert(0, folder)
    __import__(mod)
    del sys.path[0]


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

def signatures():
    dico = dict()
    for name, foo in udf_dico.items():
        doc = DocString(foo)
        detail = dict()
        if doc.short_desc:
            detail['doc'] = doc.short_desc
        # detail['class'] = None
        if doc.return_type:
            detail['type'] = doc.return_type
        detail['args'] = doc.args
        dico[name] = detail
    return dico


def eval(foo_name, *args):
    assert foo_name in udf_dico
    foo = udf_dico[foo_name]
    return foo(*args)
