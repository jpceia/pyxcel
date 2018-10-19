import inspect

udf_dico = dict()
VBA_NAMES = [
    "function",
    "sub",
    "next",
]

def export(foo):
    foo_name = foo.__name__
    assert foo_name not in udf_dico
    assert foo_name not in VBA_NAMES
    udf_dico[foo.__name__] = foo
    return foo



def udf_signatures():
    dico = dict()
    for name, foo in udf_dico.items():
        detail = dict()
        detail['args'] = inspect.getargspec(foo).args
        if foo.__doc__ is not None:
            detail['doc'] = foo.__doc__
        dico[name] = detail
    return dico


def eval(foo_name, *args):
    assert foo_name in udf_dico
    foo = udf_dico[foo_name]
    return foo(*args)