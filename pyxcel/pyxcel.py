udf_dico = dict()

def export(foo):
    udf_dico[foo.__name__] = foo
    return foo
