from pyxcel import export
import random
import numpy as np

@export
def add(y, x):
    """
    function to add two numbers
    Returns x + y

    Long description
    :param double y: First argument to be added
    :param double x: Second argument to be added
    :return double: returns x + y
    """
    return x + y

@export
def subtract(x, y):
    return x - y

@export
def reshape(x, n, m):
    res = np.array(x, dtype='object').reshape(n, m).tolist()
    return res

@export
def pyrandom():
    return random.random()