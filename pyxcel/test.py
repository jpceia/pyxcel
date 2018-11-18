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
    x = np.array(x).flatten().tolist()
    res = []
    for i in range(n):
        row = []
        for j in range(m):
            row.append(x[i * n + j])
        res.append(row)
    return res

@export
def pyrandom():
    return random.random()