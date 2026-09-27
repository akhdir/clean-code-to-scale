def add_one(x):
    return increment(x)

def increment(x):
    return plus_one(x)

def plus_one(x):
    return x + 1

# three functions, three stack frames, to compute x + 1 -
# extraction that adds indirection but no real abstraction
