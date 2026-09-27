# carts lives in this process's memory only
carts = {}

def add_to_cart(session_id, item):
    carts.setdefault(session_id, []).append(item)

def get_cart(session_id):
    return carts.get(session_id, [])
