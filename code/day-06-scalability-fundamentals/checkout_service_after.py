def add_to_cart(session_id, item):
    redis.rpush(f"cart:{session_id}", item)

def get_cart(session_id):
    return redis.lrange(f"cart:{session_id}", 0, -1)
