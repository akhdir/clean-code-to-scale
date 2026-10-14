def get_order_confirmation(order_id, session):
    if session.wrote_recently(order_id):
        return db.primary().get_order(order_id)
    return db.read_replica().get_order(order_id)
