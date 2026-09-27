def get_order_confirmation(order_id):
    return db.read_replica().get_order(order_id)
