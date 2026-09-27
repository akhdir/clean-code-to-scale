def get_shipping_estimate(order):
    rate = shipping_rate_api.get_rate(order.destination, order.weight)
    return rate

def checkout(order):
    estimate = get_shipping_estimate(order)
    order.shipping_cost = estimate
    return finalize_order(order)
