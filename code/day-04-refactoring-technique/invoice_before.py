# Day 4 — Live demo "before": all three smells, tangled together.


def _get_item_count(order):
    return len(order.items)


def generate_invoice_summary(order):
    total = 0
    for item in order.items:
        total += item.price * item.qty

    if order.discount_type == "loyalty":
        total *= 0.9
    elif order.discount_type == "seasonal":
        total *= 0.85
    elif order.discount_type == "employee":
        total *= 0.7

    count = _get_item_count(order)
    return f"{count} items, total ${total:.2f}"
