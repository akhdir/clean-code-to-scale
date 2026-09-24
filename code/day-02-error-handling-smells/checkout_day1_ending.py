# Day 2 — where Day 1 left off.
# process_checkout collapses every possible payment failure into one
# ambiguous outcome. A declined card, a network timeout, and a
# misconfigured gateway all look identical to the caller.


def process_checkout(order, customer, payment_gateway):
    total = calculate_total(order, customer)
    if not payment_gateway.charge(total):
        return {"status": "failed"}
    mark_order_paid(order, total)
    return {"status": "ok", "total": total}
