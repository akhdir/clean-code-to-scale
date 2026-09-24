# Day 1 — Naming, Functions & Formatting
# Live demo, "after" — seven functions, one job each.

TAX_RATE = 0.08
LOYALTY_DISCOUNT_THRESHOLD = 1000
LOYALTY_DISCOUNT_RATE = 0.10


def calculate_subtotal(items):
    return sum(item["price"] * item["qty"] for item in items)


def apply_loyalty_discount(subtotal, loyalty_points):
    if loyalty_points > LOYALTY_DISCOUNT_THRESHOLD:
        return subtotal * (1 - LOYALTY_DISCOUNT_RATE)
    return subtotal


def calculate_total(order, customer):
    subtotal = calculate_subtotal(order["items"])
    discounted = apply_loyalty_discount(subtotal, customer["loyalty"])
    return discounted + (discounted * TAX_RATE)


def mark_order_paid(order, total):
    order["status"] = "paid"
    order["total"] = total
    db.save(order)


def notify_by_email(customer, total):
    email.send(customer["email"], "Order paid: $%.2f" % total)


def notify_by_sms(customer, total):
    sms.send(customer["phone"], "Order paid: $%.2f" % total)


def process_checkout(order, customer, payment_gateway):
    total = calculate_total(order, customer)
    # real error handling for a failed charge is Day 2's territory —
    # for now we let it fail loudly instead of swallowing it
    if not payment_gateway.charge(total):
        return {"status": "failed"}
    mark_order_paid(order, total)
    return {"status": "ok", "total": total}


# the caller decides how to notify — no more hidden flags
result = process_checkout(order, customer, gateway)
notify_by_email(customer, result["total"])
