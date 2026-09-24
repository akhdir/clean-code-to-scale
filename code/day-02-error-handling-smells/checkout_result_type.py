# Day 2 — Approach B: Result type
# A declined card isn't a bug — it's an everyday outcome, treated as
# data instead of an interrupt.


class ChargeResult:
    def __init__(self, ok, reason=None):
        self.ok = ok
        self.reason = reason


def charge_payment(payment_gateway, amount):
    try:
        payment_gateway.charge(amount)
        return ChargeResult(ok=True)
    except CardDeclined:
        return ChargeResult(ok=False, reason="declined")
    except GatewayTimeout:
        return ChargeResult(ok=False, reason="unavailable")


def process_checkout(order, customer, payment_gateway):
    total = calculate_total(order, customer)
    result = charge_payment(payment_gateway, total)
    if not result.ok:
        return {"status": "failed", "reason": result.reason}
    mark_order_paid(order, total)
    return {"status": "ok", "total": total}


# every outcome is visible in what's returned
outcome = process_checkout(order, customer, gateway)
