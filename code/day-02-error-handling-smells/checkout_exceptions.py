# Day 2 — Approach A: Exceptions
# CardDeclined and GatewayTimeout are exceptions the payment library
# already raises; PaymentDeclined and PaymentGatewayUnavailable are ours.


class PaymentDeclined(Exception):
    pass


class PaymentGatewayUnavailable(Exception):
    pass


def charge_payment(payment_gateway, amount):
    try:
        payment_gateway.charge(amount)
    except CardDeclined as e:
        raise PaymentDeclined(str(e)) from e
    except GatewayTimeout as e:
        raise PaymentGatewayUnavailable(str(e)) from e


def process_checkout(order, customer, payment_gateway):
    total = calculate_total(order, customer)
    charge_payment(payment_gateway, total)
    mark_order_paid(order, total)
    return {"status": "ok", "total": total}


# the caller is forced to decide — no silent path through
try:
    process_checkout(order, customer, gateway)
except PaymentDeclined:
    notify_customer(customer, "card declined")
except PaymentGatewayUnavailable:
    schedule_retry(order)
