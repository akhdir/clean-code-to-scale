# Day 2 — Smell hunt, stretch (~15 min)
# One error-handling smell per function: error codes as strings/None,
# error handling mixed with business logic, except-with-pass, and
# retrying blindly.


def cancel_order(order, warehouse):
    try:
        if order.status == "shipped":
            return None
        order.status = "cancelled"
        warehouse.release_reserved_stock(order.item_id, order.qty)
        payment_gateway.refund(order.payment_id, order.total)
        send_cancellation_email(order.customer.email)
        log.info(f"order {order.id} cancelled")
        return "OK"
    except Exception:
        return "ERROR"


def sync_inventory(warehouse, feed):
    for item in feed:
        try:
            warehouse.update_stock(item["id"], item["qty"])
        except Exception:
            pass


def notify_customer(customer, message):
    attempts = 0
    while attempts < 100:
        try:
            email_service.send(customer.email, message)
            return
        except Exception:
            attempts += 1
