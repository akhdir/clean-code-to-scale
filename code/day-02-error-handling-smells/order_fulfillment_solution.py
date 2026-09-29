# Day 2 — Smell hunt, stretch (~15 min) — one valid answer shape
#
# Fixes, matched to the answer key:
#   - error codes as strings/None:      cancel_order returns one typed
#                                        CancelResult instead of None / "OK" / "ERROR"
#   - error handling mixed with logic:  the "already shipped" business
#                                        check moved out of the try — it's
#                                        not a failure, so it shouldn't be
#                                        caught like one
#   - except with pass:                 sync_inventory logs each failure
#                                        and returns the list of failed items
#                                        instead of dropping them silently
#   - retrying blindly:                 notify_customer retries 3 times
#                                        with backoff instead of 100 times
#                                        with none

import time
from dataclasses import dataclass


@dataclass
class CancelResult:
    ok: bool
    reason: str = None


def cancel_order(order, warehouse):
    if order.status == "shipped":
        return CancelResult(ok=False, reason="already_shipped")

    order.status = "cancelled"

    try:
        warehouse.release_reserved_stock(order.item_id, order.qty)
        payment_gateway.refund(order.payment_id, order.total)
    except Exception as e:
        log.error(f"cancel_order failed for order {order.id}: {e}")
        return CancelResult(ok=False, reason="refund_failed")

    send_cancellation_email(order.customer.email)
    log.info(f"order {order.id} cancelled")
    return CancelResult(ok=True)


def sync_inventory(warehouse, feed):
    failed_items = []
    for item in feed:
        try:
            warehouse.update_stock(item["id"], item["qty"])
        except Exception as e:
            log.error(f"failed to sync item {item['id']}: {e}")
            failed_items.append(item["id"])
    return failed_items


MAX_NOTIFY_ATTEMPTS = 3
NOTIFY_BACKOFF_SECONDS = 1


def notify_customer(customer, message):
    for attempt in range(1, MAX_NOTIFY_ATTEMPTS + 1):
        try:
            email_service.send(customer.email, message)
            return True
        except Exception as e:
            if attempt == MAX_NOTIFY_ATTEMPTS:
                log.error(f"notify_customer failed for {customer.email}: {e}")
                return False
            time.sleep(NOTIFY_BACKOFF_SECONDS * attempt)
