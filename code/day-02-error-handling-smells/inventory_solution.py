# Day 2 — Smell hunt (core, 40 min) — one valid answer shape
#
# Fixes, matched to the answer key:
#   - duplicate code:      both discount functions now share _discounted_price
#   - long parameter list / speculative generality:
#                           `code` deleted (never used); `is_admin`/`force`
#                           deleted (their only effect was choosing whether to
#                           swallow errors, which the fix below removes)
#   - bare except:          gone — a bad `pct` should raise, not silently
#                           return the original price
#   - blind retry:          restock_item now retries 3 times with backoff,
#                           catches a specific error, and logs the failure
#   - feature envy:         get_shipping_label asks the order for its own
#                           shipping summary instead of reaching through
#                           order.customer.profile.address itself

import time

DISCOUNT_MULTIPLIER = 1.0
BULK_DISCOUNT_MULTIPLIER = 0.95

MAX_RESTOCK_ATTEMPTS = 3
RESTOCK_BACKOFF_SECONDS = 0.5


def _discounted_price(price, pct, multiplier=DISCOUNT_MULTIPLIER):
    return price * (1 - pct) * multiplier


def apply_discount(price, pct, legacy_mode=False):
    if legacy_mode:
        return price
    return _discounted_price(price, pct)


def apply_bulk_discount(price, pct, legacy_mode=False):
    if legacy_mode:
        return price
    return _discounted_price(price, pct, BULK_DISCOUNT_MULTIPLIER)


def restock_item(warehouse, item_id, qty):
    for attempt in range(1, MAX_RESTOCK_ATTEMPTS + 1):
        try:
            warehouse.update_stock(item_id, qty)
            return True
        except WarehouseUnavailable:
            if attempt == MAX_RESTOCK_ATTEMPTS:
                log.error(f"restock failed for item {item_id} after {MAX_RESTOCK_ATTEMPTS} attempts")
                return False
            time.sleep(RESTOCK_BACKOFF_SECONDS * attempt)
    return False


def get_shipping_label(order):
    # Order (or Customer) owns its own shipping summary — the caller
    # shouldn't need to know that it lives three attributes deep.
    return order.shipping_summary()
