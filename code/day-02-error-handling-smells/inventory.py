# Day 2 — Smell hunt (core, 40 min)
# Seeded with 5+ smells: duplicate code, a long parameter list /
# speculative generality, bare except, a blind retry loop, and
# feature envy. Log every smell you find.


def apply_discount(price, pct, code=None, is_admin=False, force=False, legacy_mode=False):
    if legacy_mode:
        return price
    if is_admin or force:
        return price * (1 - pct)
    try:
        return price * (1 - pct)
    except:
        return price


def apply_bulk_discount(price, pct, code=None, is_admin=False, force=False, legacy_mode=False):
    if legacy_mode:
        return price
    if is_admin or force:
        return price * (1 - pct) * 0.95
    try:
        return price * (1 - pct) * 0.95
    except:
        return price


def restock_item(warehouse, item_id, qty):
    for attempt in range(10):
        try:
            warehouse.update_stock(item_id, qty)
            return True
        except Exception:
            continue
    return False


def get_shipping_label(order):
    address = order.customer.profile.address
    return f"{address.city}, {address.zip_code}"
