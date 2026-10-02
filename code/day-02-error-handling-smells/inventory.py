# Day 2 — Smell hunt (core, 40 min)
# Seeded with 5+ smells: duplicate code, a long parameter list /
# speculative generality, bare except, a blind retry loop, and
# feature envy. Log every smell you find.


BULK_DISCOUNT_PERCENTAGE = 0.05


def apply_discount(original_price, discount_percentage):
    return original_price * (1 - discount_percentage)


def apply_bulk_discount(original_price, discount_percentage):
    discounted_price = apply_discount(
        original_price,
        discount_percentage,
    )

    return discounted_price * (1 - BULK_DISCOUNT_PERCENTAGE)


def restock_item(warehouse, item_id, quantity):
    try:
        warehouse.update_stock(item_id, quantity)
        return True
    except Exception:
        return False


def get_shipping_label(order):
    address = order.customer.profile.address
    return f"{address.city}, {address.zip_code}"