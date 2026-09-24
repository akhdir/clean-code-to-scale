# Day 4 — Live demo "after": all three steps (Extract Method, Inline,
# Replace Conditional with Polymorphism), in one piece.
# DiscountStrategy, LoyaltyDiscount, etc. — from Day 3, unchanged.

from order_discounts import DiscountStrategy


def calculate_subtotal(items):
    return sum(item.price * item.qty for item in items)


def generate_invoice_summary(order, discount):
    total = calculate_subtotal(order.items)
    total = discount.apply(total)
    return f"{len(order.items)} items, total ${total:.2f}"
