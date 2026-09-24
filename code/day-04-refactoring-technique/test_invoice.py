# Day 4 — the safety net every red-green checkpoint in the demo refers to.


def test_generate_invoice_summary():
    order = Order(
        items=[Item(price=10, qty=1), Item(price=5, qty=2)],
        discount_type="loyalty",
    )
    assert generate_invoice_summary(order) == "2 items, total $18.00"


# after Step 3 (Replace Conditional with Polymorphism), the call shape
# changes but the assertion doesn't:
#
# def test_generate_invoice_summary():
#     order = Order(items=[Item(price=10, qty=1), Item(price=5, qty=2)])
#     assert generate_invoice_summary(order, LoyaltyDiscount()) == "2 items, total $18.00"
