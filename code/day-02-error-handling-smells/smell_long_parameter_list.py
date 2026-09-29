def schedule_delivery(order_id, address, city, state,
        zip_code, carrier, weight_kg, is_fragile,
        requires_signature):
    ...

# 9 positional args - a caller has to read the
# function body to know what's required
schedule_delivery(4821, "12 Elm St", "Austin", "TX",
    "78701", "UPS", 3.4, True, False)
