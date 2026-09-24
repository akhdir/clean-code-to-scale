# Day 10 — the punchline: PlaceOrder.run() is untouched in all three.
# Whoever constructs PlaceOrder decides which adapter to wire in;
# PlaceOrder itself never gets a vote.

# production
place_order = PlaceOrder(PostgresOrderRepository())

# tests — no network, no database
place_order = PlaceOrder(InMemoryOrderRepository())

# production, with an audit trail — wraps the real adapter
place_order = PlaceOrder(
    LoggingOrderRepository(PostgresOrderRepository(), logger)
)

# PlaceOrder.run() never changes. Not one line, in any of the three.
