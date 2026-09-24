# Day 10 — Adapter 2: tests. No network, no database, no setup.


class InMemoryOrderRepository(OrderRepository):
    def __init__(self):
        self.orders = {}

    def save(self, order):
        self.orders[order.id] = order
