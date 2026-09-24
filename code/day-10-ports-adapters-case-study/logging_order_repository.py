# Day 10 — Adapter 3: production + audit. Wraps another adapter,
# doesn't replace it — composition, not modification.


class LoggingOrderRepository(OrderRepository):
    def __init__(self, wrapped, logger):
        self.wrapped = wrapped
        self.logger = logger

    def save(self, order):
        self.logger.info(f"saving {order.id}")
        self.wrapped.save(order)
