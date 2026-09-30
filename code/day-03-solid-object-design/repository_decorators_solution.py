# Day 3 — Decorator kata: answer key.
# OrderProcessor doesn't change at all — it still just calls
# repository.save(order_id, total). That's only possible because
# process() already receives repository instead of constructing it;
# Decorator is DIP's seam, used a second time.


class PostgresRepository:
    def save(self, order_id, total):
        db.execute(
            "INSERT INTO orders (id, total) VALUES (%s, %s)",
            order_id, total,
        )


class LoggingRepository:
    def __init__(self, repository):
        self.repository = repository

    def save(self, order_id, total):
        start = time.time()
        self.repository.save(order_id, total)
        log.info(f"saved order {order_id} in {time.time() - start:.2f}s")


class RetryingRepository:
    def __init__(self, repository, attempts=3):
        self.repository = repository
        self.attempts = attempts

    def save(self, order_id, total):
        for attempt in range(self.attempts):
            try:
                return self.repository.save(order_id, total)
            except ConnectionError:
                if attempt == self.attempts - 1:
                    raise
                time.sleep(0.5)


# compose in either order — neither class knows about the other
repository = LoggingRepository(RetryingRepository(PostgresRepository()))
processor = OrderProcessor(repository, notifier)
