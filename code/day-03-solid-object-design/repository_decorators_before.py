# Day 3 — Decorator kata (20 min): wrap it, don't edit it.
# PostgresRepository.save() below does three unrelated jobs: persist,
# time-and-log, and retry. Pull logging and retry out into decorators
# that wrap any object with a .save(order_id, total) method — so
# PostgresRepository goes back to doing exactly one thing.


class PostgresRepository:
    def save(self, order_id, total):
        start = time.time()
        for attempt in range(3):
            try:
                db.execute(
                    "INSERT INTO orders (id, total) VALUES (%s, %s)",
                    order_id, total,
                )
                log.info(f"saved order {order_id} in {time.time() - start:.2f}s")
                return
            except ConnectionError:
                if attempt == 2:
                    raise
                time.sleep(0.5)
