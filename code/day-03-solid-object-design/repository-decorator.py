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
