# Day 10 — Adapter 1: production. From Day 9, unchanged.


class PostgresOrderRepository(OrderRepository):
    def save(self, order):
        conn = psycopg2.connect("dbname=orders")
        cur = conn.cursor()
        cur.execute(
            "INSERT INTO orders (id, status) VALUES (%s, %s)",
            (order.id, order.status),
        )
        conn.commit()
