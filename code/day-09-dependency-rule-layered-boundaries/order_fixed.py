# Day 9 — fixed. PlaceOrder now receives a repository instead of
# asking Order to save itself.


# order.py — Entities ring, zero imports beyond the standard library
class Order:
    def __init__(self, id, items):
        self.id = id
        self.items = items
        self.status = "pending"


# order_repository.py — the port, owned by the Use Cases ring
class OrderRepository:
    def save(self, order):
        raise NotImplementedError


# place_order.py — Use Cases ring
class PlaceOrder:
    def __init__(self, repository):
        self.repository = repository

    def run(self, id, items):
        order = Order(id, items)
        self.repository.save(order)
        return order


# postgres_order_repository.py — Interface Adapters ring, implements the port
import psycopg2


class PostgresOrderRepository(OrderRepository):
    def save(self, order):
        conn = psycopg2.connect("dbname=orders")
        cur = conn.cursor()
        cur.execute(
            "INSERT INTO orders (id, status) VALUES (%s, %s)",
            (order.id, order.status),
        )
        conn.commit()
