# Day 9 — the dependency rule's dashed arrow, in actual code.
# PlaceOrder.run() looks clean on its own — it never imports psycopg2
# itself. The violation is hidden one call deeper, inside Order.save().


# order.py — Entities ring
import psycopg2


class Order:
    def __init__(self, id, items):
        self.id = id
        self.items = items
        self.status = "pending"

    def save(self):
        conn = psycopg2.connect("dbname=orders")
        cur = conn.cursor()
        cur.execute(
            "INSERT INTO orders (id, status) VALUES (%s, %s)",
            (self.id, self.status),
        )
        conn.commit()


# place_order.py — Use Cases ring
class PlaceOrder:
    def run(self, id, items):
        order = Order(id, items)
        # smell: looks harmless from here — the violation is one call away
        order.save()
        return order
