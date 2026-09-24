# Day 3 — Live demo "before": SRP, OCP, and DIP all violated in one class.


class OrderProcessor:
    def __init__(self):
        self.db = PostgresDatabase()
        self.email = SmtpEmailService()

    def process(self, order):
        if not order.items:
            raise ValueError("empty order")

        total = sum(i.price * i.qty for i in order.items)

        if order.discount_type == "loyalty":
            total *= 0.9
        elif order.discount_type == "seasonal":
            total *= 0.85
        elif order.discount_type == "employee":
            total *= 0.7

        self.db.save(order.id, total)
        self.email.send(order.customer_email, f"Your total: ${total:.2f}")
        return total
