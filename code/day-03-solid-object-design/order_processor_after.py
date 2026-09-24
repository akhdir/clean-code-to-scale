# Day 3 — Live demo "after": one job each, new discounts add a class.


class DiscountStrategy:
    def apply(self, total):
        raise NotImplementedError


class LoyaltyDiscount(DiscountStrategy):
    def apply(self, total):
        return total * 0.9


class SeasonalDiscount(DiscountStrategy):
    def apply(self, total):
        return total * 0.85


class EmployeeDiscount(DiscountStrategy):
    def apply(self, total):
        return total * 0.7


class NoDiscount(DiscountStrategy):
    def apply(self, total):
        return total


# repository and notifier are received, not constructed (DIP)
class OrderProcessor:
    def __init__(self, repository, notifier):
        self.repository = repository
        self.notifier = notifier

    def process(self, order, discount: DiscountStrategy):
        if not order.items:
            raise ValueError("empty order")
        total = sum(i.price * i.qty for i in order.items)
        total = discount.apply(total)
        self.repository.save(order.id, total)
        self.notifier.notify(order.customer_email, total)
        return total
