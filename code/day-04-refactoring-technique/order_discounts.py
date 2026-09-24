# Day 4 — the DiscountStrategy classes generate_invoice_summary imports,
# identical to Day 3's OrderProcessor demo. One class per branch the
# if/elif used to have.


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
