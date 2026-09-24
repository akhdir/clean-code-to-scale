# Day 3 — SRP kata: answer key (one valid shape).
# A third plan is a new entry in PRICES/VALID_PLANS, not a new elif
# in renew() — the OCP payoff of splitting pricing out on its own.


class SubscriptionValidator:
    VALID_PLANS = ("monthly", "annual")

    def validate(self, plan):
        if plan not in self.VALID_PLANS:
            raise ValueError("unknown plan")


class SubscriptionPricing:
    PRICES = {"monthly": 29.00, "annual": 290.00}

    def price_for(self, plan):
        return self.PRICES[plan]


# subscription.mark_renewed() moves the state mutation onto the entity itself
class SubscriptionManager:
    def __init__(self, payment_gateway, email_service, validator, pricing):
        self.payment_gateway = payment_gateway
        self.email_service = email_service
        self.validator = validator
        self.pricing = pricing

    def renew(self, subscription, plan):
        self.validator.validate(plan)
        price = self.pricing.price_for(plan)
        self.payment_gateway.charge(subscription.customer_id, price)
        subscription.mark_renewed(plan)
        self.email_service.send_renewal_notice(subscription, plan, price)
