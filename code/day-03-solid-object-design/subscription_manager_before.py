# Day 3 — SRP kata, stretch (~15 min): reduce the responsibilities.
# renew() validates, prices, charges, updates state, and emails —
# four reasons to change in one method.


class SubscriptionManager:
    def __init__(self, payment_gateway, email_service):
        self.payment_gateway = payment_gateway
        self.email_service = email_service

    def renew(self, subscription, plan):
        if plan not in ("monthly", "annual"):
            raise ValueError("unknown plan")

        price = 29.00 if plan == "monthly" else 290.00

        self.payment_gateway.charge(subscription.customer_id, price)

        subscription.plan = plan
        subscription.renewed_at = datetime.now()

        self.email_service.send(
            subscription.customer_email,
            f"Your {plan} subscription has been renewed for ${price:.2f}"
        )
