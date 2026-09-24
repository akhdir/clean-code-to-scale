# Day 3 — DIP kata: answer key.


class InventoryAlert:
    def __init__(self, threshold, notifier, clock):
        self.threshold = threshold
        self.notifier = notifier
        self.clock = clock

    def check(self, item):
        if item.stock < self.threshold:
            timestamp = self.clock.now()
            self.notifier.send(f"[{timestamp}] Low stock: {item.name} ({item.stock} left)")


# production
alert = InventoryAlert(10, SlackNotifier(webhook_url="..."), SystemClock())

# test — no network, no real clock
alert = InventoryAlert(10, FakeNotifier(), FixedClock("2026-01-01"))
