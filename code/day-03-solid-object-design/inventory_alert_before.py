# Day 3 — DIP kata (35 min): invert two dependencies.
# InventoryAlert builds its own notifier and clock instead of
# receiving them. Change it so __init__ receives both as parameters.


class InventoryAlert:
    def __init__(self, threshold):
        self.threshold = threshold
        self.notifier = SlackNotifier(webhook_url="https://hooks.slack.com/T000/B000/xxxx")
        self.clock = SystemClock()

    def check(self, item):
        if item.stock < self.threshold:
            timestamp = self.clock.now()
            self.notifier.send(f"[{timestamp}] Low stock: {item.name} ({item.stock} left)")
