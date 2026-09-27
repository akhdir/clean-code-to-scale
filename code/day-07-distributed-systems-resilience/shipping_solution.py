shipping_breaker = CircuitBreaker(failure_threshold=5, reset_after_seconds=30)
FLAT_RATE_FALLBACK = 7.99

def get_shipping_estimate(order):
    def attempt():
        return shipping_rate_api.get_rate(
            order.destination, order.weight, timeout=1.5
        )

    for attempt_number in range(2):  # one retry — a rate lookup is a safe GET
        try:
            return shipping_breaker.call(attempt)
        except CircuitOpenError:
            return FLAT_RATE_FALLBACK
        except TimeoutError:
            if attempt_number == 1:
                return FLAT_RATE_FALLBACK
            time.sleep(0.2)  # brief backoff before the one retry

def checkout(order):
    order.shipping_cost = get_shipping_estimate(order)
    return finalize_order(order)
