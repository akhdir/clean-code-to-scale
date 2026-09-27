fraud_breaker = CircuitBreaker(
    failure_threshold=5,
    reset_after_seconds=30,
)

def confirm_payment(order):
    try:
        fraud_result = fraud_breaker.call(
            lambda: fraud_check_api.check(order, timeout=2.0)
        )
    except (CircuitOpenError, TimeoutError):
        # the dependency is having a bad day — don't have one too
        return flag_for_manual_review(order)

    if fraud_result.flagged:
        return reject(order)
    return charge_payment(order)
