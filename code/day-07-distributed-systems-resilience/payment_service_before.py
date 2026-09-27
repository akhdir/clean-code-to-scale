def confirm_payment(order):
    fraud_result = fraud_check_api.check(order)
    if fraud_result.flagged:
        return reject(order)
    return charge_payment(order)
