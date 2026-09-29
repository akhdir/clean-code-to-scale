def apply_refund(order):
    try:
        if order.status != "delivered":
            return {"error": "not eligible"}
        amount = order.total * 0.95  # restocking fee
        payment_gateway.refund(order.payment_id, amount)
        order.status = "refunded"
        db.save(order)
        send_refund_email(order.customer.email, amount)
    except Exception as e:
        log.error(e)
        return {"error": "failed"}

# eligibility, the fee rule, the refund call, and the
# notification are all one undifferentiated block
