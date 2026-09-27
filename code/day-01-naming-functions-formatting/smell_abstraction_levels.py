def process_order(order):
    # arithmetic
    total = sum(i.price * i.qty for i in order.items)

    # infrastructure
    conn = psycopg2.connect(DB_URL)
    conn.cursor().execute(
        "UPDATE orders SET total=%s WHERE id=%s",
        (total, order.id))
    conn.commit()

    # business action
    send_email(order.customer.email, "Order confirmed")

# arithmetic, raw SQL, and a business notification, all in one function
