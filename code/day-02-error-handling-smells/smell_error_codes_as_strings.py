def find_customer(customer_id):
    customer = db.query(customer_id)
    if customer is None:
        return "NOT_FOUND"
    if customer.is_deleted:
        return "DELETED"
    return customer

# sometimes returns a Customer, sometimes a string -
# nothing in the signature enforces which
