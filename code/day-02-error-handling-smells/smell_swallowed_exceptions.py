def update_customer_email(customer, new_email):
    try:
        customer.email = new_email
        db.save(customer)
    except:
        pass

# the save can fail for a dozen reasons - a bad
# connection, a constraint violation, a typo'd column -
# and every one disappears without a trace
