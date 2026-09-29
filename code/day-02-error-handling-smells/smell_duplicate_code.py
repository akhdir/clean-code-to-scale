def calculate_monthly_invoice(customer):
    total = sum(i.price * i.qty for i in customer.items)
    total = total * (1 - customer.discount_rate)
    total = total * 1.08  # tax
    return round(total, 2)

def calculate_annual_invoice(customer):
    total = sum(i.price * i.qty for i in customer.items)
    total = total * (1 - customer.discount_rate)
    total = total * 1.08  # tax
    return round(total * 12, 2)

# the pricing formula is copy-pasted - fix the tax
# rate in one function and the other keeps the old rate
