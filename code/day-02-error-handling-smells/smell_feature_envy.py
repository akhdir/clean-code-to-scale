def format_invoice_header(invoice):
    name = invoice.customer.profile.legal_name
    addr = invoice.customer.profile.billing_address
    tax_id = invoice.customer.profile.tax_id
    return f"{name}\n{addr}\nTax ID: {tax_id}"

# barely touches invoice itself - it's three
# levels deep into customer.profile instead
