# Day 1 — Naming, Functions & Formatting
# Live demo, "before" — one function, five responsibilities.
# db, email, sms are stand-ins for real services.

TAX = 0.08


def process(o, u, p, send=True, n=False):
    total = 0
    for i in o["items"]:
        total += i["price"] * i["qty"]
    if u["loyalty"] > 1000:
        total = total * 0.9
    tax_amt = total * TAX
    total = total + tax_amt
    try:
        ok = p.charge(total)
    except Exception as e:
        ok = False
    if not ok:
        return {"status": "failed"}
    o["status"] = "paid"
    o["total"] = total
    db.save(o)
    if send:
        if n:
            sms.send(u["phone"], "Order paid: $%.2f" % total)
        else:
            email.send(u["email"], "Order paid: $%.2f" % total)
    return {"status": "ok", "total": total}
