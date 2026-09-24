# Day 4 — Refactoring kata (48 min), against a red-green timer.
# Before refactoring anything, write your own test that pins down
# describe_shipment's current behavior for at least one realistic
# shipment. Only once that test passes against the untouched function
# do you start refactoring — then apply Extract Method, Inline, and
# Replace Conditional with Polymorphism one at a time, running your
# test after every single edit.
#
# No test file handed out on purpose — write test_shipment.py yourself.
# Hint: describe_shipment(Shipment(carrier="ups",
#   packages=[Package(weight_kg=10)], method="standard")) should return
# something — run it once by hand and write down what it actually returns.


def _is_express(s):
    return s.method == "express"


def describe_shipment(shipment):
    weight_fee = 0
    for pkg in shipment.packages:
        weight_fee += pkg.weight_kg * 0.5

    if shipment.carrier == "ups":
        surcharge = weight_fee * 0.1
    elif shipment.carrier == "fedex":
        surcharge = weight_fee * 0.08
    elif shipment.carrier == "dhl":
        surcharge = weight_fee * 0.12
    else:
        surcharge = 0

    express = _is_express(shipment)
    total = weight_fee + surcharge
    return f"{'EXPRESS ' if express else ''}shipment via {shipment.carrier}: ${total:.2f}"
