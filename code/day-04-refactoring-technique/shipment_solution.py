# Day 4 — Refactoring kata: answer key.
# Reference test, written before the session, against the original,
# unrefactored function:
#   describe_shipment(Shipment(carrier="ups", packages=[Package(weight_kg=10)],
#       method="standard")) == "shipment via ups: $5.50"
# After the refactor, the same shipment goes through
#   describe_shipment(shipment, UpsSurcharge()) and returns the same string.


class CarrierSurcharge:
    def apply(self, weight_fee):
        raise NotImplementedError


class UpsSurcharge(CarrierSurcharge):
    def apply(self, weight_fee):
        return weight_fee * 0.1


class FedexSurcharge(CarrierSurcharge):
    def apply(self, weight_fee):
        return weight_fee * 0.08


class DhlSurcharge(CarrierSurcharge):
    def apply(self, weight_fee):
        return weight_fee * 0.12


class NoSurcharge(CarrierSurcharge):
    def apply(self, weight_fee):
        return 0


def calculate_weight_fee(packages):
    return sum(pkg.weight_kg * 0.5 for pkg in packages)


def describe_shipment(shipment, surcharge):
    weight_fee = calculate_weight_fee(shipment.packages)
    total = weight_fee + surcharge.apply(weight_fee)
    prefix = "EXPRESS " if shipment.method == "express" else ""
    return f"{prefix}shipment via {shipment.carrier}: ${total:.2f}"
