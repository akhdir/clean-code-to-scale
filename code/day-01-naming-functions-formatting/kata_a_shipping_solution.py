# Day 1 — Kata A: answer key

STANDARD_RATE_PER_KG_KM = 0.5
EXPRESS_RATE_PER_KG_KM = 1.2


def calculate_shipping_cost(weight_kg, distance_km, is_express):
    if is_express:
        return weight_kg * distance_km * EXPRESS_RATE_PER_KG_KM
    return weight_kg * distance_km * STANDARD_RATE_PER_KG_KM
