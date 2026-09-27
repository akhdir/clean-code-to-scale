IS_HOLIDAY_SEASON = False

def calculate_shipping_cost(weight_kg):
    rate = 8.0
    if IS_HOLIDAY_SEASON:
        rate = 12.0
    return weight_kg * rate

# calculate_shipping_cost(5) returns a different answer in December
# than in June, with no argument telling you why
