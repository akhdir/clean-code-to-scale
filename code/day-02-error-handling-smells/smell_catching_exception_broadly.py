def parse_shipping_weight(raw_value):
    try:
        return float(raw_value) * 2.20462
    except Exception:
        return 0.0

# a malformed input and a typo in this function's own
# code both produce the same silent 0.0
