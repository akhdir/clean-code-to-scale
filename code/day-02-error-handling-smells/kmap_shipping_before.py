def qualifies_for_free_shipping(is_member, total_over_50, is_promo_day):
    if is_member and total_over_50 and is_promo_day:
        return True
    if is_member and total_over_50 and not is_promo_day:
        return True
    if is_member and not total_over_50 and is_promo_day:
        return True
    if not is_member and total_over_50 and is_promo_day:
        return True
    return False
