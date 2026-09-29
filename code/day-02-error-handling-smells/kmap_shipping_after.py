def qualifies_for_free_shipping(is_member, total_over_50, is_promo_day):
    return sum([is_member, total_over_50, is_promo_day]) >= 2
