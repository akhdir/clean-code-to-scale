def should_auto_approve_refund(is_member, under_threshold, item_verified):
    return item_verified and (under_threshold or is_member)
