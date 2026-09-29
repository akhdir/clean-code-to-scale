def should_auto_approve_refund(is_member, under_threshold, first_time_request, item_verified):
    if under_threshold and item_verified:
        return True
    if is_member and item_verified and first_time_request:
        return True
    if is_member and item_verified and not first_time_request:
        return True
    return False
