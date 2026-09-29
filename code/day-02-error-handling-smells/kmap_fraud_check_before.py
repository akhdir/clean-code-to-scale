def should_flag_for_review(is_new_account, failed_attempts_high, unusual_location):
    if is_new_account and unusual_location:
        return True
    if failed_attempts_high:
        return True
    if is_new_account and failed_attempts_high and unusual_location:
        return True
    return False
