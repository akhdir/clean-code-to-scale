def should_flag_for_review(is_new_account, failed_attempts_high, unusual_location):
    return failed_attempts_high or (is_new_account and unusual_location)
