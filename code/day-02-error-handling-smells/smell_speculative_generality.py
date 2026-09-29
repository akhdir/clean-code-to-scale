def send_notification(user, message, channel="email",
                       retry_strategy=None):
    if channel == "email":
        email_service.send(user.email, message)
    elif channel == "sms":
        sms_service.send(user.phone, message)
    # retry_strategy has never been passed

# added "in case we need custom retries someday" -
# no caller uses it, no one remembers why it's there
