def save_user(user, send_welcome_email):
    db.save(user)
    if send_welcome_email:
        email_service.send(user.email, "Welcome!")

# one function, two different jobs, chosen by a flag
save_user(new_user, True)
save_user(imported_user, False)
