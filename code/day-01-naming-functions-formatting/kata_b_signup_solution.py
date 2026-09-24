# Day 1 — Kata B: answer key (one valid shape)

MIN_PASSWORD_LENGTH = 8


def is_valid_email(email):
    return "@" in email and "." in email


def is_strong_password(password):
    return len(password) >= MIN_PASSWORD_LENGTH and any(c.isdigit() for c in password)


def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()


def save_user(email, hashed_password):
    user_id = str(uuid.uuid4())
    users_db[user_id] = {"email": email, "password": hashed_password}
    return user_id


def send_welcome_email(email):
    email_service.send(email, "Welcome!")


def register_user(email, password):
    if not is_valid_email(email):
        return {"error": "invalid email"}
    if not is_strong_password(password):
        return {"error": "weak password"}
    user_id = save_user(email, hash_password(password))
    log.info("user registered: " + user_id)
    return {"id": user_id}


# caller decides whether to send the welcome email
result = register_user(email, password)
send_welcome_email(email)
