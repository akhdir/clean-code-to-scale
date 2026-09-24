# Day 1 — Kata B: Extract responsibilities (core, ~20 min)
# Break this into single-purpose functions: validation, hashing,
# persistence, and notification should each be their own function.
# Remove skip_email by moving the decision to the caller.


def register_user(email, password, skip_email=False):
    if "@" not in email or "." not in email:
        return {"error": "invalid email"}
    if len(password) < 8:
        return {"error": "weak password"}
    has_digit = False
    for c in password:
        if c.isdigit():
            has_digit = True
    if not has_digit:
        return {"error": "weak password"}
    hashed = hashlib.sha256(password.encode()).hexdigest()
    user_id = str(uuid.uuid4())
    users_db[user_id] = {"email": email, "password": hashed}
    if not skip_email:
        email_service.send(email, "Welcome!")
    log.info("user registered: " + user_id)
    return {"id": user_id}
